public with sharing class RollupService {

    public class RollupResult {
        @AuraEnabled public String  value;
        @AuraEnabled public Integer recordCount;
        @AuraEnabled public String  errorMessage;
        // Formatting hints for the client
        @AuraEnabled public Boolean isCurrency;
        @AuraEnabled public Boolean isPercent;
        // Front-end label for the aggregate field (not the API name)
        @AuraEnabled public String  fieldLabel;
    }

    @AuraEnabled
    public static RollupResult getRollup(
        Id parentId,
        String childObjectApiName,
        String relationshipFieldApiName,
        String aggregateFieldApiName,
        String aggregateType,
        String filterCondition,
        String grandchildObjectApiName,
        String grandchildRelationshipFieldApiName
    ) {
        RollupResult result = new RollupResult();

        // Basic guard rails
        if (parentId == null) {
            result.errorMessage = 'Configuration error: parentId was not supplied.';
            return result;
        }

        // Keep originals for nicer error messages
        String originalChildObject                = childObjectApiName;
        String originalRelationshipField          = relationshipFieldApiName;
        String originalAggregateField             = aggregateFieldApiName;
        String originalGrandchildObject           = grandchildObjectApiName;
        String originalGrandchildRelationshipField = grandchildRelationshipFieldApiName;

        // Normalise values coming from picklists (child + aggregate)
        childObjectApiName       = normalizeApiName(childObjectApiName);
        relationshipFieldApiName = normalizeApiName(relationshipFieldApiName);
        aggregateFieldApiName    = normalizeApiName(aggregateFieldApiName);

        // Normalise values for the optional grandchild configuration
        grandchildObjectApiName =
            normalizeApiName(grandchildObjectApiName);
        grandchildRelationshipFieldApiName =
            normalizeApiName(grandchildRelationshipFieldApiName);

        if (String.isBlank(childObjectApiName)) {
            result.errorMessage = 'Configuration error: child object is blank.';
            return result;
        }
        if (String.isBlank(relationshipFieldApiName)) {
            result.errorMessage = 'Configuration error: relationship field is blank.';
            return result;
        }

        // Normalise aggregation type
        String aggType;
        if (String.isBlank(aggregateType)) {
            aggType = 'SUM';
        } else {
            aggType = aggregateType.trim().toUpperCase();
        }
        if (aggType == 'AVG') {
            aggType = 'AVERAGE';
        }

        try {
            Map<String, Schema.SObjectType> globalDescribe = Schema.getGlobalDescribe();

            // ==========================================================
            // GRANDCHILD MODE (Parent -> Child -> Grandchild)
            // ==========================================================
            Boolean hasGrandchildConfig =
                !String.isBlank(grandchildObjectApiName) ||
                !String.isBlank(grandchildRelationshipFieldApiName);

            if (hasGrandchildConfig) {
                // 1) Validate child object (pivot between parent and grandchild)
                if (!globalDescribe.containsKey(childObjectApiName)) {
                    String labelForMessage = String.isBlank(originalChildObject)
                        ? childObjectApiName
                        : originalChildObject;
                    result.errorMessage =
                        'Configuration error: child object "' + labelForMessage + '" was not found.';
                    return result;
                }

                Schema.SObjectType childType = globalDescribe.get(childObjectApiName);
                Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
                Map<String, Schema.SObjectField> childFields = childDescribe.fields.getMap();

                // 2) Validate relationship field on child (child -> parent)
                if (!childFields.containsKey(relationshipFieldApiName)) {
                    String labelForMessage = String.isBlank(originalRelationshipField)
                        ? relationshipFieldApiName
                        : originalRelationshipField;
                    result.errorMessage =
                        'Configuration error: relationship field "' + labelForMessage +
                        '" was not found on ' + childObjectApiName + '.';
                    return result;
                }

                // 3) Validate grandchild object
                if (String.isBlank(grandchildObjectApiName)) {
                    String labelForMessage = String.isBlank(originalGrandchildObject)
                        ? 'Grandchild Object'
                        : originalGrandchildObject;
                    result.errorMessage =
                        'Configuration error: grandchild object "' + labelForMessage +
                        '" is blank or invalid.';
                    return result;
                }
                if (String.isBlank(grandchildRelationshipFieldApiName)) {
                    result.errorMessage =
                        'Configuration error: relationship field (lookup on grandchild) is blank.';
                    return result;
                }

                if (!globalDescribe.containsKey(grandchildObjectApiName)) {
                    String labelForMessage = String.isBlank(originalGrandchildObject)
                        ? grandchildObjectApiName
                        : originalGrandchildObject;
                    result.errorMessage =
                        'Configuration error: grandchild object "' + labelForMessage + '" was not found.';
                    return result;
                }

                Schema.SObjectType grandchildType = globalDescribe.get(grandchildObjectApiName);
                Schema.DescribeSObjectResult grandchildDescribe = grandchildType.getDescribe();
                Map<String, Schema.SObjectField> grandchildFields =
                    grandchildDescribe.fields.getMap();

                // 4) Validate relationship field on grandchild (grandchild -> child)
                if (!grandchildFields.containsKey(grandchildRelationshipFieldApiName)) {
                    String labelForMessage = String.isBlank(originalGrandchildRelationshipField)
                        ? grandchildRelationshipFieldApiName
                        : originalGrandchildRelationshipField;
                    result.errorMessage =
                        'Configuration error: relationship field "' + labelForMessage +
                        '" was not found on ' + grandchildObjectApiName + '.';
                    return result;
                }

                Schema.DescribeFieldResult grandRelFieldDescribe =
                    grandchildFields.get(grandchildRelationshipFieldApiName).getDescribe();

                if (grandRelFieldDescribe.getType() != Schema.DisplayType.REFERENCE) {
                    String labelForMessage = String.isBlank(originalGrandchildRelationshipField)
                        ? grandchildRelationshipFieldApiName
                        : originalGrandchildRelationshipField;
                    result.errorMessage =
                        'Configuration error: relationship field "' + labelForMessage +
                        '" on ' + grandchildObjectApiName +
                        ' must be a lookup field that points to ' + childObjectApiName + '.';
                    return result;
                }

                Boolean pointsToChild = false;
                for (Schema.SObjectType refType : grandRelFieldDescribe.getReferenceTo()) {
                    if (refType != null &&
                        refType.getDescribe().getName() == childObjectApiName) {
                        pointsToChild = true;
                        break;
                    }
                }
                if (!pointsToChild) {
                    String labelForMessage = String.isBlank(originalGrandchildRelationshipField)
                        ? grandchildRelationshipFieldApiName
                        : originalGrandchildRelationshipField;
                    result.errorMessage =
                        'Configuration error: relationship field "' + labelForMessage +
                        '" on ' + grandchildObjectApiName +
                        ' must look up to ' + childObjectApiName + '.';
                    return result;
                }

                // 5) Aggregate field validation (on the GRANDCHILD object)
                Boolean requiresAggregateField = (aggType != 'COUNT');
                Boolean shouldValidateAggregateField =
                    requiresAggregateField || !String.isBlank(aggregateFieldApiName);

                if (shouldValidateAggregateField) {
                    if (String.isBlank(aggregateFieldApiName)) {
                        result.errorMessage =
                            'Configuration error: aggregate field is required for aggregation type ' +
                            aggType + '.';
                        return result;
                    }
                    if (!grandchildFields.containsKey(aggregateFieldApiName)) {
                        String labelForMessage = String.isBlank(originalAggregateField)
                            ? aggregateFieldApiName
                            : originalAggregateField;
                        result.errorMessage =
                            'Configuration error: aggregate field "' + labelForMessage +
                            '" was not found on ' + grandchildObjectApiName + '.';
                        return result;
                    }

                    // Capture basic type information so the client can format values nicely.
                    try {
                        Schema.DescribeFieldResult aggFieldDescribe =
                            grandchildFields.get(aggregateFieldApiName).getDescribe();
                        if (aggFieldDescribe != null) {
                            Schema.DisplayType t = aggFieldDescribe.getType();
                            result.isCurrency = (t == Schema.DisplayType.Currency);
                            result.isPercent  = (t == Schema.DisplayType.Percent);
                            result.fieldLabel = aggFieldDescribe.getLabel();
                        }
                    } catch (Exception ignoreTypeEx) {
                        // Never let describe/formatting hints break the rollup itself.
                    }
                }

                // 6) WHERE clause:
                //    grandchildRelField IN (
                //        SELECT Id FROM Child WHERE childRelToParent = :parentId
                //    )
                String whereClause =
                    grandchildRelationshipFieldApiName +
                    ' IN (SELECT Id FROM ' + childObjectApiName +
                    ' WHERE ' + relationshipFieldApiName + ' = :parentId)';
                if (!String.isBlank(filterCondition)) {
                    whereClause += ' AND (' + filterCondition + ')';
                }

                // =========================
                // COUNT
                // =========================
                if (aggType == 'COUNT') {
                    String countSoql =
                        'SELECT COUNT() ' +
                        'FROM ' + grandchildObjectApiName +
                        ' WHERE ' + whereClause;

                    Integer total = Database.countQuery(countSoql);
                    result.recordCount = total;
                    result.value = (total == null) ? null : String.valueOf(total);
                    return result;
                }

                // =========================
                // COUNT DISTINCT
                // =========================
                if (aggType == 'COUNT_DISTINCT') {
                    String soql =
                        'SELECT COUNT_DISTINCT(' + aggregateFieldApiName + ') aggValue,' +
                        '       COUNT(Id) recordCount ' +
                        'FROM ' + grandchildObjectApiName +
                        ' WHERE ' + whereClause;

                    List<AggregateResult> ars = Database.query(soql);
                    if (!ars.isEmpty()) {
                        AggregateResult ar = ars[0];

                        Object distinctCount = ar.get('aggValue');
                        Object total = ar.get('recordCount');

                        result.value =
                            (distinctCount == null) ? null : String.valueOf(distinctCount);
                        result.recordCount =
                            (total == null) ? null : Integer.valueOf(String.valueOf(total));
                    }
                    return result;
                }

                // =========================
                // CONCATENATE / CONCATENATE_DISTINCT
                // =========================
                if (aggType == 'CONCATENATE' || aggType == 'CONCATENATE_DISTINCT') {

                    String soql =
                        'SELECT ' + aggregateFieldApiName +
                        ' FROM ' + grandchildObjectApiName +
                        ' WHERE ' + whereClause +
                        ' ORDER BY ' + aggregateFieldApiName;

                    List<SObject> rows = Database.query(soql);
                    List<String> values = new List<String>();
                    Set<String> seen = new Set<String>();

                    for (SObject row : rows) {
                        Object raw = row.get(aggregateFieldApiName);
                        if (raw == null) continue;

                        String strVal = String.valueOf(raw).trim();
                        if (strVal == '') continue;

                        if (aggType == 'CONCATENATE_DISTINCT') {
                            if (!seen.contains(strVal)) {
                                seen.add(strVal);
                                values.add(strVal);
                            }
                        } else {
                            values.add(strVal);
                        }
                    }

                    result.recordCount = rows.size();
                    result.value = values.isEmpty()
                        ? null
                        : String.join(values, '; ');
                    return result;
                }

                // =========================
                // FIRST / LAST
                // =========================
                if (aggType == 'FIRST' || aggType == 'LAST') {
                    String direction = (aggType == 'FIRST') ? 'ASC' : 'DESC';

                    String soql =
                        'SELECT ' + aggregateFieldApiName +
                        ' FROM ' + grandchildObjectApiName +
                        ' WHERE ' + whereClause +
                        ' ORDER BY ' + aggregateFieldApiName + ' ' + direction +
                        ' NULLS ' + ((aggType == 'FIRST') ? 'LAST' : 'FIRST') +
                        ' LIMIT 1';

                    List<SObject> rows = Database.query(soql);
                    if (!rows.isEmpty()) {
                        Object raw = rows[0].get(aggregateFieldApiName);
                        result.value = (raw == null) ? null : String.valueOf(raw);
                    }

                    String countSoql =
                        'SELECT COUNT() ' +
                        'FROM ' + grandchildObjectApiName +
                        ' WHERE ' + whereClause;

                    result.recordCount = Database.countQuery(countSoql);
                    return result;
                }

                // =========================
                // SUM / AVERAGE / MIN / MAX
                // =========================
                String aggFunction;
                if (aggType == 'AVERAGE') {
                    aggFunction = 'AVG';
                } else if (aggType == 'MAX') {
                    aggFunction = 'MAX';
                } else if (aggType == 'MIN') {
                    aggFunction = 'MIN';
                } else {
                    aggFunction = 'SUM';
                }

                String soql =
                    'SELECT ' + aggFunction + '(' + aggregateFieldApiName + ') aggValue,' +
                    '       COUNT(Id) recordCount ' +
                    'FROM ' + grandchildObjectApiName +
                    ' WHERE ' + whereClause;

                List<AggregateResult> ars = Database.query(soql);
                if (!ars.isEmpty()) {
                    AggregateResult ar = ars[0];

                    Object rawValue = ar.get('aggValue');
                    Object total = ar.get('recordCount');

                    result.value =
                        (rawValue == null) ? null : String.valueOf(rawValue);
                    result.recordCount =
                        (total == null) ? null : Integer.valueOf(String.valueOf(total));
                }

                return result;
            }

            // ==========================================================
            // EXISTING CHILD-ONLY MODE (unchanged)
            // ==========================================================

            // ----- CHILD OBJECT CHECK -----
            if (!globalDescribe.containsKey(childObjectApiName)) {
                String labelForMessage = String.isBlank(originalChildObject)
                    ? childObjectApiName
                    : originalChildObject;
                result.errorMessage =
                    'Configuration error: child object "' + labelForMessage + '" was not found.';
                return result;
            }

            Schema.SObjectType childType2 = globalDescribe.get(childObjectApiName);
            Schema.DescribeSObjectResult childDescribe2 = childType2.getDescribe();
            Map<String, Schema.SObjectField> childFields2 = childDescribe2.fields.getMap();

            // ----- RELATIONSHIP FIELD CHECK -----
            if (!childFields2.containsKey(relationshipFieldApiName)) {
                String labelForMessage = String.isBlank(originalRelationshipField)
                    ? relationshipFieldApiName
                    : originalRelationshipField;
                result.errorMessage =
                    'Configuration error: relationship field "' + labelForMessage +
                    '" was not found on ' + childObjectApiName + '.';
                return result;
            }

            // ----- AGGREGATE FIELD CHECK -----
            Boolean requiresAggregateField2 = (aggType != 'COUNT');
            Boolean shouldValidateAggregateField2 =
                requiresAggregateField2 || !String.isBlank(aggregateFieldApiName);

            if (shouldValidateAggregateField2) {
                if (String.isBlank(aggregateFieldApiName)) {
                    result.errorMessage =
                        'Configuration error: aggregate field is required for aggregation type ' +
                        aggType + '.';
                    return result;
                }
                if (!childFields2.containsKey(aggregateFieldApiName)) {
                    String labelForMessage = String.isBlank(originalAggregateField)
                        ? aggregateFieldApiName
                        : originalAggregateField;
                    result.errorMessage =
                        'Configuration error: aggregate field "' + labelForMessage +
                        '" was not found on ' + childObjectApiName + '.';
                    return result;
                }

                // Capture basic type information so the client can format values nicely.
                try {
                    Schema.DescribeFieldResult aggFieldDescribe2 =
                        childFields2.get(aggregateFieldApiName).getDescribe();
                    if (aggFieldDescribe2 != null) {
                        Schema.DisplayType t2 = aggFieldDescribe2.getType();
                        result.isCurrency = (t2 == Schema.DisplayType.Currency);
                        result.isPercent  = (t2 == Schema.DisplayType.Percent);
                        result.fieldLabel = aggFieldDescribe2.getLabel();
                    }
                } catch (Exception ignoreTypeEx2) {
                    // Never let describe/formatting hints break the rollup itself.
                }
            }

            // ----- WHERE CLAUSE -----
            String whereClause2 = relationshipFieldApiName + ' = :parentId';
            if (!String.isBlank(filterCondition)) {
                whereClause2 += ' AND (' + filterCondition + ')';
            }

            // =========================
            // COUNT
            // =========================
            if (aggType == 'COUNT') {
                String countSoql2 =
                    'SELECT COUNT() ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause2;

                Integer total2 = Database.countQuery(countSoql2);
                result.recordCount = total2;
                result.value = (total2 == null) ? null : String.valueOf(total2);
                return result;
            }

            // =========================
            // COUNT DISTINCT
            // =========================
            if (aggType == 'COUNT_DISTINCT') {
                String soql2 =
                    'SELECT COUNT_DISTINCT(' + aggregateFieldApiName + ') aggValue,' +
                    '       COUNT(Id) recordCount ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause2;

                List<AggregateResult> ars2 = Database.query(soql2);
                if (!ars2.isEmpty()) {
                    AggregateResult ar2 = ars2[0];

                    Object distinctCount2 = ar2.get('aggValue');
                    Object total2 = ar2.get('recordCount');

                    result.value =
                        (distinctCount2 == null) ? null : String.valueOf(distinctCount2);
                    result.recordCount =
                        (total2 == null) ? null : Integer.valueOf(String.valueOf(total2));
                }
                return result;
            }

            // =========================
            // CONCATENATE / CONCATENATE_DISTINCT
            // =========================
            if (aggType == 'CONCATENATE' || aggType == 'CONCATENATE_DISTINCT') {

                String soql2 =
                    'SELECT ' + aggregateFieldApiName +
                    ' FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause2 +
                    ' ORDER BY ' + aggregateFieldApiName;

                List<SObject> rows2 = Database.query(soql2);
                List<String> values2 = new List<String>();
                Set<String> seen2 = new Set<String>();

                for (SObject row2 : rows2) {
                    Object raw2 = row2.get(aggregateFieldApiName);
                    if (raw2 == null) continue;

                    String strVal2 = String.valueOf(raw2).trim();
                    if (strVal2 == '') continue;

                    if (aggType == 'CONCATENATE_DISTINCT') {
                        if (!seen2.contains(strVal2)) {
                            seen2.add(strVal2);
                            values2.add(strVal2);
                        }
                    } else {
                        values2.add(strVal2);
                    }
                }

                result.recordCount = rows2.size();
                result.value = values2.isEmpty()
                    ? null
                    : String.join(values2, '; ');
                return result;
            }

            // =========================
            // FIRST / LAST
            // =========================
            if (aggType == 'FIRST' || aggType == 'LAST') {
                String direction2 = (aggType == 'FIRST') ? 'ASC' : 'DESC';

                String soql2 =
                    'SELECT ' + aggregateFieldApiName +
                    ' FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause2 +
                    ' ORDER BY ' + aggregateFieldApiName + ' ' + direction2 +
                    ' NULLS ' + ((aggType == 'FIRST') ? 'LAST' : 'FIRST') +
                    ' LIMIT 1';

                List<SObject> rows2 = Database.query(soql2);
                if (!rows2.isEmpty()) {
                    Object raw2 = rows2[0].get(aggregateFieldApiName);
                    result.value = (raw2 == null) ? null : String.valueOf(raw2);
                }

                String countSoql2 =
                    'SELECT COUNT() ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause2;

                result.recordCount = Database.countQuery(countSoql2);
                return result;
            }

            // =========================
            // SUM / AVERAGE / MIN / MAX
            // =========================
            String aggFunction2;
            if (aggType == 'AVERAGE') {
                aggFunction2 = 'AVG';
            } else if (aggType == 'MAX') {
                aggFunction2 = 'MAX';
            } else if (aggType == 'MIN') {
                aggFunction2 = 'MIN';
            } else {
                aggFunction2 = 'SUM';
            }

            String soql3 =
                'SELECT ' + aggFunction2 + '(' + aggregateFieldApiName + ') aggValue,' +
                '       COUNT(Id) recordCount ' +
                'FROM ' + childObjectApiName +
                ' WHERE ' + whereClause2;

            List<AggregateResult> ars3 = Database.query(soql3);
            if (!ars3.isEmpty()) {
                AggregateResult ar3 = ars3[0];

                Object rawValue2 = ar3.get('aggValue');
                Object total3 = ar3.get('recordCount');

                result.value =
                    (rawValue2 == null) ? null : String.valueOf(rawValue2);
                result.recordCount =
                    (total3 == null) ? null : Integer.valueOf(String.valueOf(total3));
            }

        } catch (Exception ex) {
            System.debug('RollupService.getRollup error: ' + ex);
            result.errorMessage =
                'Unexpected error while calculating rollup: ' + ex.getMessage();
        }

        return result;
    }

    /**
     * Normalises values coming from picklists that return things like:
     *  - "Project (sitetracker__Project__c)"
     *  - "Child • Parent Lookup (Child__c.Parent__c)"
     *  - "sitetracker__Project__c.Amount__c"
     *
     * After normalisation we want just the field/object API name piece:
     *  - "sitetracker__Project__c"
     *  - "Parent__c"
     *  - "Amount__c"
     */
    private static String normalizeApiName(String raw) {
        if (String.isBlank(raw)) {
            return null;
        }

        String trimmed = raw.trim();

        // If there is a "(...)" suffix, keep only the content inside.
        Integer open = trimmed.lastIndexOf('(');
        Integer close = trimmed.lastIndexOf(')');
        if (open != -1 && close != -1 && close > open) {
            trimmed = trimmed.substring(open + 1, close).trim();
        }

        // If we now have something like "Object__c.Field__c", keep only the field part.
        Integer dot = trimmed.lastIndexOf('.');
        if (dot != -1 && dot < trimmed.length() - 1) {
            trimmed = trimmed.substring(dot + 1).trim();
        }

        return trimmed;
    }
}
