public with sharing class RollupService {

    @TestVisible private static Integer maxConcatenateRows = 2000;

    public class RollupResult {
        @AuraEnabled public String  value;
        @AuraEnabled public Integer recordCount;
        @AuraEnabled public String  errorMessage;
        // Formatting hints for the client
        @AuraEnabled public Boolean isCurrency;
        @AuraEnabled public Boolean isPercent;
        @AuraEnabled public Boolean isDate;
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
        if (!isSupportedAggregationType(aggType)) {
            result.errorMessage =
                'Configuration error: aggregation type "' + aggType + '" is not supported.';
            return result;
        }

        String filterValidationError = validateFilterCondition(filterCondition);
        if (filterValidationError != null) {
            result.errorMessage = filterValidationError;
            return result;
        }
        String safeFilterCondition =
            String.isBlank(filterCondition) ? null : filterCondition.trim();

        String parentObjectApiName = parentId.getSObjectType().getDescribe().getName();

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
                String childAccessError =
                    validateObjectReadable(childDescribe, 'child object', childObjectApiName);
                if (childAccessError != null) {
                    result.errorMessage = childAccessError;
                    return result;
                }
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
                Schema.DescribeFieldResult childRelFieldDescribe =
                    childFields.get(relationshipFieldApiName).getDescribe();
                String childRelValidationError = validateRelationshipField(
                    childRelFieldDescribe,
                    String.isBlank(originalRelationshipField)
                        ? relationshipFieldApiName
                        : originalRelationshipField,
                    childObjectApiName,
                    parentObjectApiName
                );
                if (childRelValidationError != null) {
                    result.errorMessage = childRelValidationError;
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
                String grandchildAccessError =
                    validateObjectReadable(grandchildDescribe, 'grandchild object', grandchildObjectApiName);
                if (grandchildAccessError != null) {
                    result.errorMessage = grandchildAccessError;
                    return result;
                }
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

                if (!grandRelFieldDescribe.isAccessible()) {
                    String labelForMessage = String.isBlank(originalGrandchildRelationshipField)
                        ? grandchildRelationshipFieldApiName
                        : originalGrandchildRelationshipField;
                    result.errorMessage =
                        'Configuration error: relationship field "' + labelForMessage +
                        '" on ' + grandchildObjectApiName + ' is not readable for the current user.';
                    return result;
                }

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

                    Schema.DescribeFieldResult aggFieldDescribe =
                        grandchildFields.get(aggregateFieldApiName).getDescribe();
                    String aggregateValidationError = validateAggregateField(
                        aggFieldDescribe,
                        String.isBlank(originalAggregateField)
                            ? aggregateFieldApiName
                            : originalAggregateField,
                        aggType,
                        grandchildObjectApiName
                    );
                    if (aggregateValidationError != null) {
                        result.errorMessage = aggregateValidationError;
                        return result;
                    }
                    setFormattingHints(result, aggFieldDescribe);
                }

                // 6) WHERE clause:
                //    grandchildRelField IN (
                //        SELECT Id FROM Child WHERE childRelToParent = :parentId
                //    )
                String whereClause =
                    grandchildRelationshipFieldApiName +
                    ' IN (SELECT Id FROM ' + childObjectApiName +
                    ' WHERE ' + relationshipFieldApiName + ' = :parentId)';
                if (!String.isBlank(safeFilterCondition)) {
                    whereClause += ' AND (' + safeFilterCondition + ')';
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
                    Integer concatTotal = Database.countQuery(
                        'SELECT COUNT() FROM ' + grandchildObjectApiName +
                        ' WHERE ' + whereClause
                    );
                    result.recordCount = concatTotal;
                    if (concatTotal > maxConcatenateRows) {
                        result.errorMessage =
                            'Configuration error: CONCATENATE rollups are limited to ' +
                            maxConcatenateRows +
                            ' records. Add a filter or use a counted aggregation instead.';
                        return result;
                    }

                    String soql =
                        'SELECT ' + aggregateFieldApiName +
                        ' FROM ' + grandchildObjectApiName +
                        ' WHERE ' + whereClause +
                        ' ORDER BY ' + aggregateFieldApiName +
                        ' LIMIT ' + maxConcatenateRows;

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
                        if (raw == null) {
                            result.value = null;
                        } else if (result.isDate == true) {
                            result.value = formatDateForClient(raw);
                        } else {
                            result.value = String.valueOf(raw);
                        }
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

                    if (rawValue == null) {
                        result.value = null;
                    } else if (result.isDate == true &&
                               (aggType == 'MIN' || aggType == 'MAX')) {
                        result.value = formatDateForClient(rawValue);
                    } else {
                        result.value = String.valueOf(rawValue);
                    }

                    result.recordCount =
                        (total == null) ? null : Integer.valueOf(String.valueOf(total));
                }

                return result;
            }

            // ==========================================================
            // EXISTING CHILD-ONLY MODE (unchanged structure)
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
            String childAccessError2 =
                validateObjectReadable(childDescribe2, 'child object', childObjectApiName);
            if (childAccessError2 != null) {
                result.errorMessage = childAccessError2;
                return result;
            }
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
            Schema.DescribeFieldResult relationshipFieldDescribe2 =
                childFields2.get(relationshipFieldApiName).getDescribe();
            String relationshipValidationError2 = validateRelationshipField(
                relationshipFieldDescribe2,
                String.isBlank(originalRelationshipField)
                    ? relationshipFieldApiName
                    : originalRelationshipField,
                childObjectApiName,
                parentObjectApiName
            );
            if (relationshipValidationError2 != null) {
                result.errorMessage = relationshipValidationError2;
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

                Schema.DescribeFieldResult aggFieldDescribe2 =
                    childFields2.get(aggregateFieldApiName).getDescribe();
                String aggregateValidationError2 = validateAggregateField(
                    aggFieldDescribe2,
                    String.isBlank(originalAggregateField)
                        ? aggregateFieldApiName
                        : originalAggregateField,
                    aggType,
                    childObjectApiName
                );
                if (aggregateValidationError2 != null) {
                    result.errorMessage = aggregateValidationError2;
                    return result;
                }
                setFormattingHints(result, aggFieldDescribe2);
            }

            // ----- WHERE CLAUSE -----
            String whereClause2 = relationshipFieldApiName + ' = :parentId';
            if (!String.isBlank(safeFilterCondition)) {
                whereClause2 += ' AND (' + safeFilterCondition + ')';
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
                Integer concatTotal2 = Database.countQuery(
                    'SELECT COUNT() FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause2
                );
                result.recordCount = concatTotal2;
                if (concatTotal2 > maxConcatenateRows) {
                    result.errorMessage =
                        'Configuration error: CONCATENATE rollups are limited to ' +
                        maxConcatenateRows +
                        ' records. Add a filter or use a counted aggregation instead.';
                    return result;
                }

                String soql2 =
                    'SELECT ' + aggregateFieldApiName +
                    ' FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause2 +
                    ' ORDER BY ' + aggregateFieldApiName +
                    ' LIMIT ' + maxConcatenateRows;

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
                    if (raw2 == null) {
                        result.value = null;
                    } else if (result.isDate == true) {
                        result.value = formatDateForClient(raw2);
                    } else {
                        result.value = String.valueOf(raw2);
                    }
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

                if (rawValue2 == null) {
                    result.value = null;
                } else if (result.isDate == true &&
                           (aggType == 'MIN' || aggType == 'MAX')) {
                    result.value = formatDateForClient(rawValue2);
                } else {
                    result.value = String.valueOf(rawValue2);
                }

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

    private static Boolean isSupportedAggregationType(String aggType) {
        return aggType == 'SUM' ||
            aggType == 'AVERAGE' ||
            aggType == 'MAX' ||
            aggType == 'MIN' ||
            aggType == 'COUNT' ||
            aggType == 'COUNT_DISTINCT' ||
            aggType == 'CONCATENATE' ||
            aggType == 'CONCATENATE_DISTINCT' ||
            aggType == 'FIRST' ||
            aggType == 'LAST';
    }

    private static String validateObjectReadable(
        Schema.DescribeSObjectResult objectDescribe,
        String objectRole,
        String objectApiName
    ) {
        if (objectDescribe == null) {
            return 'Configuration error: ' + objectRole + ' "' + objectApiName + '" was not found.';
        }
        if (!objectDescribe.isAccessible() || !objectDescribe.isQueryable()) {
            return 'Configuration error: ' + objectRole + ' "' + objectApiName +
                '" is not readable for the current user.';
        }
        return null;
    }

    private static String validateRelationshipField(
        Schema.DescribeFieldResult fieldDescribe,
        String fieldLabelForMessage,
        String sourceObjectApiName,
        String targetObjectApiName
    ) {
        if (fieldDescribe == null) {
            return 'Configuration error: relationship field "' + fieldLabelForMessage +
                '" was not found on ' + sourceObjectApiName + '.';
        }
        if (!fieldDescribe.isAccessible()) {
            return 'Configuration error: relationship field "' + fieldLabelForMessage +
                '" on ' + sourceObjectApiName + ' is not readable for the current user.';
        }
        if (fieldDescribe.getType() != Schema.DisplayType.REFERENCE) {
            return 'Configuration error: relationship field "' + fieldLabelForMessage +
                '" on ' + sourceObjectApiName +
                ' must be a lookup field that points to ' + targetObjectApiName + '.';
        }

        Boolean pointsToTarget = false;
        for (Schema.SObjectType refType : fieldDescribe.getReferenceTo()) {
            if (refType != null && refType.getDescribe().getName() == targetObjectApiName) {
                pointsToTarget = true;
                break;
            }
        }
        if (!pointsToTarget) {
            return 'Configuration error: relationship field "' + fieldLabelForMessage +
                '" on ' + sourceObjectApiName +
                ' must look up to ' + targetObjectApiName + '.';
        }
        return null;
    }

    private static String validateAggregateField(
        Schema.DescribeFieldResult fieldDescribe,
        String fieldLabelForMessage,
        String aggType,
        String objectApiName
    ) {
        if (fieldDescribe == null) {
            return 'Configuration error: aggregate field "' + fieldLabelForMessage +
                '" was not found on ' + objectApiName + '.';
        }
        if (!fieldDescribe.isAccessible()) {
            return 'Configuration error: aggregate field "' + fieldLabelForMessage +
                '" on ' + objectApiName + ' is not readable for the current user.';
        }
        if (aggType == 'COUNT') {
            return null;
        }

        Schema.DisplayType t = fieldDescribe.getType();
        if ((aggType == 'SUM' || aggType == 'AVERAGE') && !isNumericFieldType(t)) {
            return 'Configuration error: aggregation type ' + aggType +
                ' requires a numeric aggregate field.';
        }
        if (aggType == 'COUNT_DISTINCT' && !fieldDescribe.isGroupable()) {
            return 'Configuration error: aggregate field "' + fieldLabelForMessage +
                '" cannot be grouped safely for aggregation type ' + aggType + '.';
        }

        return null;
    }

    private static Boolean isNumericFieldType(Schema.DisplayType t) {
        return t == Schema.DisplayType.Currency ||
            t == Schema.DisplayType.Double ||
            t == Schema.DisplayType.Integer ||
            t == Schema.DisplayType.Long ||
            t == Schema.DisplayType.Percent;
    }

    private static void setFormattingHints(
        RollupResult result,
        Schema.DescribeFieldResult fieldDescribe
    ) {
        if (result == null || fieldDescribe == null) {
            return;
        }

        Schema.DisplayType t = fieldDescribe.getType();
        result.isCurrency = (t == Schema.DisplayType.Currency);
        result.isPercent  = (t == Schema.DisplayType.Percent);
        result.isDate     =
            (t == Schema.DisplayType.Date ||
             t == Schema.DisplayType.Datetime);
        result.fieldLabel = fieldDescribe.getLabel();
    }

    private static String validateFilterCondition(String rawFilterCondition) {
        if (String.isBlank(rawFilterCondition)) {
            return null;
        }

        String trimmed = rawFilterCondition.trim();
        String quoteStateError = validateFilterQuotesAndParentheses(trimmed);
        if (quoteStateError != null) {
            return quoteStateError;
        }

        String stripped = stripQuotedLiterals(trimmed);
        String lowered = ' ' + stripped.toLowerCase().replaceAll('\\s+', ' ') + ' ';

        if (lowered.contains(';') ||
            lowered.contains('--') ||
            lowered.contains('/*') ||
            lowered.contains('*/')) {
            return 'Configuration error: additional SOQL filter contains unsupported comment or statement syntax.';
        }

        List<String> forbiddenClauses = new List<String>{
            ' select ',
            ' from ',
            ' where ',
            ' order by ',
            ' group by ',
            ' having ',
            ' limit ',
            ' offset ',
            ' all rows ',
            ' for update ',
            ' with ',
            ' typeof ',
            ' fields('
        };
        for (String forbiddenClause : forbiddenClauses) {
            if (lowered.contains(forbiddenClause)) {
                return 'Configuration error: additional SOQL filter contains unsupported clause "' +
                    forbiddenClause.trim() + '".';
            }
        }

        return null;
    }

    private static String validateFilterQuotesAndParentheses(String value) {
        Integer depth = 0;
        Boolean inQuote = false;
        String quote = '\'';
        String backslash = '\\';

        for (Integer i = 0; i < value.length(); i++) {
            String ch = value.substring(i, i + 1);
            if (inQuote && ch == backslash && i + 1 < value.length()) {
                i++;
                continue;
            }
            if (ch == quote) {
                inQuote = !inQuote;
                continue;
            }
            if (inQuote) {
                continue;
            }
            if (ch == '(') {
                depth++;
            } else if (ch == ')') {
                depth--;
                if (depth < 0) {
                    return 'Configuration error: additional SOQL filter has unbalanced parentheses.';
                }
            }
        }

        if (inQuote) {
            return 'Configuration error: additional SOQL filter has an unterminated string literal.';
        }
        if (depth != 0) {
            return 'Configuration error: additional SOQL filter has unbalanced parentheses.';
        }
        return null;
    }

    private static String stripQuotedLiterals(String value) {
        String output = '';
        Boolean inQuote = false;
        String quote = '\'';
        String backslash = '\\';

        for (Integer i = 0; i < value.length(); i++) {
            String ch = value.substring(i, i + 1);
            if (inQuote && ch == backslash && i + 1 < value.length()) {
                output += '  ';
                i++;
                continue;
            }
            if (ch == quote) {
                inQuote = !inQuote;
                output += ' ';
            } else if (inQuote) {
                output += ' ';
            } else {
                output += ch;
            }
        }

        return output;
    }

    /**
     * Format a Date/Datetime as MM/DD/YYYY for the client.
     * Avoids Date.format(String), which is not supported.
     */
    private static String formatDateForClient(Object raw) {
        if (raw == null) {
            return null;
        }

        Date d;
        if (raw instanceof Date) {
            d = (Date) raw;
        } else if (raw instanceof Datetime) {
            d = ((Datetime) raw).date();
        } else {
            // Not a date-like value; just stringify it
            return String.valueOf(raw);
        }

        Integer year  = d.year();
        Integer month = d.month();
        Integer day   = d.day();

        // Zero-pad month/day so we always get MM/DD/YYYY
        String mm = (month < 10 ? '0' : '') + String.valueOf(month);
        String dd = (day   < 10 ? '0' : '') + String.valueOf(day);

        return mm + '/' + dd + '/' + String.valueOf(year);
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
