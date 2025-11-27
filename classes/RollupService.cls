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
        String filterCondition
    ) {
        RollupResult result = new RollupResult();

        // Basic guard rails
        if (parentId == null) {
            result.errorMessage = 'Configuration error: parentId was not supplied.';
            return result;
        }

        // Keep originals for nicer error messages
        String originalChildObject       = childObjectApiName;
        String originalRelationshipField = relationshipFieldApiName;
        String originalAggregateField    = aggregateFieldApiName;

        // Normalise values coming from picklists
        childObjectApiName       = normalizeApiName(childObjectApiName);
        relationshipFieldApiName = normalizeApiName(relationshipFieldApiName);
        aggregateFieldApiName    = normalizeApiName(aggregateFieldApiName);

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

            // ----- CHILD OBJECT CHECK -----
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

            // ----- RELATIONSHIP FIELD CHECK -----
            if (!childFields.containsKey(relationshipFieldApiName)) {
                String labelForMessage = String.isBlank(originalRelationshipField)
                    ? relationshipFieldApiName
                    : originalRelationshipField;
                result.errorMessage =
                    'Configuration error: relationship field "' + labelForMessage +
                    '" was not found on ' + childObjectApiName + '.';
                return result;
            }

            // ----- AGGREGATE FIELD CHECK -----
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
                if (!childFields.containsKey(aggregateFieldApiName)) {
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
                    Schema.DescribeFieldResult aggFieldDescribe =
                        childFields.get(aggregateFieldApiName).getDescribe();
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

            // ----- WHERE CLAUSE -----
            String whereClause = relationshipFieldApiName + ' = :parentId';
            if (!String.isBlank(filterCondition)) {
                whereClause += ' AND (' + filterCondition + ')';
            }

            // =========================
            // COUNT
            // =========================
            if (aggType == 'COUNT') {
                String countSoql =
                    'SELECT COUNT() total ' +
                    'FROM ' + childObjectApiName +
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
                    'FROM ' + childObjectApiName +
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
                    ' FROM ' + childObjectApiName +
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
                    ' FROM ' + childObjectApiName +
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
                    'SELECT COUNT() total ' +
                    'FROM ' + childObjectApiName +
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
                'FROM ' + childObjectApiName +
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
