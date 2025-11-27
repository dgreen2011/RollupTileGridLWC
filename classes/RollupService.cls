public with sharing class RollupService {

    public class RollupResult {
        @AuraEnabled public String value;
        @AuraEnabled public Integer recordCount;
        @AuraEnabled public String errorMessage;
        // Formatting hints for the client
        @AuraEnabled public Boolean isCurrency;
        @AuraEnabled public Boolean isPercent;
        // Front-end label for the aggregate field (not the API name)
        @AuraEnabled public String fieldLabel;
    }

    // NOTE: No cacheable=true here so every refresh hits the DB.
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

        // Normalise values coming from picklists:
        // - "Project (sitetracker__Project__c)"          -> "sitetracker__Project__c"
        // - "Project -> Program (Proj__c.Program__c)"    -> "Program__c"
        // - "Project__c.Amount__c"                       -> "Amount__c"
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
            if (String.isBlank(aggregateType)) {
                aggregateType = 'SUM';
            } else {
                aggregateType = aggregateType.trim().toUpperCase();
            }

            // Only COUNT can safely omit an aggregate field; everything else needs one.
            Boolean requiresAggregateField = (aggregateType != 'COUNT');

            // Still validate the field if the admin supplied one even for COUNT.
            Boolean shouldValidateAggregateField =
                requiresAggregateField || !String.isBlank(aggregateFieldApiName);

            if (shouldValidateAggregateField) {
                if (String.isBlank(aggregateFieldApiName)) {
                    result.errorMessage =
                        'Configuration error: aggregate field is required for aggregation type ' +
                        aggregateType + '.';
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
            if (aggregateType == 'COUNT') {
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
            if (aggregateType == 'COUNT_DISTINCT') {
                // At this point aggregateFieldApiName has been validated.
                String soql =
                    'SELECT COUNT_DISTINCT(' + aggregateFieldApiName + ') aggValue,' +
                    '       COUNT(Id) recordCount ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause;

                AggregateResult ar = (AggregateResult)Database.query(soql);
                Object distinctCount = ar.get('aggValue');
                Object total = ar.get('recordCount');

                result.value =
                    (distinctCount == null) ? null : String.valueOf(distinctCount);
                result.recordCount =
                    (total == null) ? null : Integer.valueOf(String.valueOf(total));
                return result;
            }

            // =========================
            // CONCATENATE / CONCATENATE_DISTINCT
            // =========================
            if (aggregateType == 'CONCATENATE' ||
                aggregateType == 'CONCATENATE_DISTINCT') {

                // No aliasing here to avoid "only aggregate expressions use field aliasing"
                // errors if admins include GROUP BY fragments in filters.
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

                    if (aggregateType == 'CONCATENATE_DISTINCT') {
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
            if (aggregateType == 'FIRST' || aggregateType == 'LAST') {
                String direction = (aggregateType == 'FIRST') ? 'ASC' : 'DESC';

                // Again, avoid aliasing here so the query works even with GROUP BY clauses
                // in admin-supplied filters.
                String soql =
                    'SELECT ' + aggregateFieldApiName +
                    ' FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause +
                    ' ORDER BY ' + aggregateFieldApiName + ' ' + direction +
                    ' NULLS ' + ((aggregateType == 'FIRST') ? 'LAST' : 'FIRST') +
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
            // SUM / AVG / MIN / MAX
            // =========================
            String aggFunction;
            if (aggregateType == 'AVERAGE' || aggregateType ==
