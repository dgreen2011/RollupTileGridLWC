public with sharing class RollupService {

    public class RollupResponse {
        @AuraEnabled public Object  value;
        @AuraEnabled public Integer recordCount;
        @AuraEnabled public Boolean isCurrency;
        @AuraEnabled public Boolean isPercent;
        @AuraEnabled public String  fieldLabel;
        @AuraEnabled public String  errorMessage;
    }

    @AuraEnabled(cacheable=true)
    public static RollupResponse getRollup(
        Id parentId,
        String childObjectApiName,
        String relationshipFieldApiName,
        String aggregateFieldApiName,
        String aggregateType,
        String filterCondition
    ) {
        RollupResponse res = new RollupResponse();

        try {
            // Basic validation
            if (parentId == null) {
                res.errorMessage = 'Parent Id is required.';
                return res;
            }
            if (String.isBlank(childObjectApiName)) {
                res.errorMessage = 'Child object API name is required.';
                return res;
            }
            if (String.isBlank(relationshipFieldApiName)) {
                res.errorMessage = 'Relationship field API name is required.';
                return res;
            }

            String normalizedAgg = normalizeAggregateType(aggregateType);

            // Describe child object
            Schema.SObjectType childType = Schema.getGlobalDescribe().get(childObjectApiName);
            if (childType == null) {
                res.errorMessage = 'Unknown child object: ' + childObjectApiName;
                return res;
            }

            Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
            Map<String, Schema.SObjectField> fieldMap = childDescribe.fields.getMap();

            Schema.DescribeFieldResult fieldDescribe;
            if (!String.isBlank(aggregateFieldApiName)) {
                if (!fieldMap.containsKey(aggregateFieldApiName)) {
                    res.errorMessage = 'Unknown aggregate field: ' + aggregateFieldApiName;
                    return res;
                }
                fieldDescribe = fieldMap.get(aggregateFieldApiName).getDescribe();
                res.fieldLabel = fieldDescribe.getLabel();
            }

            // Currency / percent flags (only for numeric-type aggregations)
            if (fieldDescribe != null &&
                normalizedAgg != 'COUNT' &&
                normalizedAgg != 'COUNT_DISTINCT' &&
                normalizedAgg != 'CONCATENATE' &&
                normalizedAgg != 'CONCATENATE_DISTINCT') {

                Schema.DisplayType dt = fieldDescribe.getType();
                if (dt == Schema.DisplayType.Currency) {
                    res.isCurrency = true;
                } else if (dt == Schema.DisplayType.Percent) {
                    res.isPercent = true;
                }
            }

            // WHERE clause
            String whereClause = relationshipFieldApiName + ' = \'' +
                String.escapeSingleQuotes(String.valueOf(parentId)) + '\'';

            if (!String.isBlank(filterCondition)) {
                whereClause += ' AND (' + filterCondition + ')';
            }

            // CONCATENATE / CONCATENATE_DISTINCT handled separately
            if (normalizedAgg == 'CONCATENATE' || normalizedAgg == 'CONCATENATE_DISTINCT') {
                if (String.isBlank(aggregateFieldApiName)) {
                    res.errorMessage = 'Aggregate field is required for concatenate rollups.';
                    return res;
                }
                if (fieldDescribe == null) {
                    res.errorMessage = 'Unable to describe aggregate field: ' + aggregateFieldApiName;
                    return res;
                }

                Schema.DisplayType dt = fieldDescribe.getType();
                // Only allow text-ish fields for concatenation
                if (dt != Schema.DisplayType.String &&
                    dt != Schema.DisplayType.TextArea &&
                    dt != Schema.DisplayType.Phone &&
                    dt != Schema.DisplayType.Email &&
                    dt != Schema.DisplayType.Url &&
                    dt != Schema.DisplayType.Picklist &&
                    dt != Schema.DisplayType.MultiPicklist) {

                    res.errorMessage = 'Concatenate rollups are only supported on text-like fields.';
                    return res;
                }

                String soql =
                    'SELECT ' + aggregateFieldApiName +
                    ' FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause;

                List<SObject> rows = Database.query(soql);
                res.recordCount = rows.size();

                Set<String> distinctSet = new Set<String>();
                List<String> values = new List<String>();

                for (SObject row : rows) {
                    Object raw = row.get(aggregateFieldApiName);
                    if (raw == null) continue;
                    String val = String.valueOf(raw);

                    if (normalizedAgg == 'CONCATENATE_DISTINCT') {
                        if (!distinctSet.contains(val)) {
                            distinctSet.add(val);
                            values.add(val);
                        }
                    } else {
                        values.add(val);
                    }
                }

                // Join with a simple delimiter; adjust if you want commas instead.
                res.value = String.join(values, '; ');
                return res;
            }

            // Numeric / counting aggregations (SUM, AVERAGE, MAX, MIN, COUNT, COUNT_DISTINCT)
            String fieldForFunction = aggregateFieldApiName;
            if (String.isBlank(fieldForFunction)) {
                // For COUNT / COUNT_DISTINCT with no field, just use Id
                fieldForFunction = 'Id';
            }

            String functionName;
            if (normalizedAgg == 'AVERAGE') {
                functionName = 'AVG';
            } else if (normalizedAgg == 'COUNT_DISTINCT') {
                functionName = 'COUNT_DISTINCT';
            } else {
                // SUM, MAX, MIN, COUNT
                functionName = normalizedAgg;
            }

            String soql =
                'SELECT ' + functionName + '(' + fieldForFunction + ') totalValue, ' +
                'COUNT(Id) totalRecords ' +
                'FROM ' + childObjectApiName +
                ' WHERE ' + whereClause;

            List<AggregateResult> aggResults = Database.query(soql);

            if (!aggResults.isEmpty()) {
                AggregateResult ar = aggResults[0];
                res.value = ar.get('totalValue');

                Object rc = ar.get('totalRecords');
                if (rc != null) {
                    res.recordCount = Integer.valueOf(String.valueOf(rc));
                }
            } else {
                res.value = null;
                res.recordCount = 0;
            }

            return res;

        } catch (Exception e) {
            // Surface the real error back to the tile so you can see the full message.
            res.errorMessage = 'Unexpected error while calculating rollup: ' + e.getMessage();
            return res;
        }
    }

    private static String normalizeAggregateType(String raw) {
        if (String.isBlank(raw)) {
            return 'SUM';
        }

        String v = raw.trim().toUpperCase();
        if (v == 'AVG') {
            v = 'AVERAGE';
        }

        // Allowed values
        Set<String> allowed = new Set<String>{
            'SUM',
            'AVERAGE',
            'MAX',
            'MIN',
            'COUNT',
            'COUNT_DISTINCT',
            'CONCATENATE',
            'CONCATENATE_DISTINCT'
        };

        if (!allowed.contains(v)) {
            return 'SUM';
        }
        return v;
    }
}
