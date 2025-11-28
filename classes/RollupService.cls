public with sharing class RollupService {

    // DTO returned to the LWC
    public class RollupResult {
        @AuraEnabled public String  value;        // Aggregated value (always serialized as string)
        @AuraEnabled public Integer recordCount;  // Number of child records considered
        @AuraEnabled public Boolean isCurrency;   // True if aggregate field is a currency field
        @AuraEnabled public Boolean isPercent;    // True if aggregate field is a percent field
        @AuraEnabled public String  fieldLabel;   // Label for the aggregate field
        @AuraEnabled public String  errorMessage; // Any business / configuration error
    }

    // Numeric aggregation types (based on what the LWC sends)
    private static final Set<String> NUMERIC_AGG_TYPES =
        new Set<String>{ 'SUM', 'AVERAGE', 'AVG', 'MAX', 'MIN' };

    // Field types that support numeric aggregates
    private static final Set<Schema.DisplayType> NUMERIC_FIELD_TYPES =
        new Set<Schema.DisplayType>{
            Schema.DisplayType.Currency,
            Schema.DisplayType.Double,
            Schema.DisplayType.Integer,
            Schema.DisplayType.Long,
            Schema.DisplayType.Percent
        };

    // Field types that are safe for concatenation (and FIRST/LAST)
    private static final Set<Schema.DisplayType> TEXTY_FIELD_TYPES =
        new Set<Schema.DisplayType>{
            Schema.DisplayType.String,
            Schema.DisplayType.TextArea,
            Schema.DisplayType.Phone,
            Schema.DisplayType.Email,
            Schema.DisplayType.Url,
            Schema.DisplayType.Picklist,
            Schema.DisplayType.MultiPicklist
        };

    @AuraEnabled(cacheable=true)
    public static RollupResult getRollup(
        Id      parentId,
        String  childObjectApiName,
        String  relationshipFieldApiName,
        String  aggregateFieldApiName,
        String  aggregateType,
        String  filterCondition
    ) {
        RollupResult result = new RollupResult();

        try {
            // ---- Basic validation -------------------------------------------------------
            if (parentId == null) {
                result.errorMessage = 'Rollup is not correctly configured: missing parentId.';
                return result;
            }
            if (String.isBlank(childObjectApiName)) {
                result.errorMessage = 'Rollup is not correctly configured: Child Object is required.';
                return result;
            }
            if (String.isBlank(relationshipFieldApiName)) {
                result.errorMessage = 'Rollup is not correctly configured: Relationship Field is required.';
                return result;
            }

            String normalizedAggType = normalizeAggregateType(aggregateType);

            // ---- Describe child object and fields --------------------------------------
            Map<String, Schema.SObjectType> gd = Schema.getGlobalDescribe();
            if (!gd.containsKey(childObjectApiName)) {
                result.errorMessage = 'Child object ' + childObjectApiName + ' does not exist or is not accessible.';
                return result;
            }

            Schema.SObjectType childType        = gd.get(childObjectApiName);
            Schema.DescribeSObjectResult childD = childType.getDescribe();
            Map<String, Schema.SObjectField> fieldMap = childD.fields.getMap();

            if (!fieldMap.containsKey(relationshipFieldApiName)) {
                result.errorMessage = 'Relationship field ' + relationshipFieldApiName + ' does not exist on ' + childObjectApiName + '.';
                return result;
            }

            Schema.DescribeFieldResult aggFieldD;
            if (!String.isBlank(aggregateFieldApiName) && fieldMap.containsKey(aggregateFieldApiName)) {
                aggFieldD = fieldMap.get(aggregateFieldApiName).getDescribe();
            }

            // Populate fieldLabel / currency / percent flags (best effort)
            if (aggFieldD != null) {
                result.fieldLabel = aggFieldD.getLabel();
                Schema.DisplayType dt = aggFieldD.getType();
                result.isCurrency = (dt == Schema.DisplayType.Currency);
                result.isPercent  = (dt == Schema.DisplayType.Percent);
            } else {
                result.fieldLabel = aggregateFieldApiName;
                result.isCurrency = false;
                result.isPercent  = false;
            }

            // ---- Build WHERE clause (uses :parentId binding) ---------------------------
            String whereClause = relationshipFieldApiName + ' = :parentId';
            if (!String.isBlank(filterCondition)) {
                String trimmed = filterCondition.trim();
                if (!trimmed.toUpperCase().startsWith('AND ')) {
                    whereClause += ' AND ' + trimmed;
                } else {
                    whereClause += ' ' + trimmed;
                }
            }

            // ---- Branch by aggregation type -------------------------------------------

            // 1) COUNT – ALWAYS use COUNT(), never SUM on the text field
            if (normalizedAggType == 'COUNT') {
                String soql =
                    'SELECT COUNT() totalCount ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause;

                List<AggregateResult> ars =
                    (List<AggregateResult>)Database.query(soql);
                AggregateResult ar = ars.isEmpty() ? null : ars[0];

                Integer countVal = 0;
                if (ar != null) {
                    Object rawCount = ar.get('totalCount');
                    if (rawCount != null) {
                        countVal = Integer.valueOf(String.valueOf(rawCount));
                    }
                }

                result.recordCount = countVal;
                result.value       = String.valueOf(countVal);
                return result;
            }

            // 2) COUNT_DISTINCT – COUNT_DISTINCT(field) and also COUNT() for recordCount
            if (normalizedAggType == 'COUNT_DISTINCT') {
                if (aggFieldD == null || String.isBlank(aggregateFieldApiName)) {
                    result.errorMessage = 'Count Distinct requires an Aggregate Field.';
                    return result;
                }

                String soql =
                    'SELECT COUNT_DISTINCT(' + aggregateFieldApiName + ') distinctCount, ' +
                           'COUNT() totalCount ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause;

                List<AggregateResult> ars =
                    (List<AggregateResult>)Database.query(soql);
                AggregateResult ar = ars.isEmpty() ? null : ars[0];

                Integer distinctVal = 0;
                Integer totalVal    = 0;

                if (ar != null) {
                    Object rawDistinct = ar.get('distinctCount');
                    Object rawTotal    = ar.get('totalCount');

                    if (rawDistinct != null) {
                        distinctVal = Integer.valueOf(String.valueOf(rawDistinct));
                    }
                    if (rawTotal != null) {
                        totalVal = Integer.valueOf(String.valueOf(rawTotal));
                    }
                }

                result.recordCount = totalVal;
                result.value       = String.valueOf(distinctVal);
                return result;
            }

            // 3) CONCATENATE / CONCATENATE_DISTINCT – build value in Apex
            if (normalizedAggType == 'CONCATENATE' || normalizedAggType == 'CONCATENATE_DISTINCT') {
                if (aggFieldD == null || String.isBlank(aggregateFieldApiName)) {
                    result.errorMessage = 'Concatenate requires an Aggregate Field.';
                    return result;
                }

                // Ensure the field is string-ish
                if (!TEXTY_FIELD_TYPES.contains(aggFieldD.getType())) {
                    result.errorMessage =
                        'Field ' + result.fieldLabel + ' does not support concatenation.';
                    return result;
                }

                String soql =
                    'SELECT ' + aggregateFieldApiName + ' val ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause +
                    ' ORDER BY ' + aggregateFieldApiName;

                List<SObject> rows = Database.query(soql);
                result.recordCount = rows.size();

                List<String> pieces = new List<String>();
                Set<String> seen    = new Set<String>();

                for (SObject row : rows) {
                    Object raw = row.get(aggregateFieldApiName);
                    String s   = (raw == null) ? null : String.valueOf(raw);
                    if (String.isBlank(s)) {
                        continue;
                    }

                    if (normalizedAggType == 'CONCATENATE_DISTINCT') {
                        if (!seen.contains(s)) {
                            seen.add(s);
                            pieces.add(s);
                        }
                    } else {
                        pieces.add(s);
                    }
                }

                if (!pieces.isEmpty()) {
                    result.value = String.join(pieces, '; ');
                } else {
                    result.value = null;
                }
                return result;
            }

            // 4) FIRST / LAST – take first/last non-null value based on CreatedDate
            if (normalizedAggType == 'FIRST' || normalizedAggType == 'LAST') {
                if (aggFieldD == null || String.isBlank(aggregateFieldApiName)) {
                    result.errorMessage = 'First/Last requires an Aggregate Field.';
                    return result;
                }

                String sortDir = (normalizedAggType == 'FIRST') ? 'ASC' : 'DESC';

                String valueSoql =
                    'SELECT ' + aggregateFieldApiName + ' val ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause +
                    ' ORDER BY CreatedDate ' + sortDir +
                    ' LIMIT 1';

                List<SObject> rows = Database.query(valueSoql);

                String countSoql =
                    'SELECT COUNT() totalCount ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause;

                List<AggregateResult> countResults =
                    (List<AggregateResult>)Database.query(countSoql);
                AggregateResult countAr = countResults.isEmpty() ? null : countResults[0];

                Integer totalVal = 0;
                if (countAr != null) {
                    Object rawTotal = countAr.get('totalCount');
                    if (rawTotal != null) {
                        totalVal = Integer.valueOf(String.valueOf(rawTotal));
                    }
                }

                result.recordCount = totalVal;

                if (!rows.isEmpty()) {
                    Object rawVal = rows[0].get(aggregateFieldApiName);
                    result.value  = (rawVal == null) ? null : String.valueOf(rawVal);
                } else {
                    result.value = null;
                }

                return result;
            }

            // 5) Numeric aggregates: SUM / AVERAGE / MAX / MIN (and AVG alias)
            if (NUMERIC_AGG_TYPES.contains(normalizedAggType)) {
                if (aggFieldD == null || String.isBlank(aggregateFieldApiName)) {
                    result.errorMessage =
                        'Numeric aggregate ' + normalizedAggType + ' requires an Aggregate Field.';
                    return result;
                }

                // Check field type compatibility
                if (!NUMERIC_FIELD_TYPES.contains(aggFieldD.getType())) {
                    result.errorMessage =
                        'Field ' + result.fieldLabel +
                        ' does not support aggregate operator ' +
                        displayAggName(normalizedAggType) + '.';
                    return result;
                }

                String soqlFn = mapToSoqlNumericFunction(normalizedAggType); // e.g. AVERAGE -> AVG

                String soql =
                    'SELECT ' + soqlFn + '(' + aggregateFieldApiName + ') aggVal, ' +
                           'COUNT() totalCount ' +
                    'FROM ' + childObjectApiName +
                    ' WHERE ' + whereClause;

                List<AggregateResult> ars =
                    (List<AggregateResult>)Database.query(soql);
                AggregateResult ar = ars.isEmpty() ? null : ars[0];

                Object rawAgg   = (ar == null) ? null : ar.get('aggVal');
                Object rawTotal = (ar == null) ? null : ar.get('totalCount');

                Integer totalVal = 0;
                if (rawTotal != null) {
                    totalVal = Integer.valueOf(String.valueOf(rawTotal));
                }

                result.recordCount = totalVal;
                result.value       = (rawAgg == null) ? null : String.valueOf(rawAgg);

                return result;
            }

            // 6) Fallback: Unknown aggregate type – behave like COUNT()
            String fallbackSoql =
                'SELECT COUNT() totalCount ' +
                'FROM ' + childObjectApiName +
                ' WHERE ' + whereClause;

            List<AggregateResult> fbResults =
                (List<AggregateResult>)Database.query(fallbackSoql);
            AggregateResult fbAr = fbResults.isEmpty() ? null : fbResults[0];

            Integer fbTotal = 0;
            if (fbAr != null) {
                Object rawTotal = fbAr.get('totalCount');
                if (rawTotal != null) {
                    fbTotal = Integer.valueOf(String.valueOf(rawTotal));
                }
            }

            result.recordCount = fbTotal;
            result.value       = String.valueOf(fbTotal);
            return result;

        } catch (Exception ex) {
            // Catch *all* errors so they come back as an errorMessage (no hard failure in LWC)
            result.errorMessage = 'Unexpected error while calculating rollup: ' + ex.getMessage();
            return result;
        }
    }

    // --- Helper methods --------------------------------------------------------------

    private static String normalizeAggregateType(String agg) {
        if (String.isBlank(agg)) {
            return 'SUM';
        }
        String norm = agg.trim().toUpperCase();
        if (norm == 'AVG') {
            norm = 'AVERAGE';
        }
        if (norm == 'COUNT()') {
            norm = 'COUNT';
        }
        return norm;
    }

    private static String mapToSoqlNumericFunction(String agg) {
        String norm = normalizeAggregateType(agg);
        if (norm == 'AVERAGE') {
            return 'AVG';
        }
        // SUM, MAX, MIN already match SOQL function names
        return norm;
    }

    private static String displayAggName(String agg) {
        String norm = normalizeAggregateType(agg);
        if (norm == 'AVERAGE') return 'AVERAGE';
        if (norm == 'SUM')     return 'SUM';
        if (norm == 'MAX')     return 'MAX';
        if (norm == 'MIN')     return 'MIN';
        if (norm == 'COUNT')   return 'COUNT';
        if (norm == 'COUNT_DISTINCT') return 'COUNT_DISTINCT';
        return norm;
    }
}
