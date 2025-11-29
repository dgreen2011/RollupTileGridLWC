global with sharing class RollupAggregateFieldPicklist extends VisualEditor.DynamicPickList {

    private VisualEditor.DesignTimePageContext context;

    global RollupAggregateFieldPicklist(VisualEditor.DesignTimePageContext context) {
        this.context = context;
    }

    // No default – admin must choose one, or leave blank for COUNT-only rollups
    global override VisualEditor.DataRow getDefaultValue() {
        return null;
    }

    global override VisualEditor.DynamicPickListRows getValues() {
        VisualEditor.DynamicPickListRows rows = new VisualEditor.DynamicPickListRows();

        try {
            // 1) Prefer fields from child objects of the current record page object
            if (context != null && !String.isBlank(context.entityName)) {
                addFieldsForParent(context.entityName, rows);
            }

            // 2) Fallback: if we still have nothing (edge cases), add aggregateable
            //    fields from the whole org so the admin at least has options.
            if (rows.size() == 0) {
                addFieldsFromAllObjects(rows);
            }
        } catch (Exception e) {
            // Swallow everything – never let App Builder or the record page fail
            // because of this picklist.
        }

        return rows;
    }

    /**
     * Adds aggregate-friendly fields from all child (and grandchild) objects
     * of the given parent.
     *
     * Value stored in the design attribute = "ObjectApi.FieldApi".
     * Label shown in App Builder =
     *   "Child Label • Field Label (FieldApi) | field label (fieldapi)"
     *   OR
     *   "Child Label • Grandchild Label • Field Label (FieldApi) | field label (fieldapi)".
     *
     * The lowercase tail is there specifically so App Builder’s
     * case-sensitive search can match inputs like "aerial".
     */
    private void addFieldsForParent(
        String parentApiName,
        VisualEditor.DynamicPickListRows rows
    ) {
        Map<String, Schema.SObjectType> allTypes = Schema.getGlobalDescribe();
        Schema.SObjectType parentType = allTypes.get(parentApiName);
        if (parentType == null) return;

        Schema.DescribeSObjectResult parentDescribe = parentType.getDescribe();
        Set<String> seenKeys = new Set<String>();

        for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
            Schema.SObjectType childType = childRel.getChildSObject();
            if (childType == null) {
                continue;
            }

            Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
            String childApi   = childDescribe.getName();
            String childLabel = childDescribe.getLabel();

            // -------- 1) Fields on the CHILD object (existing behaviour) --------
            Map<String, Schema.SObjectField> fields = childDescribe.fields.getMap();

            for (Schema.SObjectField f : fields.values()) {
                Schema.DescribeFieldResult fd = f.getDescribe();
                if (!isAggregateFriendly(fd)) {
                    continue;
                }

                String fieldApi   = fd.getName();
                String fieldLabel = fd.getLabel();

                // Value used later by RollupService (normalised down to just the field).
                String key = childApi + '.' + fieldApi;

                // Don’t show duplicates for the same object+field combination
                if (!seenKeys.add(key)) {
                    continue;
                }

                // Core, nicely formatted label
                String coreLabel =
                    childLabel +
                    ' • ' +
                    fieldLabel +
                    ' (' + fieldApi + ')';

                // Lowercase tail for case-sensitive search to match things like "aerial"
                String searchTail =
                    ' | ' +
                    fieldLabel.toLowerCase() +
                    ' (' + fieldApi.toLowerCase() + ')';

                String label = coreLabel + searchTail;

                // DataRow(label, value)
                rows.addRow(new VisualEditor.DataRow(label, key));
            }

            // -------- 2) Fields on GRANDCHILD objects (new) --------
            // For each child of this child (i.e., grandchild of the parent),
            // surface aggregate-friendly fields as options too.
            for (Schema.ChildRelationship grandRel : childDescribe.getChildRelationships()) {
                Schema.SObjectType grandType = grandRel.getChildSObject();
                if (grandType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult grandDescribe = grandType.getDescribe();
                String grandApi   = grandDescribe.getName();
                String grandLabel = grandDescribe.getLabel();

                Map<String, Schema.SObjectField> grandFields = grandDescribe.fields.getMap();

                for (Schema.SObjectField gf : grandFields.values()) {
                    Schema.DescribeFieldResult gfd = gf.getDescribe();
                    if (!isAggregateFriendly(gfd)) {
                        continue;
                    }

                    String fieldApi   = gfd.getName();
                    String fieldLabel = gfd.getLabel();

                    // Value used later by RollupService – in grandchild mode
                    // we normalise this down to just the field API name.
                    String key = grandApi + '.' + fieldApi;

                    if (!seenKeys.add(key)) {
                        continue;
                    }

                    // Label pattern:
                    // "Child Label • Grandchild Label • Field Label (FieldApi)"
                    String coreLabel =
                        childLabel +
                        ' • ' +
                        grandLabel +
                        ' • ' +
                        fieldLabel +
                        ' (' + fieldApi + ')';

                    String searchTail =
                        ' | ' +
                        fieldLabel.toLowerCase() +
                        ' (' + fieldApi.toLowerCase() + ')';

                    String label = coreLabel + searchTail;

                    rows.addRow(new VisualEditor.DataRow(label, key));
                }
            }
        }
    }

    /**
     * Fallback: scan all objects in the org and add aggregate-friendly fields.
     *
     * Value stored in the design attribute = "ObjectApi.FieldApi".
     * Label shown in App Builder =
     *   "Object Label • Field Label (FieldApi) | field label (fieldapi)".
     */
    private void addFieldsFromAllObjects(VisualEditor.DynamicPickListRows rows) {
        Map<String, Schema.SObjectType> allTypes = Schema.getGlobalDescribe();
        Set<String> seenKeys = new Set<String>();

        for (Schema.SObjectType sType : allTypes.values()) {
            Schema.DescribeSObjectResult sDescribe = sType.getDescribe();

            if (sDescribe.isCustomSetting() || sDescribe.isDeprecatedAndHidden()) {
                continue;
            }

            String objectApi   = sDescribe.getName();
            String objectLabel = sDescribe.getLabel();

            Map<String, Schema.SObjectField> fields = sDescribe.fields.getMap();

            for (Schema.SObjectField f : fields.values()) {
                Schema.DescribeFieldResult fd = f.getDescribe();
                if (!isAggregateFriendly(fd)) {
                    continue;
                }

                String fieldApi   = fd.getName();
                String fieldLabel = fd.getLabel();

                String key = objectApi + '.' + fieldApi;
                if (!seenKeys.add(key)) {
                    continue;
                }

                String coreLabel =
                    objectLabel +
                    ' • ' +
                    fieldLabel +
                    ' (' + fieldApi + ')';

                String searchTail =
                    ' | ' +
                    fieldLabel.toLowerCase() +
                    ' (' + fieldApi.toLowerCase() + ')';

                String label = coreLabel + searchTail;

                rows.addRow(new VisualEditor.DataRow(label, key));
            }
        }
    }

    /**
     * Decide which field types make sense for SUM/AVG/COUNT/CONCATENATE/FIRST/LAST.
     */
    private static Boolean isAggregateFriendly(Schema.DescribeFieldResult fd) {
        Schema.DisplayType t = fd.getType();
        return
            t == Schema.DisplayType.Currency ||
            t == Schema.DisplayType.Double   ||
            t == Schema.DisplayType.Integer  ||
            t == Schema.DisplayType.Long     ||
            t == Schema.DisplayType.Percent  ||
            t == Schema.DisplayType.Date     ||
            t == Schema.DisplayType.DateTime ||
            t == Schema.DisplayType.String;
    }

    // Prevents Salesforce from re-calling getValues() during validation
    // and avoids weird Community/App Builder bugs.
    global override Boolean isValid(Object attributeValue) {
        return true;
    }
}
