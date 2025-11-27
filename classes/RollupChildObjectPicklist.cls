global with sharing class RollupChildObjectPicklist extends VisualEditor.DynamicPickList {

    private VisualEditor.DesignTimePageContext context;

    global RollupChildObjectPicklist(VisualEditor.DesignTimePageContext context) {
        this.context = context;
    }

    global override VisualEditor.DataRow getDefaultValue() {
        VisualEditor.DynamicPickListRows rows = getValues();
        if (rows != null && rows.size() > 0) {
            return rows.get(0);
        }
        return null;
    }

    global override VisualEditor.DynamicPickListRows getValues() {
        VisualEditor.DynamicPickListRows rows = new VisualEditor.DynamicPickListRows();

        try {
            if (context == null || String.isBlank(context.entityName)) {
                return rows;
            }

            Schema.SObjectType parentType = Schema.getGlobalDescribe().get(context.entityName);
            if (parentType == null) {
                return rows;
            }

            Schema.DescribeSObjectResult parentDescribe = parentType.getDescribe();
            String parentLabel = parentDescribe.getLabel();

            Set<String> seenApiNames = new Set<String>();

            for (Schema.ChildRelationship cr : parentDescribe.getChildRelationships()) {
                Schema.SObjectType childType = cr.getChildSObject();
                if (childType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
                String apiName = childDescribe.getName();
                if (String.isBlank(apiName) || !seenApiNames.add(apiName)) {
                    continue;
                }

                String childLabel = childDescribe.getLabel();

                // Consistent formatting with the other picklists:
                // "Parent Label • Child Label (ChildApi)"
                String label =
                    parentLabel +
                    ' • ' +
                    childLabel +
                    ' (' + apiName + ')';

                // DataRow(labelShownToAdmin, valueStoredInAttribute)
                rows.addRow(
                    new VisualEditor.DataRow(
                        label,
                        apiName
                    )
                );
            }
        } catch (Exception e) {
            // Swallow errors so Lightning pages and App Builder keep loading
            // even if something goes wrong here.
        }

        return rows;
    }

    // Important so getValues() isn’t called again at runtime and doesn’t break pages.
    global override Boolean isValid(Object attributeValue) {
        return true;
    }
}
