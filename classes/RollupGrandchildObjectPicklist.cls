global with sharing class RollupGrandchildObjectPicklist extends VisualEditor.DynamicPickList {

    private VisualEditor.DesignTimePageContext context;

    global RollupGrandchildObjectPicklist(VisualEditor.DesignTimePageContext context) {
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

            Map<String, Schema.SObjectType> allTypes = Schema.getGlobalDescribe();
            Schema.SObjectType parentType = allTypes.get(context.entityName);
            if (parentType == null) {
                return rows;
            }

            Schema.DescribeSObjectResult parentDescribe = parentType.getDescribe();
            Set<String> seenApiNames = new Set<String>();

            // For each CHILD of the parent, look for its CHILDREN (grandchildren of the parent).
            for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
                Schema.SObjectType childType = childRel.getChildSObject();
                if (childType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
                String childLabel = childDescribe.getLabel();

                for (Schema.ChildRelationship grandRel : childDescribe.getChildRelationships()) {
                    Schema.SObjectType grandType = grandRel.getChildSObject();
                    if (grandType == null) {
                        continue;
                    }

                    Schema.DescribeSObjectResult grandDescribe = grandType.getDescribe();
                    String grandApi = grandDescribe.getName();
                    if (String.isBlank(grandApi) || !seenApiNames.add(grandApi)) {
                        continue;
                    }

                    String grandLabel = grandDescribe.getLabel();

                    // Label pattern: "Child Label • Grandchild Label (GrandchildApi)"
                    String label =
                        childLabel +
                        ' • ' +
                        grandLabel +
                        ' (' + grandApi + ')';

                    rows.addRow(
                        new VisualEditor.DataRow(
                            label,
                            grandApi
                        )
                    );
                }
            }
        } catch (Exception e) {
            // Swallow errors so Lightning pages and App Builder keep loading.
        }

        return rows;
    }

    // Important so getValues() isn’t called again at runtime and doesn’t break pages.
    global override Boolean isValid(Object attributeValue) {
        return true;
    }
}
