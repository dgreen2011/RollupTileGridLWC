global with sharing class RollupGrandchildObjectPicklist extends VisualEditor.DynamicPickList {

    private VisualEditor.DesignTimePageContext context;

    global RollupGrandchildObjectPicklist(VisualEditor.DesignTimePageContext context) {
        this.context = context;
    }

    // No default – avoid doing the heavy getValues() work twice.
    // Admin will explicitly choose the grandchild object.
    global override VisualEditor.DataRow getDefaultValue() {
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

            // Ensure we only show each grandchild object once.
            Set<String> seenGrandchildApis = new Set<String>();

            for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
                Schema.SObjectType childType = childRel.getChildSObject();
                if (childType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
                String childApi   = childDescribe.getName();
                String childLabel = childDescribe.getLabel();

                // Only include custom child objects, plus Account and Contact
                if (!childDescribe.isCustom() && childApi != 'Account' && childApi != 'Contact') {
                    continue;
                }

                // For each CHILD of the parent, look for its CHILDREN (grandchildren of the parent).
                for (Schema.ChildRelationship grandRel : childDescribe.getChildRelationships()) {
                    Schema.SObjectType grandType = grandRel.getChildSObject();
                    if (grandType == null) {
                        continue;
                    }

                    Schema.DescribeSObjectResult grandDescribe = grandType.getDescribe();
                    String grandApi   = grandDescribe.getName();
                    String grandLabel = grandDescribe.getLabel();

                    // Only include custom grandchild objects, plus Account and Contact
                    if (!grandDescribe.isCustom() && grandApi != 'Account' && grandApi != 'Contact') {
                        continue;
                    }

                    if (String.isBlank(grandApi) || !seenGrandchildApis.add(grandApi)) {
                        continue;
                    }

                    // Core label similar to other picklists:
                    // "Child Label • Grandchild Label (GrandchildApi)"
                    String coreLabel =
                        childLabel +
                        ' • ' +
                        grandLabel +
                        ' (' + grandApi + ')';

                    // Lowercase search tail so App Builder’s case-sensitive search
                    // matches things like "activity" or "myobject__c".
                    String searchTail =
                        ' | ' +
                        grandLabel.toLowerCase() +
                        ' (' + grandApi.toLowerCase() + ')';

                    String label = coreLabel + searchTail;

                    // DataRow(labelShownToAdmin, valueStoredInAttribute)
                    rows.addRow(new VisualEditor.DataRow(label, grandApi));
                }
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
