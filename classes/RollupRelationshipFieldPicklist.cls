global with sharing class RollupRelationshipFieldPicklist extends VisualEditor.DynamicPickList {

    private VisualEditor.DesignTimePageContext context;

    global RollupRelationshipFieldPicklist(VisualEditor.DesignTimePageContext context) {
        this.context = context;
    }

    // No default – avoid doing the heavy getValues() work twice.
    // Admin will explicitly choose the relationship field.
    global override VisualEditor.DataRow getDefaultValue() {
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
            String parentApiName = parentDescribe.getName();

            // Use a key that includes both child + field so we don't show duplicates
            Set<String> addedKeys = new Set<String>();

            for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
                Schema.SObjectType childType = childRel.getChildSObject();
                if (childType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
                String childApi = childDescribe.getName();
                String childLabel = childDescribe.getLabel();

                Schema.SObjectField relField = childRel.getField();
                if (relField == null) {
                    continue;
                }

                Schema.DescribeFieldResult fieldDescribe = relField.getDescribe();

                // Only reference fields make sense here
                if (fieldDescribe.getType() != Schema.DisplayType.REFERENCE) {
                    continue;
                }

                // Only keep relationships that can point back to the parent object
                List<Schema.SObjectType> refTo = fieldDescribe.getReferenceTo();
                Boolean pointsToParent = false;
                for (Schema.SObjectType refType : refTo) {
                    if (refType != null && refType.getDescribe().getName() == parentApiName) {
                        pointsToParent = true;
                        break;
                    }
                }
                if (!pointsToParent) {
                    continue;
                }

                String fieldApi = fieldDescribe.getName();
                String fieldLabel = fieldDescribe.getLabel();

                // Value stored in the design attribute; used later by RollupService
                String value = childApi + '.' + fieldApi;
                if (!addedKeys.add(value)) {
                    continue;
                }

                String label =
                    childLabel +
                    ' • ' +
                    fieldLabel +
                    ' (' + value + ')';

                // DataRow(labelShownToAdmin, valueStoredInAttribute)
                rows.addRow(new VisualEditor.DataRow(label, value));
            }
        } catch (Exception e) {
            // Contain all errors inside the picklist; never let them reach the flexipage engine.
        }

        return rows;
    }

    // Prevents Salesforce from re-calling getValues() during validation
    // and avoids weird Community/App Builder bugs.
    global override Boolean isValid(Object attributeValue) {
        return true;
    }
}
