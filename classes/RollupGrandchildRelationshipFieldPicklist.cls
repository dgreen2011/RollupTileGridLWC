global with sharing class RollupGrandchildRelationshipFieldPicklist extends VisualEditor.DynamicPickList {

    private VisualEditor.DesignTimePageContext context;

    global RollupGrandchildRelationshipFieldPicklist(VisualEditor.DesignTimePageContext context) {
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

            // Collect immediate child objects of the parent.
            Set<String> childApiNames = new Set<String>();
            Map<String, String> childApiToLabel = new Map<String, String>();

            for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
                Schema.SObjectType childType = childRel.getChildSObject();
                if (childType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
                String childApi = childDescribe.getName();
                if (String.isBlank(childApi)) {
                    continue;
                }

                childApiNames.add(childApi);
                childApiToLabel.put(childApi, childDescribe.getLabel());
            }

            // For each child, inspect its child relationships to find grandchildren
            // with lookups back to that child.
            Set<String> addedValues = new Set<String>();

            for (String childApi : childApiNames) {
                Schema.SObjectType childType = allTypes.get(childApi);
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
                    String grandLabel = grandDescribe.getLabel();

                    Schema.SObjectField relField = grandRel.getField();
                    if (relField == null) {
                        continue;
                    }

                    Schema.DescribeFieldResult fd = relField.getDescribe();

                    // Only reference fields are valid relationship fields.
                    if (fd.getType() != Schema.DisplayType.REFERENCE) {
                        continue;
                    }

                    // Only keep relationships that point back to this child object.
                    Boolean pointsToChild = false;
                    for (Schema.SObjectType refType : fd.getReferenceTo()) {
                        if (refType != null &&
                            refType.getDescribe().getName() == childApi) {
                            pointsToChild = true;
                            break;
                        }
                    }
                    if (!pointsToChild) {
                        continue;
                    }

                    String fieldApi = fd.getName();
                    String fieldLabel = fd.getLabel();

                    // Value stored in the design attribute; normalised later by RollupService.
                    String value = grandApi + '.' + fieldApi;
                    if (!addedValues.add(value)) {
                        continue;
                    }

                    // Label pattern:
                    // "Child Label • Grandchild Label • Field Label (GrandchildApi.FieldApi)"
                    String label =
                        childLabel +
                        ' • ' +
                        grandLabel +
                        ' • ' +
                        fieldLabel +
                        ' (' + value + ')';

                    rows.addRow(new VisualEditor.DataRow(label, value));
                }
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
