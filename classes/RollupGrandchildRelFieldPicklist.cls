global with sharing class RollupGrandchildRelFieldPicklist extends VisualEditor.DynamicPickList {

    private VisualEditor.DesignTimePageContext context;

    global RollupGrandchildRelFieldPicklist(VisualEditor.DesignTimePageContext context) {
        this.context = context;
    }

    // No default – avoid doing the heavy getValues() work twice.
    // Admin will explicitly choose the lookup on the grandchild.
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

            // Ensure we only show each (grandchild object + field) once.
            Set<String> addedValues = new Set<String>();

            // For each CHILD of the parent…
            for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
                Schema.SObjectType childType = childRel.getChildSObject();
                if (childType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
                String childApi   = childDescribe.getName();
                String childLabel = childDescribe.getLabel();

                // …look for its CHILDREN (grandchildren of the parent).
                for (Schema.ChildRelationship grandRel : childDescribe.getChildRelationships()) {
                    Schema.SObjectType grandType = grandRel.getChildSObject();
                    if (grandType == null) {
                        continue;
                    }

                    Schema.DescribeSObjectResult grandDescribe = grandType.getDescribe();
                    String grandApi   = grandDescribe.getName();
                    String grandLabel = grandDescribe.getLabel();

                    // Relationship field on the GRANDCHILD that points to the CHILD.
                    Schema.SObjectField relField = grandRel.getField();
                    if (relField == null) {
                        continue;
                    }

                    Schema.DescribeFieldResult fd = relField.getDescribe();

                    // Only reference fields make sense here.
                    if (fd.getType() != Schema.DisplayType.REFERENCE) {
                        continue;
                    }

                    // Only keep relationships that actually point back to this child object.
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

                    String fieldApi   = fd.getName();
                    String fieldLabel = fd.getLabel();

                    // Value stored in the design attribute; normalised later by RollupService.
                    // Example: "sitetracker__Activity__c.sitetracker__Project__c"
                    String value = grandApi + '.' + fieldApi;
                    if (!addedValues.add(value)) {
                        continue;
                    }

                    // Core label:
                    // "Child Label • Grandchild Label • Field Label (GrandchildApi.FieldApi)"
                    String coreLabel =
                        childLabel +
                        ' • ' +
                        grandLabel +
                        ' • ' +
                        fieldLabel +
                        ' (' + value + ')';

                    // Lowercase search tail so App Builder’s case-sensitive search
                    // also matches things like "activity", "project", etc.
                    String searchTail =
                        ' | ' +
                        childLabel.toLowerCase() +
                        ' • ' +
                        grandLabel.toLowerCase() +
                        ' • ' +
                        fieldLabel.toLowerCase() +
                        ' (' + fieldApi.toLowerCase() + ')';

                    String label = coreLabel + searchTail;

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
