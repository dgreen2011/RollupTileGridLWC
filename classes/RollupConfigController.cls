public with sharing class RollupConfigController {

    // ------------------------------------------------------------
    // Simple DTOs for LWC
    // ------------------------------------------------------------

    public class Option {
        @AuraEnabled public String label;
        @AuraEnabled public String value;
    }

    public class RollupConfigDTO {
        @AuraEnabled public String parentObjectApiName;
        @AuraEnabled public String configKey;
        @AuraEnabled public String childObjectApiName;
        @AuraEnabled public String relationshipFieldApiName;
        @AuraEnabled public String grandchildObjectApiName;
        @AuraEnabled public String grandchildRelationshipFieldApiName;
    }

    // ------------------------------------------------------------
    // Schema-driven option loaders
    // ------------------------------------------------------------

    /**
     * Child objects of the given parent.
     * Label pattern: "Parent Label • Child Label (ChildApi)"
     * Value: Child object API name (e.g., "Contact").
     */
    @AuraEnabled(cacheable=true)
    public static List<Option> getChildObjects(String parentObjectApiName) {
        List<Option> options = new List<Option>();

        if (String.isBlank(parentObjectApiName)) {
            return options;
        }

        Map<String, Schema.SObjectType> allTypes = Schema.getGlobalDescribe();
        if (!allTypes.containsKey(parentObjectApiName)) {
            return options;
        }

        Schema.SObjectType parentType = allTypes.get(parentObjectApiName);
        if (parentType == null) {
            return options;
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

            Option opt = new Option();
            opt.label =
                parentLabel +
                ' • ' +
                childLabel +
                ' (' + apiName + ')';
            opt.value = apiName; // plain object API name

            options.add(opt);
        }

        return options;
    }

    /**
     * Relationship fields on the *selected child* that point back to the parent.
     * Label pattern: "Child Label • Field Label (ChildApi.FieldApi)"
     * Value: Field API name (e.g., "AccountId").
     */
    @AuraEnabled(cacheable=true)
    public static List<Option> getChildRelationshipFields(
        String parentObjectApiName,
        String childObjectApiName
    ) {
        List<Option> options = new List<Option>();

        if (String.isBlank(parentObjectApiName) || String.isBlank(childObjectApiName)) {
            return options;
        }

        Map<String, Schema.SObjectType> allTypes = Schema.getGlobalDescribe();
        if (!allTypes.containsKey(parentObjectApiName)) {
            return options;
        }

        Schema.SObjectType parentType = allTypes.get(parentObjectApiName);
        if (parentType == null) {
            return options;
        }

        Schema.DescribeSObjectResult parentDescribe = parentType.getDescribe();
        String parentApiName = parentDescribe.getName();

        Set<String> addedFieldApis = new Set<String>();

        for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
            Schema.SObjectType childType = childRel.getChildSObject();
            if (childType == null) {
                continue;
            }

            Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
            String childApi = childDescribe.getName();

            // Only relationships for the chosen child object
            if (childApi != childObjectApiName) {
                continue;
            }

            String childLabel = childDescribe.getLabel();

            Schema.SObjectField relField = childRel.getField();
            if (relField == null) {
                continue;
            }

            Schema.DescribeFieldResult fieldDescribe = relField.getDescribe();

            // Only reference fields make sense
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
            if (!addedFieldApis.add(fieldApi)) {
                continue;
            }

            String fieldLabel = fieldDescribe.getLabel();

            Option opt = new Option();
            opt.label =
                childLabel +
                ' • ' +
                fieldLabel +
                ' (' + childApi + '.' + fieldApi + ')';
            opt.value = fieldApi; // plain field API name

            options.add(opt);
        }

        return options;
    }

    /**
     * Grandchild objects that are children of the selected child object.
     * Label pattern: "Child Label • Grandchild Label (GrandchildApi) | grandchild label (grandchildapi)"
     * Value: Grandchild object API name.
     */
    @AuraEnabled(cacheable=true)
    public static List<Option> getGrandchildObjects(
        String parentObjectApiName,
        String childObjectApiName
    ) {
        List<Option> options = new List<Option>();

        if (String.isBlank(parentObjectApiName) || String.isBlank(childObjectApiName)) {
            return options;
        }

        Map<String, Schema.SObjectType> allTypes = Schema.getGlobalDescribe();
        if (!allTypes.containsKey(parentObjectApiName)) {
            return options;
        }

        Schema.SObjectType parentType = allTypes.get(parentObjectApiName);
        if (parentType == null) {
            return options;
        }

        Schema.DescribeSObjectResult parentDescribe = parentType.getDescribe();

        Set<String> seenGrandchildApis = new Set<String>();

        for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
            Schema.SObjectType childType = childRel.getChildSObject();
            if (childType == null) {
                continue;
            }

            Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
            String childApi = childDescribe.getName();

            // Only expand the selected child
            if (childApi != childObjectApiName) {
                continue;
            }

            String childLabel = childDescribe.getLabel();

            // For this CHILD, look for its CHILDREN (grandchildren of the parent)
            for (Schema.ChildRelationship grandRel : childDescribe.getChildRelationships()) {
                Schema.SObjectType grandType = grandRel.getChildSObject();
                if (grandType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult grandDescribe = grandType.getDescribe();
                String grandApi   = grandDescribe.getName();
                String grandLabel = grandDescribe.getLabel();

                if (String.isBlank(grandApi) || !seenGrandchildApis.add(grandApi)) {
                    continue;
                }

                String coreLabel =
                    childLabel +
                    ' • ' +
                    grandLabel +
                    ' (' + grandApi + ')';

                String searchTail =
                    ' | ' +
                    grandLabel.toLowerCase() +
                    ' (' + grandApi.toLowerCase() + ')';

                Option opt = new Option();
                opt.label = coreLabel + searchTail;
                opt.value = grandApi; // plain object API name

                options.add(opt);
            }
        }

        return options;
    }

    /**
     * Relationship fields on the selected grandchild that point back to the selected child.
     * Label pattern:
     *   "Child Label • Grandchild Label • Field Label (GrandchildApi.FieldApi) | ..."
     * Value: Field API name on the grandchild (e.g., "Project__c").
     */
    @AuraEnabled(cacheable=true)
    public static List<Option> getGrandchildRelationshipFields(
        String parentObjectApiName,
        String childObjectApiName,
        String grandchildObjectApiName
    ) {
        List<Option> options = new List<Option>();

        if (String.isBlank(parentObjectApiName)
            || String.isBlank(childObjectApiName)
            || String.isBlank(grandchildObjectApiName)) {
            return options;
        }

        Map<String, Schema.SObjectType> allTypes = Schema.getGlobalDescribe();
        if (!allTypes.containsKey(parentObjectApiName)) {
            return options;
        }

        Schema.SObjectType parentType = allTypes.get(parentObjectApiName);
        if (parentType == null) {
            return options;
        }

        Schema.DescribeSObjectResult parentDescribe = parentType.getDescribe();

        Set<String> addedFieldApis = new Set<String>();

        // For each CHILD of the parent…
        for (Schema.ChildRelationship childRel : parentDescribe.getChildRelationships()) {
            Schema.SObjectType childType = childRel.getChildSObject();
            if (childType == null) {
                continue;
            }

            Schema.DescribeSObjectResult childDescribe = childType.getDescribe();
            String childApi   = childDescribe.getName();
            String childLabel = childDescribe.getLabel();

            // Only expand the selected child object
            if (childApi != childObjectApiName) {
                continue;
            }

            // …look for its CHILDREN (grandchildren of the parent)
            for (Schema.ChildRelationship grandRel : childDescribe.getChildRelationships()) {
                Schema.SObjectType grandType = grandRel.getChildSObject();
                if (grandType == null) {
                    continue;
                }

                Schema.DescribeSObjectResult grandDescribe = grandType.getDescribe();
                String grandApi   = grandDescribe.getName();
                String grandLabel = grandDescribe.getLabel();

                // Only show relationships for the selected grandchild object
                if (grandApi != grandchildObjectApiName) {
                    continue;
                }

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

                if (!addedFieldApis.add(fieldApi)) {
                    continue;
                }

                String coreLabel =
                    childLabel +
                    ' • ' +
                    grandLabel +
                    ' • ' +
                    fieldLabel +
                    ' (' + grandApi + '.' + fieldApi + ')';

                String searchTail =
                    ' | ' +
                    childLabel.toLowerCase() +
                    ' • ' +
                    grandLabel.toLowerCase() +
                    ' • ' +
                    fieldLabel.toLowerCase() +
                    ' (' + fieldApi.toLowerCase() + ')';

                Option opt = new Option();
                opt.label = coreLabel + searchTail;
                opt.value = fieldApi; // plain field API name on the grandchild

                options.add(opt);
            }
        }

        return options;
    }

    // ------------------------------------------------------------
    // Persisted config (Custom Object)
    // ------------------------------------------------------------
    // NOTE: This assumes you create a custom object:
    //   Rollup_Relationship_Config__c
    // with these Text fields:
    //   - Parent_Object__c
    //   - Config_Key__c
    //   - Child_Object__c
    //   - Relationship_Field__c
    //   - Grandchild_Object__c
    //   - Grandchild_Relationship_Field__c

    @AuraEnabled(cacheable=true)
    public static RollupConfigDTO getConfig(
        String parentObjectApiName,
        String configKey
    ) {
        if (String.isBlank(parentObjectApiName)) {
            return null;
        }

        Rollup_Relationship_Config__c rec;

        if (!String.isBlank(configKey)) {
            List<Rollup_Relationship_Config__c> rowsWithKey = [
                SELECT Id,
                       Parent_Object__c,
                       Config_Key__c,
                       Child_Object__c,
                       Relationship_Field__c,
                       Grandchild_Object__c,
                       Grandchild_Relationship_Field__c
                FROM Rollup_Relationship_Config__c
                WHERE Parent_Object__c = :parentObjectApiName
                  AND Config_Key__c = :configKey
                LIMIT 1
            ];
            if (!rowsWithKey.isEmpty()) {
                rec = rowsWithKey[0];
            }
        } else {
            List<Rollup_Relationship_Config__c> rowsNoKey = [
                SELECT Id,
                       Parent_Object__c,
                       Config_Key__c,
                       Child_Object__c,
                       Relationship_Field__c,
                       Grandchild_Object__c,
                       Grandchild_Relationship_Field__c
                FROM Rollup_Relationship_Config__c
                WHERE Parent_Object__c = :parentObjectApiName
                  AND Config_Key__c = NULL
                ORDER BY CreatedDate DESC
                LIMIT 1
            ];
            if (!rowsNoKey.isEmpty()) {
                rec = rowsNoKey[0];
            }
        }

        if (rec == null) {
            return null;
        }

        RollupConfigDTO dto = new RollupConfigDTO();
        dto.parentObjectApiName              = rec.Parent_Object__c;
        dto.configKey                        = rec.Config_Key__c;
        dto.childObjectApiName               = rec.Child_Object__c;
        dto.relationshipFieldApiName         = rec.Relationship_Field__c;
        dto.grandchildObjectApiName          = rec.Grandchild_Object__c;
        dto.grandchildRelationshipFieldApiName = rec.Grandchild_Relationship_Field__c;

        return dto;
    }

    @AuraEnabled
    public static void saveConfig(
        String parentObjectApiName,
        String configKey,
        String childObjectApiName,
        String relationshipFieldApiName,
        String grandchildObjectApiName,
        String grandchildRelationshipFieldApiName
    ) {
        if (String.isBlank(parentObjectApiName)) {
            throw new AuraHandledException('parentObjectApiName is required.');
        }
        if (String.isBlank(childObjectApiName)) {
            throw new AuraHandledException('Child Object is required.');
        }
        if (String.isBlank(relationshipFieldApiName)) {
            throw new AuraHandledException('Relationship Field (lookup on child) is required.');
        }

        // Grandchild config must be "all or nothing"
        Boolean hasGrandchildObject = !String.isBlank(grandchildObjectApiName);
        Boolean hasGrandchildRel    = !String.isBlank(grandchildRelationshipFieldApiName);
        if (hasGrandchildObject ^ hasGrandchildRel) {
            throw new AuraHandledException(
                'If you use a grandchild object, you must also choose its relationship field (and vice versa).'
            );
        }

        Rollup_Relationship_Config__c rec;

        if (!String.isBlank(configKey)) {
            List<Rollup_Relationship_Config__c> rowsWithKey = [
                SELECT Id
                FROM Rollup_Relationship_Config__c
                WHERE Parent_Object__c = :parentObjectApiName
                  AND Config_Key__c = :configKey
                LIMIT 1
            ];
            if (!rowsWithKey.isEmpty()) {
                rec = rowsWithKey[0];
            }
        } else {
            List<Rollup_Relationship_Config__c> rowsNoKey = [
                SELECT Id
                FROM Rollup_Relationship_Config__c
                WHERE Parent_Object__c = :parentObjectApiName
                  AND Config_Key__c = NULL
                LIMIT 1
            ];
            if (!rowsNoKey.isEmpty()) {
                rec = rowsNoKey[0];
            }
        }

        if (rec == null) {
            rec = new Rollup_Relationship_Config__c();
        }

        rec.Parent_Object__c                  = parentObjectApiName;
        rec.Config_Key__c                     = configKey;
        rec.Child_Object__c                   = childObjectApiName;
        rec.Relationship_Field__c             = relationshipFieldApiName;
        rec.Grandchild_Object__c              = grandchildObjectApiName;
        rec.Grandchild_Relationship_Field__c  = grandchildRelationshipFieldApiName;

        upsert rec;
    }
}
