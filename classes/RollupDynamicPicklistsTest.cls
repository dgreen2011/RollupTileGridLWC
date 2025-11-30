@IsTest
private class RollupDynamicPicklistsTest {

    // Helper to build a context for a specific parent object
    private static VisualEditor.DesignTimePageContext makeContext(String entityApiName) {
        VisualEditor.DesignTimePageContext ctx = new VisualEditor.DesignTimePageContext();
        ctx.entityName = entityApiName;
        return ctx;
    }

    // ------------------------------------------------------------
    // Aggregate field picklist – basic smoke test (child/grandchild path)
    // ------------------------------------------------------------
    @IsTest
    static void testAggregateFieldPicklistWithValidContext() {
        VisualEditor.DesignTimePageContext ctx = makeContext('Account');

        RollupAggregateFieldPicklist picklist =
            new RollupAggregateFieldPicklist(ctx);

        Test.startTest();
        VisualEditor.DynamicPickListRows rows = picklist.getValues();
        Test.stopTest();

        System.assertNotEquals(null, rows, 'Rows should never be null.');

        // This picklist intentionally has no default value
        System.assertEquals(null, picklist.getDefaultValue(),
            'Default should always be null for aggregate field picklist.');
        System.assertEquals(true, picklist.isValid(null));
    }

    // ------------------------------------------------------------
    // Aggregate field picklist – fallback path (all objects, null context)
    // ------------------------------------------------------------
    @IsTest
    static void testAggregateFieldPicklistFallbackWithNullContext() {
        // No context means getValues() should fall back to scanning all objects
        // via addFieldsFromAllObjects(), which is currently uncovered.
        RollupAggregateFieldPicklist picklist =
            new RollupAggregateFieldPicklist(null);

        Test.startTest();
        VisualEditor.DynamicPickListRows rows = picklist.getValues();
        Test.stopTest();

        System.assertNotEquals(
            null,
            rows,
            'Rows should not be null even when context is null.'
        );

        // There should be at least one aggregate-friendly field in the org
        System.assert(
            rows.size() > 0,
            'Fallback scan of all objects should surface at least one aggregate-friendly field.'
        );

        System.assertEquals(true, picklist.isValid(null),
            'isValid should always return true.');
    }

    // ------------------------------------------------------------
    // Child object picklist
    // ------------------------------------------------------------
    @IsTest
    static void testChildObjectPicklistWithNullContext() {
        RollupChildObjectPicklist picklist =
            new RollupChildObjectPicklist(null);

        VisualEditor.DynamicPickListRows rows = picklist.getValues();

        System.assertNotEquals(null, rows,
            'Rows collection itself should not be null.');
        System.assertEquals(0, rows.size(),
            'With a null context there should be no child values.');
        System.assertEquals(null, picklist.getDefaultValue(),
            'Default value should be null when there are no rows.');
        System.assertEquals(true, picklist.isValid(null));
    }

    @IsTest
    static void testChildObjectPicklistWithValidContext() {
        VisualEditor.DesignTimePageContext ctx = makeContext('Account');

        RollupChildObjectPicklist picklist =
            new RollupChildObjectPicklist(ctx);

        Test.startTest();
        VisualEditor.DynamicPickListRows rows = picklist.getValues();
        Test.stopTest();

        System.assertNotEquals(null, rows,
            'Rows should not be null for a valid context.');

        VisualEditor.DataRow defaultRow = picklist.getDefaultValue();
        if (rows.size() > 0) {
            System.assertNotEquals(null, defaultRow,
                'Default row should be populated when rows exist.');
        } else {
            System.assertEquals(null, defaultRow,
                'If there are no rows, default should be null.');
        }
        System.assertEquals(true, picklist.isValid(null));
    }

    // ------------------------------------------------------------
    // Grandchild object picklist
    // ------------------------------------------------------------
    @IsTest
    static void testGrandchildObjectPicklistWithNullContext() {
        RollupGrandchildObjectPicklist picklist =
            new RollupGrandchildObjectPicklist(null);

        VisualEditor.DynamicPickListRows rows = picklist.getValues();

        System.assertNotEquals(null, rows);
        System.assertEquals(0, rows.size(),
            'No grandchild objects should be returned with null context.');
        System.assertEquals(null, picklist.getDefaultValue(),
            'Default is explicitly null.');
        System.assertEquals(true, picklist.isValid(null));
    }

    @IsTest
    static void testGrandchildObjectPicklistWithValidContext() {
        VisualEditor.DesignTimePageContext ctx = makeContext('Account');

        RollupGrandchildObjectPicklist picklist =
            new RollupGrandchildObjectPicklist(ctx);

        Test.startTest();
        VisualEditor.DynamicPickListRows rows = picklist.getValues();
        Test.stopTest();

        System.assertNotEquals(null, rows,
            'Grandchild object picklist should return a non-null row collection.');
        System.assertEquals(true, picklist.isValid(null));
    }

    // ------------------------------------------------------------
    // Grandchild relationship field picklist
    // ------------------------------------------------------------
    @IsTest
    static void testGrandchildRelFieldPicklistWithNullContext() {
        RollupGrandchildRelFieldPicklist picklist =
            new RollupGrandchildRelFieldPicklist(null);

        VisualEditor.DynamicPickListRows rows = picklist.getValues();

        System.assertNotEquals(null, rows);
        System.assertEquals(0, rows.size(),
            'No relationship fields should be returned with null context.');
        System.assertEquals(null, picklist.getDefaultValue(),
            'Default is explicitly null.');
        System.assertEquals(true, picklist.isValid(null));
    }

    @IsTest
    static void testGrandchildRelFieldPicklistWithValidContext() {
        VisualEditor.DesignTimePageContext ctx = makeContext('Account');

        RollupGrandchildRelFieldPicklist picklist =
            new RollupGrandchildRelFieldPicklist(ctx);

        Test.startTest();
        VisualEditor.DynamicPickListRows rows = picklist.getValues();
        Test.stopTest();

        System.assertNotEquals(null, rows,
            'Grandchild relationship picklist should return a non-null row collection.');
        System.assertEquals(true, picklist.isValid(null));
    }

    // ------------------------------------------------------------
    // Relationship field picklist (child lookup back to parent)
    // ------------------------------------------------------------
    @IsTest
    static void testRelationshipFieldPicklistWithNullContext() {
        RollupRelationshipFieldPicklist picklist =
            new RollupRelationshipFieldPicklist(null);

        VisualEditor.DynamicPickListRows rows = picklist.getValues();

        System.assertNotEquals(null, rows);
        System.assertEquals(0, rows.size(),
            'No relationship fields should be returned with null context.');
        System.assertEquals(null, picklist.getDefaultValue(),
            'Default is explicitly null.');
        System.assertEquals(true, picklist.isValid(null));
    }

    @IsTest
    static void testRelationshipFieldPicklistWithValidContext() {
        VisualEditor.DesignTimePageContext ctx = makeContext('Account');

        RollupRelationshipFieldPicklist picklist =
            new RollupRelationshipFieldPicklist(ctx);

        Test.startTest();
        VisualEditor.DynamicPickListRows rows = picklist.getValues();
        Test.stopTest();

        System.assertNotEquals(null, rows,
            'Relationship field picklist should return a non-null row collection.');
        System.assertEquals(true, picklist.isValid(null));
    }
}