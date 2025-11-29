@IsTest
private class RollupDynamicPicklistsTest {

    @IsTest
    static void testAggregateFieldPicklistWithNullContext() {
        RollupAggregateFieldPicklist picklist =
            new RollupAggregateFieldPicklist(null);

        VisualEditor.DynamicPickListRows rows = picklist.getValues();

        System.assertNotEquals(null, rows, 'Rows should never be null.');
        // getDefaultValue() is explicitly null for this picklist
        System.assertEquals(null, picklist.getDefaultValue());
        System.assertEquals(true, picklist.isValid(null));
    }

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
}
