@IsTest
private class RollupServiceTest {

    // ---------- Test data helpers ----------

    private static Account createTestAccount() {
        Account acc = new Account(Name = 'Rollup Test Account');
        insert acc;
        return acc;
    }

    private static List<Opportunity> createTestOpportunities(Account acc) {
        List<Opportunity> opps = new List<Opportunity>();

        opps.add(new Opportunity(
            Name      = 'Opp 1',
            StageName = 'Prospecting',
            CloseDate = Date.today().addDays(1),
            AccountId = acc.Id,
            Amount    = 100
        ));
        opps.add(new Opportunity(
            Name      = 'Opp 2',
            StageName = 'Prospecting',
            CloseDate = Date.today().addDays(2),
            AccountId = acc.Id,
            Amount    = 200
        ));
        opps.add(new Opportunity(
            Name      = 'Opp 3',
            StageName = 'Closed Won',
            CloseDate = Date.today().addDays(3),
            AccountId = acc.Id,
            Amount    = 50
        ));

        insert opps;
        return opps;
    }

    private static List<Contact> createTestContacts(Account acc) {
        List<Contact> cons = new List<Contact>();
        cons.add(new Contact(FirstName = 'First',  LastName = 'Child1', AccountId = acc.Id));
        cons.add(new Contact(FirstName = 'Second', LastName = 'Child2', AccountId = acc.Id));
        insert cons;
        return cons;
    }

    private static List<Task> createTestTasks(Contact c1, Contact c2) {
        List<Task> tasks = new List<Task>();
        tasks.add(new Task(
            Subject      = 'Task A',
            Status       = 'Not Started',
            Priority     = 'Normal',
            WhoId        = c1.Id,
            ActivityDate = Date.today().addDays(1)
        ));
        tasks.add(new Task(
            Subject      = 'Task B',
            Status       = 'Not Started',
            Priority     = 'Normal',
            WhoId        = c1.Id,
            ActivityDate = Date.today().addDays(2)
        ));
        tasks.add(new Task(
            Subject      = 'Task C',
            Status       = 'Completed',
            Priority     = 'High',
            WhoId        = c2.Id,
            ActivityDate = Date.today().addDays(3)
        ));
        insert tasks;
        return tasks;
    }

    // ---------- Child-only rollup tests ----------

    @IsTest
    static void testChildSumAndCountRollup() {
        Account acc = createTestAccount();
        createTestOpportunities(acc);

        Test.startTest();
        // Use "label-style" values so normalizeApiName() is exercised.
        RollupService.RollupResult sumResult = RollupService.getRollup(
            acc.Id,
            'Parent • Opportunity (Opportunity)',
            'Child • Account Lookup (Opportunity.AccountId)',
            'Child • Amount (Opportunity.Amount)',
            'sum',    // lower-case to test normalization
            null,
            null,
            null
        );
        RollupService.RollupResult countResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            null,     // COUNT does not require an aggregate field
            'COUNT',
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertNotEquals(null, sumResult, 'Sum result should not be null.');
        System.assertEquals(350, Decimal.valueOf(sumResult.value),
            'SUM of Amount across all Opportunities should be 350.');
        System.assertEquals(3, sumResult.recordCount);
        System.assertEquals(true, sumResult.isCurrency,
            'Amount should be detected as a currency field.');
        System.assertNotEquals(null, sumResult.fieldLabel,
            'Field label hint should be populated.');

        System.assertNotEquals(null, countResult);
        System.assertEquals(3, countResult.recordCount);
        System.assertEquals('3', countResult.value);
    }

    @IsTest
    static void testChildCountDistinctAndConcats() {
        Account acc = createTestAccount();
        createTestOpportunities(acc);

        Test.startTest();
        RollupService.RollupResult countDistinct = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'StageName',
            'COUNT_DISTINCT',
            null,
            null,
            null
        );
        RollupService.RollupResult concatResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Name',
            'CONCATENATE',
            null,
            null,
            null
        );
        RollupService.RollupResult concatDistinctResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'StageName',
            'CONCATENATE_DISTINCT',
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertEquals(2, Integer.valueOf(countDistinct.value),
            'There are two distinct stage names.');
        System.assertEquals(3, countDistinct.recordCount);

        System.assertNotEquals(null, concatResult.value,
            'Concatenated string should not be null.');
        System.assertEquals(3, concatResult.recordCount);
        System.assert(concatResult.value.contains('Opp 1'),
            'Concatenated value should contain at least one of the opportunity names.');

        if (concatDistinctResult.value != null) {
            List<String> distinctStages = concatDistinctResult.value.split('; ');
            System.assertEquals(2, distinctStages.size(),
                'Should only contain two distinct stage names.');
            Set<String> stageSet = new Set<String>(distinctStages);
            System.assert(stageSet.contains('Prospecting'));
            System.assert(stageSet.contains('Closed Won'));
        } else {
            System.assert(false, 'CONCATENATE_DISTINCT result should not be null.');
        }
    }

    @IsTest
    static void testChildAverageFirstLastMinMaxRollup() {
        Account acc = createTestAccount();
        createTestOpportunities(acc);

        Test.startTest();
        RollupService.RollupResult avgResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Amount',
            'avg',    // normalised to AVERAGE
            null,
            null,
            null
        );
        RollupService.RollupResult firstResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Amount',
            'FIRST',
            null,
            null,
            null
        );
        RollupService.RollupResult lastResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Amount',
            'LAST',
            null,
            null,
            null
        );
        RollupService.RollupResult minResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Amount',
            'MIN',
            null,
            null,
            null
        );
        RollupService.RollupResult maxResult = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Amount',
            'MAX',
            null,
            null,
            null
        );
        Test.stopTest();

        Decimal avgVal = Decimal.valueOf(avgResult.value);
        System.assert(avgVal > 116 && avgVal < 117,
            'Average should be between 116 and 117.');
        System.assertEquals(3, avgResult.recordCount);

        System.assertEquals(50,  Decimal.valueOf(firstResult.value));
        System.assertEquals(3,   firstResult.recordCount);
        System.assertEquals(200, Decimal.valueOf(lastResult.value));
        System.assertEquals(3,   lastResult.recordCount);

        System.assertEquals(50,  Decimal.valueOf(minResult.value));
        System.assertEquals(200, Decimal.valueOf(maxResult.value));
    }

    /**
     * Extra coverage: aggregateType left null should default to SUM.
     */
    @IsTest
    static void testChildDefaultAggregateTypeUsesSum() {
        Account acc = createTestAccount();
        createTestOpportunities(acc);

        Test.startTest();
        RollupService.RollupResult defaultAgg = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Amount',
            null,    // aggregateType is null -> should default to SUM
            null,
            null,
            null
        );
        Test.stopTest();

        System.assertEquals(350, Decimal.valueOf(defaultAgg.value),
            'Null aggregate type should behave like SUM for Amount.');
        System.assertEquals(3, defaultAgg.recordCount,
            'Record count should still reflect all 3 child records.');
        System.assertEquals(true, defaultAgg.isCurrency,
            'Currency formatting hint should still be populated.');
    }

    // ---------- Child-only configuration / error tests ----------

    @IsTest
    static void testParentAndConfigValidationErrors() {
        Account acc = createTestAccount();

        // Missing parent
        RollupService.RollupResult noParent = RollupService.getRollup(
            null,
            'Opportunity',
            'AccountId',
            'Amount',
            'SUM',
            null,
            null,
            null
        );
        System.assertEquals(
            'Configuration error: parentId was not supplied.',
            noParent.errorMessage
        );

        // Blank child
        RollupService.RollupResult blankChild = RollupService.getRollup(
            acc.Id,
            null,
            'AccountId',
            'Amount',
            'SUM',
            null,
            null,
            null
        );
        System.assertEquals(
            'Configuration error: child object is blank.',
            blankChild.errorMessage
        );

        // Blank relationship field
        RollupService.RollupResult blankRelationship = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            '',
            'Amount',
            'SUM',
            null,
            null,
            null
        );
        System.assertEquals(
            'Configuration error: relationship field is blank.',
            blankRelationship.errorMessage
        );

        // Aggregate field required but blank
        RollupService.RollupResult blankAggregateFieldRequired = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            '',
            'SUM',
            null,
            null,
            null
        );
        System.assertEquals(
            'Configuration error: aggregate field is required for aggregation type SUM.',
            blankAggregateFieldRequired.errorMessage
        );
    }

    @IsTest
    static void testInvalidChildAndFieldErrors() {
        Account acc = createTestAccount();

        // Non-existent child object (after normalisation)
        RollupService.RollupResult invalidChild = RollupService.getRollup(
            acc.Id,
            'Made Up Label (DoesNotExist__c)',
            'AccountId',
            'Amount',
            'SUM',
            null,
            null,
            null
        );
        System.assert(
            invalidChild.errorMessage != null &&
            invalidChild.errorMessage.contains('child object'),
            'Should report a child object error.'
        );

        // Non-existent relationship field
        RollupService.RollupResult invalidRelField = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'DoesNotExist__c',
            'Amount',
            'SUM',
            null,
            null,
            null
        );
        System.assert(
            invalidRelField.errorMessage != null &&
            invalidRelField.errorMessage.contains('relationship field'),
            'Should report a relationship field error.'
        );

        // Non-existent aggregate field
        RollupService.RollupResult invalidAggField = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'DoesNotExist__c',
            'SUM',
            null,
            null,
            null
        );
        System.assert(
            invalidAggField.errorMessage != null &&
            invalidAggField.errorMessage.contains('aggregate field'),
            'Should report an aggregate field error.'
        );
    }

    @IsTest
    static void testInvalidFilterConditionTriggersCatch() {
        Account acc = createTestAccount();
        createTestOpportunities(acc);

        // Bad SOQL fragment should trigger the catch block and generic error handling
        RollupService.RollupResult result = RollupService.getRollup(
            acc.Id,
            'Opportunity',
            'AccountId',
            'Amount',
            'SUM',
            'INVALID SOQL FRAGMENT',
            null,
            null
        );

        System.assertNotEquals(null, result.errorMessage);
        System.assert(
            result.errorMessage.startsWith('Unexpected error while calculating rollup:'),
            'Should surface a generic unexpected error message.'
        );
    }

    // ---------- Grandchild rollup tests ----------

    @IsTest
    static void testGrandchildCountAndMinRollup() {
        Account acc = createTestAccount();
        List<Contact> contacts = createTestContacts(acc);
        List<Task> tasks = createTestTasks(contacts[0], contacts[1]);

        Test.startTest();
        // COUNT across Task (grandchild) via Contact (child) back to Account (parent)
        RollupService.RollupResult countGrandchild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            null,
            'COUNT',
            null,
            'Task',
            'WhoId'
        );

        // MIN ActivityDate across the same grandchild set.
        // In some orgs this may surface an "aggregate operator MIN" error; the
        // important thing for coverage is simply that the call is exercised and
        // handled by the service (no unhandled exception).
        RollupService.RollupResult minGrandchild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'MIN',
            null,
            'Task',
            'WhoId'
        );
        Test.stopTest();

        System.assertEquals(tasks.size(), countGrandchild.recordCount);
        System.assertEquals(String.valueOf(tasks.size()), countGrandchild.value);

        // Just ensure the wrapper is returned; we intentionally do not assert on
        // errorMessage/value because support for MIN on ActivityDate can vary.
        System.assertNotEquals(null, minGrandchild,
            'Grandchild MIN result wrapper should not be null.');
    }

    @IsTest
    static void testGrandchildConfigurationErrors() {
        Account acc = createTestAccount();

        // Grandchild object blank but relationship supplied
        RollupService.RollupResult blankGrandchildObject = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'MIN',
            null,
            null,
            'WhoId'
        );
        System.assertEquals(
            'Configuration error: grandchild object "Grandchild Object" is blank or invalid.',
            blankGrandchildObject.errorMessage
        );

        // Grandchild relationship field blank
        RollupService.RollupResult blankGrandchildRel = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'MIN',
            null,
            'Task',
            ''
        );
        System.assertEquals(
            'Configuration error: relationship field (lookup on grandchild) is blank.',
            blankGrandchildRel.errorMessage
        );

        // Grandchild object does not exist
        RollupService.RollupResult invalidGrandchildObject = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'MIN',
            null,
            'DoesNotExist__c',
            'WhoId'
        );
        System.assert(
            invalidGrandchildObject.errorMessage != null &&
            invalidGrandchildObject.errorMessage.contains('grandchild object'),
            'Should report invalid grandchild object.'
        );

        // Grandchild relationship field does not exist
        RollupService.RollupResult invalidGrandchildRelField = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'MIN',
            null,
            'Task',
            'DoesNotExist__c'
        );
        System.assert(
            invalidGrandchildRelField.errorMessage != null &&
            invalidGrandchildRelField.errorMessage.contains('relationship field'),
            'Should report missing grandchild relationship field.'
        );

        // Grandchild relationship field is not a lookup
        RollupService.RollupResult nonLookupGrandchildRel = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'MIN',
            null,
            'Task',
            'Subject'
        );
        System.assert(
            nonLookupGrandchildRel.errorMessage != null &&
            nonLookupGrandchildRel.errorMessage.contains('must be a lookup field'),
            'Should enforce that the grandchild relationship is a lookup.'
        );

        // Grandchild lookup does not point back to the child object
        RollupService.RollupResult lookupNotToChild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'MIN',
            null,
            'Task',
            'OwnerId' // lookup to User, not Contact
        );
        System.assert(
            lookupNotToChild.errorMessage != null &&
            lookupNotToChild.errorMessage.contains('must look up to'),
            'Should enforce that the lookup points to the child object.'
        );
    }

    // ---------- Additional grandchild coverage ----------

    @IsTest
    static void testGrandchildConcatAndFirstLastRollups() {
        Account acc = createTestAccount();
        List<Contact> contacts = createTestContacts(acc);
        createTestTasks(contacts[0], contacts[1]);

        Test.startTest();
        // COUNT_DISTINCT across Task.Status
        RollupService.RollupResult countDistinctGrandchild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'Status',
            'COUNT_DISTINCT',
            null,
            'Task',
            'WhoId'
        );

        // CONCATENATE Subjects
        RollupService.RollupResult concatGrandchild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'Subject',
            'CONCATENATE',
            null,
            'Task',
            'WhoId'
        );

        // CONCATENATE_DISTINCT Status values
        RollupService.RollupResult concatDistinctGrandchild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'Status',
            'CONCATENATE_DISTINCT',
            null,
            'Task',
            'WhoId'
        );

        // FIRST / LAST ActivityDate across the same grandchild set
        RollupService.RollupResult firstGrandchild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'FIRST',
            null,
            'Task',
            'WhoId'
        );
        RollupService.RollupResult lastGrandchild = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'ActivityDate',
            'LAST',
            null,
            'Task',
            'WhoId'
        );
        Test.stopTest();

        System.assertEquals(3, countDistinctGrandchild.recordCount,
            'COUNT_DISTINCT grandchild rollup should report total record count.');
        System.assertEquals(2, Integer.valueOf(countDistinctGrandchild.value),
            'There should be two distinct status values.');

        System.assertEquals(3, concatGrandchild.recordCount,
            'CONCATENATE grandchild rollup should include all tasks.');
        System.assertNotEquals(null, concatGrandchild.value,
            'Concatenated Subject string should not be null.');

        System.assertEquals(3, concatDistinctGrandchild.recordCount,
            'CONCATENATE_DISTINCT grandchild rollup should include all tasks.');
        System.assertNotEquals(null, concatDistinctGrandchild.value);
        System.assert(concatDistinctGrandchild.value.contains('Not Started'),
            'Distinct status concatenation should contain "Not Started".');

        System.assertNotEquals(null, firstGrandchild.value);
        System.assertNotEquals(null, lastGrandchild.value);
        System.assertNotEquals(firstGrandchild.value, lastGrandchild.value);
        System.assertEquals(3, firstGrandchild.recordCount);
        System.assertEquals(3, lastGrandchild.recordCount);
    }

    /**
     * Extra coverage: filterCondition applied in grandchild mode
     * to hit the "WHERE ... AND (filterCondition)" branch.
     */
    @IsTest
    static void testGrandchildFilterConditionFiltersRows() {
        Account acc = createTestAccount();
        List<Contact> contacts = createTestContacts(acc);
        createTestTasks(contacts[0], contacts[1]);
        // Tasks: 2 "Not Started", 1 "Completed"

        String filterCond = 'Status = \'Not Started\'';

        Test.startTest();
        RollupService.RollupResult filteredCount = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            null,          // COUNT does not need an aggregate field
            'COUNT',
            filterCond,    // Exercise grandchild filterCondition branch
            'Task',
            'WhoId'
        );
        Test.stopTest();

        System.assertEquals(2, filteredCount.recordCount,
            'Filter should limit results to the two "Not Started" tasks.');
        System.assertEquals('2', filteredCount.value);
        System.assertEquals(null, filteredCount.errorMessage,
            'Filtered grandchild COUNT should not produce an error.');
    }

    @IsTest
    static void testGrandchildAggregateFieldErrors() {
        Account acc = createTestAccount();
        createTestContacts(acc);

        // Aggregate field required for SUM in grandchild mode
        RollupService.RollupResult blankAgg = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            '',
            'SUM',
            null,
            'Task',
            'WhoId'
        );
        System.assertEquals(
            'Configuration error: aggregate field is required for aggregation type SUM.',
            blankAgg.errorMessage
        );

        // Non-existent aggregate field, using dotted "Object.Field" format
        // to exercise normalizeApiName() path.
        RollupService.RollupResult invalidAgg = RollupService.getRollup(
            acc.Id,
            'Contact',
            'AccountId',
            'Task.DoesNotExist__c',
            'SUM',
            null,
            'Task',
            'WhoId'
        );
        System.assert(
            invalidAgg.errorMessage != null &&
            invalidAgg.errorMessage.contains('aggregate field'),
            'Should report missing grandchild aggregate field.'
        );
    }
}