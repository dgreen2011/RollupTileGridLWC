import { LightningElement, api } from 'lwc';
import getRollup from '@salesforce/apex/RollupService.getRollup';
import LOCALE from '@salesforce/i18n/locale';

const MAX_ROWS = 5;
const MAX_COLUMNS = 5;
const MAX_TILES = MAX_ROWS * MAX_COLUMNS;

// Canonical aggregation types this component + RollupService understand.
const VALID_AGGREGATION_TYPES = [
    'SUM',
    'AVERAGE',
    'MAX',
    'MIN',
    'COUNT',
    'COUNT_DISTINCT',
    'CONCATENATE',
    'CONCATENATE_DISTINCT',
    'FIRST',
    'LAST'
];

// Aggregations that produce numeric values (used for number formatting).
const NUMERIC_TYPES = ['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNT', 'COUNT_DISTINCT'];

// Aggregations that imply a numeric field (for filtering by field type).
const NUMERIC_FIELD_AGG_TYPES = ['SUM', 'AVERAGE', 'MAX', 'MIN'];

// Aggregations that imply a text-like field (for filtering by field type).
const TEXT_FIELD_AGG_TYPES = ['CONCATENATE', 'CONCATENATE_DISTINCT'];

// Simple "field category" buckets used for deciding which aggregation
// types to show in the dropdown for a tile.
const FIELD_CATEGORY = {
    NUMERIC: 'numeric',
    TEXT: 'text',
    DATE: 'date',
    UNKNOWN: 'unknown'
};

// Allowed aggregation sets per field category.
const NUMERIC_FIELD_ALLOWED_AGG_TYPES = new Set([
    'SUM',
    'AVERAGE',
    'MAX',
    'MIN',
    'COUNT',
    'COUNT_DISTINCT',
    'FIRST',
    'LAST'
]);

const TEXT_FIELD_ALLOWED_AGG_TYPES = new Set([
    'CONCATENATE',
    'CONCATENATE_DISTINCT',
    'COUNT',
    'COUNT_DISTINCT',
    'FIRST',
    'LAST'
]);

const DATE_FIELD_ALLOWED_AGG_TYPES = new Set([
    'MAX',
    'MIN',
    'COUNT',
    'COUNT_DISTINCT'
]);

// Soft timeout so we never spin forever on a bad call.
const LOAD_TIMEOUT_MS = 15000;

// Base aggregation options (shared by all tiles).
const BASE_AGGREGATION_OPTIONS = [
    { label: 'Average', value: 'AVERAGE' },
    { label: 'Concatenate', value: 'CONCATENATE' },
    { label: 'Concatenate Distinct', value: 'CONCATENATE_DISTINCT' },
    { label: 'Count', value: 'COUNT' },
    { label: 'Count Distinct', value: 'COUNT_DISTINCT' },
    { label: 'Max', value: 'MAX' },
    { label: 'Min', value: 'MIN' },
    { label: 'Sum', value: 'SUM' }
];

/**
 * Normalize an aggregation type value into the canonical form used by
 * this component and RollupService (e.g., "avg" -> "AVERAGE").
 */
function normalizeAggTypeValue(raw) {
    if (raw === null || raw === undefined) {
        return null;
    }
    const upper = raw.toString().trim().toUpperCase();
    if (!upper) {
        return null;
    }
    return upper === 'AVG' ? 'AVERAGE' : upper;
}

/**
 * Infer a broad field category from an aggregation type. This is a heuristic
 * used when we don't have explicit field metadata.
 */
function inferFieldCategoryFromAggregationType(rawAggType) {
    const canonical = normalizeAggTypeValue(rawAggType);
    if (!canonical) {
        return FIELD_CATEGORY.UNKNOWN;
    }
    if (NUMERIC_FIELD_AGG_TYPES.includes(canonical)) {
        return FIELD_CATEGORY.NUMERIC;
    }
    if (TEXT_FIELD_AGG_TYPES.includes(canonical)) {
        return FIELD_CATEGORY.TEXT;
    }
    return FIELD_CATEGORY.UNKNOWN;
}

/**
 * Finalize the field category for a tile, taking into account any runtime
 * hints coming back from Apex (currency / percent / date flags).
 */
function resolveFieldCategory(tile) {
    let category =
        tile && tile.fieldCategory ? tile.fieldCategory : FIELD_CATEGORY.UNKNOWN;

    // If Apex tells us this is a date/datetime field, give it its own category
    // so we can tailor the aggregation menu.
    if (tile && tile.isDate) {
        category = FIELD_CATEGORY.DATE;
    } else if (tile && (tile.isCurrency || tile.isPercent)) {
        // Currency/percent values are always treated as numeric.
        category = FIELD_CATEGORY.NUMERIC;
    }

    return category;
}

/**
 * Return the set of allowed aggregation types for a given field category.
 * If we return null, it means "no filtering" – show all aggregation options.
 */
function getAllowedAggregationsForCategory(category) {
    if (category === FIELD_CATEGORY.NUMERIC) {
        return NUMERIC_FIELD_ALLOWED_AGG_TYPES;
    }
    if (category === FIELD_CATEGORY.TEXT) {
        return TEXT_FIELD_ALLOWED_AGG_TYPES;
    }
    if (category === FIELD_CATEGORY.DATE) {
        return DATE_FIELD_ALLOWED_AGG_TYPES;
    }
    return null;
}

// Unique id generator for each grid instance (used for cross-instance coordination).
let NEXT_INSTANCE_ID = 1;

export default class RollupTileGrid extends LightningElement {
    @api recordId;

    // Layout
    @api rows = 1;
    @api columns = 1;

    // Optional header shown above the grid
    @api headerText;
    @api headerHelpText;

    // Shared rollup config for all tiles
    @api childObjectApiName;
    @api relationshipFieldApiName;

    // NEW: optional grandchild path (Parent -> Child -> Grandchild).
    // If both are populated, the grid rolls up from the GRANDCHILD object.
    @api grandchildObjectApiName;
    @api grandchildRelationshipFieldApiName;

    @api styleVariant = 'Medium';
    @api allowUserToChangeAggregation = false; // kept for compatibility
    @api decimalPlaces = 2;

    // Refresh behavior – single Refresh button in header
    @api showRefreshButton; // default from meta.xml; treated as true if undefined

    // Control whether to show summary text under the rollup value
    @api showSummaryBelowValue; // default true if undefined

    // ---- Tile-specific @api properties (1–25) ----
    @api tile1Label;
    @api tile1AggregateFieldApiName;
    @api tile1InitialAggregationType;
    @api tile1FilterCondition;

    @api tile2Label;
    @api tile2AggregateFieldApiName;
    @api tile2InitialAggregationType;
    @api tile2FilterCondition;

    @api tile3Label;
    @api tile3AggregateFieldApiName;
    @api tile3InitialAggregationType;
    @api tile3FilterCondition;

    @api tile4Label;
    @api tile4AggregateFieldApiName;
    @api tile4InitialAggregationType;
    @api tile4FilterCondition;

    @api tile5Label;
    @api tile5AggregateFieldApiName;
    @api tile5InitialAggregationType;
    @api tile5FilterCondition;

    @api tile6Label;
    @api tile6AggregateFieldApiName;
    @api tile6InitialAggregationType;
    @api tile6FilterCondition;

    @api tile7Label;
    @api tile7AggregateFieldApiName;
    @api tile7InitialAggregationType;
    @api tile7FilterCondition;

    @api tile8Label;
    @api tile8AggregateFieldApiName;
    @api tile8InitialAggregationType;
    @api tile8FilterCondition;

    @api tile9Label;
    @api tile9AggregateFieldApiName;
    @api tile9InitialAggregationType;
    @api tile9FilterCondition;

    @api tile10Label;
    @api tile10AggregateFieldApiName;
    @api tile10InitialAggregationType;
    @api tile10FilterCondition;

    @api tile11Label;
    @api tile11AggregateFieldApiName;
    @api tile11InitialAggregationType;
    @api tile11FilterCondition;

    @api tile12Label;
    @api tile12AggregateFieldApiName;
    @api tile12InitialAggregationType;
    @api tile12FilterCondition;

    @api tile13Label;
    @api tile13AggregateFieldApiName;
    @api tile13InitialAggregationType;
    @api tile13FilterCondition;

    @api tile14Label;
    @api tile14AggregateFieldApiName;
    @api tile14InitialAggregationType;
    @api tile14FilterCondition;

    @api tile15Label;
    @api tile15AggregateFieldApiName;
    @api tile15InitialAggregationType;
    @api tile15FilterCondition;

    @api tile16Label;
    @api tile16AggregateFieldApiName;
    @api tile16InitialAggregationType;
    @api tile16FilterCondition;

    @api tile17Label;
    @api tile17AggregateFieldApiName;
    @api tile17InitialAggregationType;
    @api tile17FilterCondition;

    @api tile18Label;
    @api tile18AggregateFieldApiName;
    @api tile18InitialAggregationType;
    @api tile18FilterCondition;

    @api tile19Label;
    @api tile19AggregateFieldApiName;
    @api tile19InitialAggregationType;
    @api tile19FilterCondition;

    @api tile20Label;
    @api tile20AggregateFieldApiName;
    @api tile20InitialAggregationType;
    @api tile20FilterCondition;

    @api tile21Label;
    @api tile21AggregateFieldApiName;
    @api tile21InitialAggregationType;
    @api tile21FilterCondition;

    @api tile22Label;
    @api tile22AggregateFieldApiName;
    @api tile22InitialAggregationType;
    @api tile22FilterCondition;

    @api tile23Label;
    @api tile23AggregateFieldApiName;
    @api tile23InitialAggregationType;
    @api tile23FilterCondition;

    @api tile24Label;
    @api tile24AggregateFieldApiName;
    @api tile24InitialAggregationType;
    @api tile24FilterCondition;

    @api tile25Label;
    @api tile25AggregateFieldApiName;
    @api tile25InitialAggregationType;
    @api tile25FilterCondition;

    // Internal state: array of tile view models (config + runtime state).
    tiles = [];

    // Track whether we've kicked off the initial load.
    _initialized = false;

    // Track per-tile timeouts (not reactive).
    _tileTimeouts = {};

    // Bound window click handler (for outside-click closing).
    _windowClickHandler;

    // Bound handler for global refresh events across multiple grids.
    _globalRefreshHandler;

    // Bound handler for global "close menus" events across multiple grids.
    _globalCloseMenusHandler;

    // Unique id for this component instance (for cross-instance coordination).
    _instanceId;

    // ------------- Lifecycle -------------

    connectedCallback() {
        // Assign a unique id for this grid instance.
        if (!this._instanceId) {
            this._instanceId = `rtg_${NEXT_INSTANCE_ID++}`;
        }

        // Build the tiles from the design-time attributes.
        this.initializeTilesFromConfig();

        // Global click listener to close menus when clicking outside the component.
        if (typeof window !== 'undefined') {
            this._windowClickHandler = this.handleWindowClick.bind(this);
            window.addEventListener('click', this._windowClickHandler);

            // Listen for global refresh events so multiple grids stay in sync.
            this._globalRefreshHandler = this.handleGlobalRefresh.bind(this);
            window.addEventListener(
                'st_rollup_tile_grid_refresh',
                this._globalRefreshHandler
            );

            // Listen for global "close menus" events fired by other grids.
            this._globalCloseMenusHandler =
                this.handleGlobalCloseMenus.bind(this);
            window.addEventListener(
                'st_rollup_tile_grid_close_menus',
                this._globalCloseMenusHandler
            );
        }
    }

    renderedCallback() {
        // Only run once, and only after recordId is available.
        if (this._initialized) {
            return;
        }
        if (!this.recordId) {
            return;
        }

        this._initialized = true;

        if (!this.globalConfigError) {
            this.refreshAllTiles();
        }
    }

    disconnectedCallback() {
        // Clean up any pending timeouts.
        Object.keys(this._tileTimeouts).forEach((key) => {
            const id = this._tileTimeouts[key];
            if (id) {
                clearTimeout(id);
            }
        });
        this._tileTimeouts = {};

        // Remove global click listener.
        if (this._windowClickHandler && typeof window !== 'undefined') {
            window.removeEventListener('click', this._windowClickHandler);
            this._windowClickHandler = null;
        }

        // Remove global refresh listener.
        if (this._globalRefreshHandler && typeof window !== 'undefined') {
            window.removeEventListener(
                'st_rollup_tile_grid_refresh',
                this._globalRefreshHandler
            );
            this._globalRefreshHandler = null;
        }

        // Remove global "close menus" listener.
        if (this._globalCloseMenusHandler && typeof window !== 'undefined') {
            window.removeEventListener(
                'st_rollup_tile_grid_close_menus',
                this._globalCloseMenusHandler
            );
            this._globalCloseMenusHandler = null;
        }
    }

    // Catch unexpected errors so they don't break the entire record page.
    errorCallback(error, stack) {
        // If the configuration itself is incomplete, surface that directly so
        // the message is more descriptive than the generic "Invalid component".
        let msg = this.globalConfigError;

        if (!msg) {
            msg =
                'Unexpected error while rendering these rollup tiles. ' +
                'Please check the configuration and try again.';

            if (error) {
                if (error.body && error.body.message) {
                    msg += ' ' + error.body.message;
                } else if (error.message) {
                    msg += ' ' + error.message;
                }
            }
        }

        // eslint-disable-next-line no-console
        console.error('RollupTileGrid errorCallback', {
            error,
            stack,
            recordId: this.recordId,
            childObjectApiName: this.childObjectApiName,
            relationshipFieldApiName: this.relationshipFieldApiName,
            grandchildObjectApiName: this.grandchildObjectApiName,
            grandchildRelationshipFieldApiName: this.grandchildRelationshipFieldApiName
        });

        this.tiles = this.tiles.map((tile) =>
            this.recomputeTileDerivedFields({
                ...tile,
                isLoading: false,
                error: msg
            })
        );
    }

    // ------------- Layout helpers -------------

    get normalizedRows() {
        return this.normalizeDimension(this.rows, 1, MAX_ROWS);
    }

    get normalizedColumns() {
        return this.normalizeDimension(this.columns, 1, MAX_COLUMNS);
    }

    normalizeDimension(value, min, max) {
        const num = parseInt(value, 10);
        if (isNaN(num) || num < min) {
            return min;
        }
        if (num > max) {
            return max;
        }
        return num;
    }

    get gridStyle() {
        const cols = this.normalizedColumns;
        return `grid-template-columns: repeat(${cols}, minmax(0, 1fr));`;
    }

    // Size / style mapping (Small / Medium / Large with legacy support).
    get normalizedStyleVariant() {
        let variant =
            this.styleVariant && typeof this.styleVariant === 'string'
                ? this.styleVariant.toLowerCase()
                : 'medium';

        // Map legacy values to new names
        if (variant === 'compact') {
            variant = 'small';
        } else if (variant === 'square') {
            variant = 'large';
        }

        if (variant !== 'small' && variant !== 'medium' && variant !== 'large') {
            variant = 'medium';
        }

        return variant;
    }

    get tileContainerClass() {
        let base = 'st-rollup-tile';
        const variant = this.normalizedStyleVariant;

        if (variant === 'medium') {
            base += ' st-rollup-tile_large';
        } else if (variant === 'large') {
            base += ' st-rollup-tile_square';
        } else {
            // "Small" (or anything unknown) -> compact style
            base += ' st-rollup-tile_compact';
        }
        return base;
    }

    // ------------- Header helpers -------------

    get hasHeader() {
        const text = this.headerText ? this.headerText.trim() : '';
        const help = this.headerHelpText ? this.headerHelpText.trim() : '';
        return !!(text || help);
    }

    get hasHeaderOrToolbar() {
        const showRefresh =
            this.showRefreshButton === undefined || this.showRefreshButton === null
                ? true
                : this.showRefreshButton;
        return this.hasHeader || showRefresh;
    }

    get showRefreshButtonEffective() {
        // If the admin never touches the property, treat it as true (default).
        return (
            this.showRefreshButton === true ||
            this.showRefreshButton === 'true' ||
            this.showRefreshButton === undefined ||
            this.showRefreshButton === null
        );
    }

    get allowUserToChangeAggregationEffective() {
        return (
            this.allowUserToChangeAggregation === true ||
            this.allowUserToChangeAggregation === 'true' ||
            this.allowUserToChangeAggregation === undefined ||
            this.allowUserToChangeAggregation === null
        );
    }

    // Effective flag for showing summary text under values.
    get showSummaryBelowValueEffective() {
        return (
            this.showSummaryBelowValue === true ||
            this.showSummaryBelowValue === 'true' ||
            this.showSummaryBelowValue === undefined ||
            this.showSummaryBelowValue === null
        );
    }

    // ------------- Grandchild-mode helpers -------------

    /**
     * True when the admin has configured both grandchild properties.
     * When true, queries roll up from the grandchild object instead of the child.
     */
    get isGrandchildMode() {
        const obj =
            this.grandchildObjectApiName &&
            this.grandchildObjectApiName.toString().trim();
        const rel =
            this.grandchildRelationshipFieldApiName &&
            this.grandchildRelationshipFieldApiName.toString().trim();
        return !!(obj && rel);
    }

    /**
     * Object API name to use when building summary labels:
     *  - Child object in normal mode
     *  - Grandchild object when grandchild mode is enabled
     */
    get aggregateObjectApiNameForLabel() {
        if (this.isGrandchildMode && this.grandchildObjectApiName) {
            return this.grandchildObjectApiName;
        }
        return this.childObjectApiName;
    }

    // ------------- Config error (shared across tiles) -------------

    get globalConfigError() {
        const missing = [];

        if (!this.childObjectApiName) {
            missing.push('Child Object');
        }
        if (!this.relationshipFieldApiName) {
            missing.push('Relationship Field (lookup on child)');
        }

        // Only require grandchild properties if the admin has started to configure them.
        const hasAnyGrandchildConfig =
            (this.grandchildObjectApiName &&
                this.grandchildObjectApiName.toString().trim()) ||
            (this.grandchildRelationshipFieldApiName &&
                this.grandchildRelationshipFieldApiName.toString().trim());

        if (hasAnyGrandchildConfig) {
            if (!this.grandchildObjectApiName) {
                missing.push('Grandchild Object');
            }
            if (!this.grandchildRelationshipFieldApiName) {
                missing.push('Relationship Field (lookup on grandchild)');
            }
        }

        if (!missing.length) {
            return null;
        }

        return (
            'Rollup Tile Grid is not fully configured yet. Missing: ' +
            missing.join(', ') +
            '. In the Lightning App Builder, set these properties (and at minimum configure "Tile 1 Aggregate Field") before using this component.'
        );
    }

    /**
     * Derive a human-readable singular object label from the object we’re
     * actually aggregating over:
     *  - Child object in normal mode
     *  - Grandchild object in grandchild mode
     */
    get childObjectLabelSingular() {
        const apiName = this.aggregateObjectApiNameForLabel;
        if (!apiName || typeof apiName !== 'string') {
            return 'record';
        }

        let name = apiName.trim();

        // Remove common suffixes
        name = name.replace(/__c$/i, '').replace(/__x$/i, '');

        // Strip namespace prefix if present (ns__Object)
        const parts = name.split('__');
        if (parts.length > 1) {
            name = parts[1];
        }

        // Convert from PascalCase / camelCase / underscores to nice words
        name = name
            .replace(/_/g, ' ')
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .trim();

        if (!name) {
            return 'record';
        }

        name = name
            .split(' ')
            .map((w) =>
                w ? w[0].toUpperCase() + w.slice(1).toLowerCase() : ''
            )
            .join(' ')
            .trim();

        return name || 'record';
    }

    // ------------- Tile initialization -------------

    initializeTilesFromConfig() {
        const rows = this.normalizedRows;
        const cols = this.normalizedColumns;
        const maxSlots = Math.min(rows * cols, MAX_TILES);
        const tiles = [];

        for (let index = 1; index <= maxSlots; index++) {
            tiles.push(this.buildInitialTileConfig(index));
        }

        this.tiles = tiles;
    }

    buildInitialTileConfig(index) {
        const suffix = index.toString();

        const label = this[`tile${suffix}Label`];
        const aggregateFieldApiName = this[`tile${suffix}AggregateFieldApiName`];
        const rawAggregationType = this[`tile${suffix}InitialAggregationType`];
        const filterCondition = this[`tile${suffix}FilterCondition`];

        const initialAggregationType =
            this.normalizeAggregationType(rawAggregationType);
        const fieldCategory =
            inferFieldCategoryFromAggregationType(initialAggregationType);

        // Use a single global decimal setting for all tiles.
        const decimalPlaces =
            this.decimalPlaces !== null && this.decimalPlaces !== undefined
                ? this.decimalPlaces
                : 2;

        const baseTile = {
            index,
            label: label || `Tile ${index}`,
            aggregateFieldApiName,
            initialAggregationType,
            filterCondition,
            decimalPlaces,
            fieldCategory,

            // runtime state
            aggregateType: initialAggregationType,
            isLoading: false,
            error: null,
            value: null,
            recordCount: null,
            isCurrency: false,
            isPercent: false,
            isDate: false,
            fieldLabel: null,
            isAggregationMenuOpen: false,

            // derived view fields (filled by recomputeTileDerivedFields)
            displayValue: '-',
            hasRecordCount: false,
            summaryRecordLabel: null,
            summaryLabel: null,
            aggregationMenuOptions: [],
            gearMenuClass: ''
        };

        return this.recomputeTileDerivedFields(baseTile);
    }

    normalizeAggregationType(raw) {
        const canonical = normalizeAggTypeValue(raw);

        if (!canonical) {
            return 'SUM';
        }

        if (VALID_AGGREGATION_TYPES.includes(canonical)) {
            return canonical;
        }

        // If App Builder somehow stored an unexpected value, be defensive
        // and fall back to SUM so the SOQL stays valid.
        return 'SUM';
    }

    // ------------- Tile view-model helpers -------------

    recomputeTileDerivedFields(tile) {
        const aggregateType =
            normalizeAggTypeValue(
                tile.aggregateType || tile.initialAggregationType || 'SUM'
            ) || 'SUM';

        const isDateAggregate =
            tile &&
            tile.isDate &&
            (aggregateType === 'MIN' || aggregateType === 'MAX');

        // Numeric aggregate?
        const isNumericAggregate =
            NUMERIC_TYPES.includes(aggregateType) && !isDateAggregate;

        // Decide which "field category" this tile belongs to for dropdown filtering.
        const fieldCategory = resolveFieldCategory(tile);
        const allowedSet = getAllowedAggregationsForCategory(fieldCategory);

        // displayValue
        let displayValue;
        if (tile.value === null || tile.value === undefined || tile.value === '') {
            displayValue = '-';
        } else if (isNumericAggregate) {
            const num = parseFloat(tile.value);
            if (isNaN(num)) {
                displayValue = tile.value;
            } else {
                const fractionDigits =
                    tile.decimalPlaces !== undefined && tile.decimalPlaces !== null
                        ? parseInt(tile.decimalPlaces, 10)
                        : 2;

                try {
                    const formatted = new Intl.NumberFormat(LOCALE, {
                        minimumFractionDigits: 0,
                        maximumFractionDigits: fractionDigits
                    }).format(num);

                    if (tile.isCurrency) {
                        displayValue = '$' + formatted;
                    } else if (tile.isPercent) {
                        displayValue = formatted + '%';
                    } else {
                        displayValue = formatted;
                    }
                } catch (_e) {
                    displayValue = tile.value;
                }
            }
        } else {
            // For text-like and date aggregates, use the string as-is.
            displayValue = tile.value;
        }

        // hasRecordCount + summaryRecordLabel (now includes object name)
        let hasRecordCount = false;
        let summaryRecordLabel = null;

        if (tile.recordCount !== undefined && tile.recordCount !== null) {
            hasRecordCount = true;
            const count = Number(tile.recordCount);
            const objectLabel = this.childObjectLabelSingular || 'record';

            if (Number.isNaN(count)) {
                const recordWord = 'records';
                summaryRecordLabel = `${tile.recordCount} ${objectLabel} ${recordWord}`;
            } else {
                let formattedCount;
                try {
                    formattedCount = new Intl.NumberFormat(LOCALE).format(count);
                } catch (_e) {
                    formattedCount = count.toString();
                }

                const recordWord = count === 1 ? 'record' : 'records';
                summaryRecordLabel = `${formattedCount} ${objectLabel} ${recordWord}`;
            }
        }

        // fieldLabelForSummary
        let fieldLabelForSummary = '';
        const serverLabel =
            tile.fieldLabel && typeof tile.fieldLabel === 'string'
                ? tile.fieldLabel.trim()
                : '';
        if (serverLabel) {
            fieldLabelForSummary = serverLabel;
        } else {
            const explicitLabel =
                tile.label && typeof tile.label === 'string'
                    ? tile.label.trim()
                    : '';
            if (explicitLabel) {
                fieldLabelForSummary = explicitLabel;
            } else {
                const apiName =
                    tile.aggregateFieldApiName &&
                    typeof tile.aggregateFieldApiName === 'string'
                        ? tile.aggregateFieldApiName.trim()
                        : '';
                fieldLabelForSummary = apiName || 'this field';
            }
        }

        // friendlyAggregationLabel
        let friendlyAggregationLabel;
        switch (aggregateType) {
            case 'AVERAGE':
            case 'AVG':
                friendlyAggregationLabel = 'Average';
                break;
            case 'COUNT':
                friendlyAggregationLabel = 'Count';
                break;
            case 'COUNT_DISTINCT':
                friendlyAggregationLabel = 'Distinct count';
                break;
            case 'MAX':
                friendlyAggregationLabel = 'Maximum';
                break;
            case 'MIN':
                friendlyAggregationLabel = 'Minimum';
                break;
            case 'FIRST':
                friendlyAggregationLabel = 'First value';
                break;
            case 'LAST':
                friendlyAggregationLabel = 'Last value';
                break;
            case 'CONCATENATE':
                friendlyAggregationLabel = 'Combined values';
                break;
            case 'CONCATENATE_DISTINCT':
                friendlyAggregationLabel = 'Unique combined values';
                break;
            default:
                friendlyAggregationLabel = 'Sum';
                break;
        }

        // summaryLabel (tooltip + text under value)
        let summaryLabel;
        if (summaryRecordLabel) {
            // Example: "Sum of 'Aerial Footage' across 2 Project records"
            summaryLabel = `${friendlyAggregationLabel} of '${fieldLabelForSummary}' across ${summaryRecordLabel}`;
        } else {
            // No record count available
            summaryLabel = `${friendlyAggregationLabel} of '${fieldLabelForSummary}'`;
        }

        // aggregationMenuOptions (gear dropdown)
        const aggregationMenuOptions = BASE_AGGREGATION_OPTIONS
            .filter((opt) => !allowedSet || allowedSet.has(opt.value))
            .map((opt) => {
                const isSelected = opt.value === aggregateType;
                return {
                    label: opt.label,
                    value: opt.value,
                    isSelected,
                    itemClass:
                        'slds-dropdown__item' +
                        (isSelected ? ' slds-is-selected' : ''),
                    ariaChecked: isSelected ? 'true' : 'false'
                };
            });

        // gearMenuClass
        let gearMenuClass =
            'st-rollup-gear-menu slds-dropdown-trigger slds-dropdown-trigger_click';
        if (tile.isAggregationMenuOpen) {
            gearMenuClass += ' slds-is-open';
        }

        return {
            ...tile,
            aggregateType,
            fieldCategory,
            displayValue,
            hasRecordCount,
            summaryRecordLabel,
            summaryLabel,
            aggregationMenuOptions,
            gearMenuClass
        };
    }

    // ------------- Refresh / loading -------------

    handleRefreshAllClick() {
        // Clicking Refresh on one grid should refresh all grids on the same page.
        if (typeof window !== 'undefined') {
            const evt = new CustomEvent('st_rollup_tile_grid_refresh', {
                bubbles: false,
                composed: false,
                detail: {
                    sourceRecordId: this.recordId || null
                }
            });
            window.dispatchEvent(evt);
        } else {
            // Fallback – just refresh this grid.
            this.refreshAllTiles();
        }
    }

    refreshAllTiles() {
        this.tiles.forEach((tile) => {
            this.loadTile(tile.index);
        });
    }

    async loadTile(index) {
        const globalConfigError = this.globalConfigError;
        if (globalConfigError) {
            // If the shared config is bad, set an error on this tile and bail.
            this.tiles = this.tiles.map((tile) =>
                tile.index === index
                    ? this.recomputeTileDerivedFields({
                          ...tile,
                          isLoading: false,
                          error: globalConfigError,
                          value: null,
                          recordCount: null,
                          isCurrency: false,
                          isPercent: false,
                          isDate: false,
                          fieldLabel: null
                      })
                    : tile
            );
            return;
        }

        const currentTile = this.tiles.find((t) => t.index === index);
        if (!currentTile) {
            return;
        }

        // If this tile doesn't have an aggregate field configured, surface a
        // clear configuration message instead of letting Apex error out.
        if (!currentTile.aggregateFieldApiName) {
            const configMsg = `Tile ${index} is not fully configured. Set "Tile ${index} Aggregate Field" in the Lightning App Builder.`;
            this.tiles = this.tiles.map((tile) =>
                tile.index === index
                    ? this.recomputeTileDerivedFields({
                          ...tile,
                          isLoading: false,
                          error: configMsg,
                          value: null,
                          recordCount: null,
                          isCurrency: false,
                          isPercent: false,
                          isDate: false,
                          fieldLabel: null
                      })
                    : tile
            );
            return;
        }

        // Clear any previous timeout for this tile.
        const existingTimeout = this._tileTimeouts[index];
        if (existingTimeout) {
            clearTimeout(existingTimeout);
            delete this._tileTimeouts[index];
        }

        // Reset tile state to "loading"
        this.tiles = this.tiles.map((tile) =>
            tile.index === index
                ? this.recomputeTileDerivedFields({
                      ...tile,
                      isLoading: true,
                      error: null,
                      value: null,
                      recordCount: null,
                      isCurrency: false,
                      isPercent: false,
                      isDate: false,
                      fieldLabel: null
                  })
                : tile
        );

        const timeoutError = new Error(
            'Timed out while loading rollup. Please refresh the page or contact your admin.'
        );

        const timeoutPromise = new Promise((_, reject) => {
            const id = setTimeout(() => {
                reject(timeoutError);
            }, LOAD_TIMEOUT_MS);
            this._tileTimeouts[index] = id;
        });

        try {
            const tileAfterReset = this.tiles.find((t) => t.index === index);
            if (!tileAfterReset) {
                return;
            }

            let aggregateTypeToUse =
                tileAfterReset.aggregateType ||
                tileAfterReset.initialAggregationType ||
                'SUM';

            // NOTE: we now pass 'COUNT' straight through to Apex so it can
            // execute a true COUNT() branch instead of behaving like SUM.

            const apexPromise = getRollup({
                parentId: this.recordId,
                childObjectApiName: this.childObjectApiName,
                relationshipFieldApiName: this.relationshipFieldApiName,
                aggregateFieldApiName: tileAfterReset.aggregateFieldApiName,
                aggregateType: aggregateTypeToUse,
                filterCondition: tileAfterReset.filterCondition,
                grandchildObjectApiName: this.isGrandchildMode
                    ? this.grandchildObjectApiName
                    : null,
                grandchildRelationshipFieldApiName: this.isGrandchildMode
                    ? this.grandchildRelationshipFieldApiName
                    : null
            });

            const data = await Promise.race([apexPromise, timeoutPromise]);

            if (!data) {
                // No data returned
                this.tiles = this.tiles.map((tile) =>
                    tile.index === index
                        ? this.recomputeTileDerivedFields({
                              ...tile,
                              isLoading: false,
                              error: 'No data was returned for this rollup.',
                              value: null,
                              recordCount: null,
                              isCurrency: false,
                              isPercent: false,
                              isDate: false,
                              fieldLabel: null
                          })
                        : tile
                );
                return;
            }

            let errorMsg = null;
            let value = null;
            let recordCount = null;
            let isCurrency = false;
            let isPercent = false;
            let isDate = false;
            let fieldLabel = null;

            if (data.errorMessage) {
                // Business / configuration error from Apex – show it in the tile.
                errorMsg = data.errorMessage;
                value = undefined;
                recordCount = data.recordCount;
                isCurrency = !!data.isCurrency;
                isPercent = !!data.isPercent;
                isDate = !!data.isDate;
                fieldLabel = data.fieldLabel;
            } else {
                errorMsg = null;
                value = data.value;
                recordCount = data.recordCount;
                isCurrency = !!data.isCurrency;
                isPercent = !!data.isPercent;
                isDate = !!data.isDate;
                fieldLabel = data.fieldLabel;
            }

            this.tiles = this.tiles.map((tile) =>
                tile.index === index
                    ? this.recomputeTileDerivedFields({
                          ...tile,
                          isLoading: false,
                          error: errorMsg,
                          value,
                          recordCount,
                          isCurrency,
                          isPercent,
                          isDate,
                          fieldLabel
                      })
                    : tile
            );
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error('RollupTileGrid loadTile error', {
                error,
                index,
                recordId: this.recordId,
                childObjectApiName: this.childObjectApiName,
                relationshipFieldApiName: this.relationshipFieldApiName,
                grandchildObjectApiName: this.grandchildObjectApiName,
                grandchildRelationshipFieldApiName: this.grandchildRelationshipFieldApiName
            });

            let msg = 'Unexpected error while loading rollup.';
            if (error) {
                if (error.body && error.body.message) {
                    msg = error.body.message;
                } else if (error.message) {
                    msg = error.message;
                } else if (typeof error === 'string') {
                    msg = error;
                }
            }

            this.tiles = this.tiles.map((tile) =>
                tile.index === index
                    ? this.recomputeTileDerivedFields({
                          ...tile,
                          isLoading: false,
                          error: msg,
                          value: null,
                          recordCount: null,
                          isCurrency: false,
                          isPercent: false,
                          isDate: false,
                          fieldLabel: null
                      })
                    : tile
            );
        } finally {
            const timeoutId = this._tileTimeouts[index];
            if (timeoutId) {
                clearTimeout(timeoutId);
                delete this._tileTimeouts[index];
            }
        }
    }

    // ------------- Dropdown / click handling -------------

    /**
     * Root click handler for clicks *inside* the component.
     * If you click anywhere that's not inside a gear menu,
     * close any open aggregation menus.
     */
    handleRootClick() {
        this.closeAllAggregationMenus();
    }

    /**
     * Global window click handler, to close menus when clicking completely
     * outside the component.
     */
    handleWindowClick(event) {
        if (!this.template) {
            return;
        }

        // If click is inside this component, let handleRootClick deal with it.
        if (this.template.contains(event.target)) {
            return;
        }

        this.closeAllAggregationMenus();
    }

    /**
     * Handler for global refresh events so that clicking Refresh on one
     * grid refreshes all grids on the same record page.
     */
    handleGlobalRefresh(event) {
        const detail = event && event.detail ? event.detail : {};
        const sourceRecordId = detail.sourceRecordId;

        // If the event specifies a recordId, and we also have one, require them to match.
        if (sourceRecordId && this.recordId && sourceRecordId !== this.recordId) {
            return;
        }

        this.refreshAllTiles();
    }

    /**
     * Handler for global "close menus" events so that opening a gear dropdown
     * in one grid closes all open dropdowns in other grids on the page.
     */
    handleGlobalCloseMenus(event) {
        if (!event || !event.detail) {
            return;
        }

        const { instanceId } = event.detail;

        // Ignore events we originated ourselves.
        if (!instanceId || instanceId === this._instanceId) {
            return;
        }

        this.closeAllAggregationMenus();
    }

    /**
     * Close all aggregation menus (used by outside-click + root click + global events).
     */
    closeAllAggregationMenus() {
        let anyOpen = false;
        const updated = this.tiles.map((tile) => {
            if (tile.isAggregationMenuOpen) {
                anyOpen = true;
                const next = {
                    ...tile,
                    isAggregationMenuOpen: false
                };
                return this.recomputeTileDerivedFields(next);
            }
            return tile;
        });

        if (anyOpen) {
            this.tiles = updated;
        }
    }

    // ------------- UI handlers for per-tile controls -------------

    handleGearClick(event) {
        event.stopPropagation();
        const index = Number(event.currentTarget.dataset.index);
        if (!index) {
            return;
        }

        // Before toggling this tile's menu, tell all *other* grids on the page
        // to close any open aggregation menus.
        if (typeof window !== 'undefined') {
            const closeEvt = new CustomEvent('st_rollup_tile_grid_close_menus', {
                bubbles: false,
                composed: false,
                detail: {
                    instanceId: this._instanceId,
                    tileIndex: index
                }
            });
            window.dispatchEvent(closeEvt);
        }

        this.tiles = this.tiles.map((tile) => {
            if (tile.index === index) {
                // Toggle this tile's menu.
                const next = {
                    ...tile,
                    isAggregationMenuOpen: !tile.isAggregationMenuOpen
                };
                return this.recomputeTileDerivedFields(next);
            }

            // Close all other menus when opening a new one.
            if (tile.isAggregationMenuOpen) {
                const next = {
                    ...tile,
                    isAggregationMenuOpen: false
                };
                return this.recomputeTileDerivedFields(next);
            }

            return tile;
        });
    }

    handleAggregationMenuClick(event) {
        event.preventDefault();
        event.stopPropagation();

        const index = Number(event.currentTarget.dataset.index);
        const newTypeRaw = event.currentTarget.dataset.value;

        if (!index) {
            return;
        }

        const newTypeUpper = (newTypeRaw || '').toString().toUpperCase();
        if (!newTypeUpper) {
            // Just close the menu.
            this.tiles = this.tiles.map((tile) =>
                tile.index === index
                    ? this.recomputeTileDerivedFields({
                          ...tile,
                          isAggregationMenuOpen: false
                      })
                    : tile
            );
            return;
        }

        this.tiles = this.tiles.map((tile) => {
            if (tile.index !== index) {
                return tile;
            }

            const normalizedType = this.normalizeAggregationType(newTypeUpper);
            const next = this.recomputeTileDerivedFields({
                ...tile,
                aggregateType: normalizedType,
                isAggregationMenuOpen: false
            });
            return next;
        });

        // After updating the aggregation type, reload the tile.
        this.loadTile(index);
    }
}
