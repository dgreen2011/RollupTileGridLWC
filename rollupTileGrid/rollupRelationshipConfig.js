import { LightningElement, api } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';

import getChildObjects from '@salesforce/apex/RollupConfigController.getChildObjects';
import getChildRelationshipFields from '@salesforce/apex/RollupConfigController.getChildRelationshipFields';
import getGrandchildObjects from '@salesforce/apex/RollupConfigController.getGrandchildObjects';
import getGrandchildRelationshipFields from '@salesforce/apex/RollupConfigController.getGrandchildRelationshipFields';
import getConfig from '@salesforce/apex/RollupConfigController.getConfig';
import saveConfig from '@salesforce/apex/RollupConfigController.saveConfig';

/**
 * rollupRelationshipConfig
 *
 * A runtime config screen that lets an admin choose:
 * - Child Object (child of parentObjectApiName)
 * - Relationship Field (lookup on child back to parent)
 * - Grandchild Object (child of the chosen child)
 * - Relationship Field (lookup on grandchild back to child)
 *
 * The options are fully dependent:
 * - Child list is based on parentObjectApiName
 * - Child relationship fields are based on selected child
 * - Grandchild objects are based on selected child
 * - Grandchild relationship fields are based on selected grandchild
 *
 * Configuration is loaded/saved via RollupConfigController.
 */
export default class RollupRelationshipConfig extends LightningElement {
    /**
     * @api parentObjectApiName
     * Example: 'Program__c' or 'Account'
     * This is the root object for the rollup tile grid.
     */
    @api parentObjectApiName;

    /**
     * @api configKey
     * A logical key to identify a config row (e.g. tile API name).
     * The Apex controller can decide whether this maps to a Custom
     * Metadata record, a custom object record, etc.
     */
    @api configKey;

    // --- Child Object state -----------------------------------

    childObjectOptions = [];
    childObjectValue;

    // Relationship field (lookup on child back to parent)
    childRelationshipFieldOptions = [];
    childRelationshipFieldValue;

    // --- Grandchild Object state ------------------------------

    grandchildObjectOptions = [];
    grandchildObjectValue;

    // Relationship field (lookup on grandchild back to child)
    grandchildRelationshipFieldOptions = [];
    grandchildRelationshipFieldValue;

    // --- UI state ---------------------------------------------

    isLoading = false;
    loadError;

    // ----------------------------------------------------------
    // Lifecycle
    // ----------------------------------------------------------

    connectedCallback() {
        if (!this.parentObjectApiName) {
            this.loadError = 'parentObjectApiName @api property is required for rollupRelationshipConfig.';
            return;
        }

        this.initialize();
    }

    async initialize() {
        this.isLoading = true;
        this.loadError = undefined;

        try {
            // 1) Load any existing persisted config (optional)
            const existing = await getConfig({
                parentObjectApiName: this.parentObjectApiName,
                configKey: this.configKey
            });

            if (existing) {
                this.childObjectValue = existing.childObjectApiName || null;
                this.childRelationshipFieldValue = existing.relationshipFieldApiName || null;
                this.grandchildObjectValue = existing.grandchildObjectApiName || null;
                this.grandchildRelationshipFieldValue =
                    existing.grandchildRelationshipFieldApiName || null;
            }

            // 2) Load base options + dependent options, in sequence
            await this.loadChildObjects();

            if (this.childObjectValue) {
                await this.loadChildRelationshipFields();
                await this.loadGrandchildObjects();

                if (this.grandchildObjectValue) {
                    await this.loadGrandchildRelationshipFields();
                }
            }
        } catch (error) {
            this.handleError(error, 'Error loading rollup relationship configuration.');
        } finally {
            this.isLoading = false;
        }
    }

    // ----------------------------------------------------------
    // Data loading helpers
    // ----------------------------------------------------------

    async loadChildObjects() {
        this.childObjectOptions = [];

        const results = await getChildObjects({
            parentObjectApiName: this.parentObjectApiName
        });

        this.childObjectOptions = (results || []).map(opt => ({
            label: opt.label,
            value: opt.value
        }));

        // Ensure the stored value is still valid
        if (
            this.childObjectValue &&
            !this.childObjectOptions.find(o => o.value === this.childObjectValue)
        ) {
            this.childObjectValue = null;
        }
    }

    async loadChildRelationshipFields() {
        this.childRelationshipFieldOptions = [];
        this.childRelationshipFieldValue = this.childRelationshipFieldValue || null;

        if (!this.childObjectValue) {
            return;
        }

        const results = await getChildRelationshipFields({
            parentObjectApiName: this.parentObjectApiName,
            childObjectApiName: this.childObjectValue
        });

        this.childRelationshipFieldOptions = (results || []).map(opt => ({
            label: opt.label,
            value: opt.value
        }));

        if (
            this.childRelationshipFieldValue &&
            !this.childRelationshipFieldOptions.find(o => o.value === this.childRelationshipFieldValue)
        ) {
            this.childRelationshipFieldValue = null;
        }
    }

    async loadGrandchildObjects() {
        this.grandchildObjectOptions = [];

        if (!this.childObjectValue) {
            this.grandchildObjectValue = null;
            return;
        }

        const results = await getGrandchildObjects({
            parentObjectApiName: this.parentObjectApiName,
            childObjectApiName: this.childObjectValue
        });

        this.grandchildObjectOptions = (results || []).map(opt => ({
            label: opt.label,
            value: opt.value
        }));

        if (
            this.grandchildObjectValue &&
            !this.grandchildObjectOptions.find(o => o.value === this.grandchildObjectValue)
        ) {
            this.grandchildObjectValue = null;
        }
    }

    async loadGrandchildRelationshipFields() {
        this.grandchildRelationshipFieldOptions = [];
        this.grandchildRelationshipFieldValue = this.grandchildRelationshipFieldValue || null;

        if (!this.childObjectValue || !this.grandchildObjectValue) {
            return;
        }

        const results = await getGrandchildRelationshipFields({
            parentObjectApiName: this.parentObjectApiName,
            childObjectApiName: this.childObjectValue,
            grandchildObjectApiName: this.grandchildObjectValue
        });

        this.grandchildRelationshipFieldOptions = (results || []).map(opt => ({
            label: opt.label,
            value: opt.value
        }));

        if (
            this.grandchildRelationshipFieldValue &&
            !this.grandchildRelationshipFieldOptions.find(
                o => o.value === this.grandchildRelationshipFieldValue
            )
        ) {
            this.grandchildRelationshipFieldValue = null;
        }
    }

    // ----------------------------------------------------------
    // Event handlers
    // ----------------------------------------------------------

    async handleChildObjectChange(event) {
        const newValue = event.detail.value || null;
        if (newValue === this.childObjectValue) {
            return;
        }

        this.childObjectValue = newValue;

        // Reset dependent selections
        this.childRelationshipFieldValue = null;
        this.childRelationshipFieldOptions = [];
        this.grandchildObjectValue = null;
        this.grandchildObjectOptions = [];
        this.grandchildRelationshipFieldValue = null;
        this.grandchildRelationshipFieldOptions = [];

        if (!this.childObjectValue) {
            return;
        }

        await this.withSpinner(async () => {
            await this.loadChildRelationshipFields();
            await this.loadGrandchildObjects();
        });
    }

    async handleChildRelationshipFieldChange(event) {
        this.childRelationshipFieldValue = event.detail.value || null;
    }

    async handleGrandchildObjectChange(event) {
        const newValue = event.detail.value || null;
        if (newValue === this.grandchildObjectValue) {
            return;
        }

        this.grandchildObjectValue = newValue;

        // Reset dependent selection
        this.grandchildRelationshipFieldValue = null;
        this.grandchildRelationshipFieldOptions = [];

        if (!this.grandchildObjectValue) {
            return;
        }

        await this.withSpinner(async () => {
            await this.loadGrandchildRelationshipFields();
        });
    }

    handleGrandchildRelationshipFieldChange(event) {
        this.grandchildRelationshipFieldValue = event.detail.value || null;
    }

    async handleSaveClick() {
        this.isLoading = true;

        try {
            await saveConfig({
                parentObjectApiName: this.parentObjectApiName,
                configKey: this.configKey,
                childObjectApiName: this.childObjectValue,
                relationshipFieldApiName: this.childRelationshipFieldValue,
                grandchildObjectApiName: this.grandchildObjectValue,
                grandchildRelationshipFieldApiName: this.grandchildRelationshipFieldValue
            });

            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Saved',
                    message: 'Rollup relationship configuration saved.',
                    variant: 'success'
                })
            );

            // Bubble a simple event so rollupTileGrid (or a parent) can react
            this.dispatchEvent(
                new CustomEvent('configsave', {
                    detail: {
                        parentObjectApiName: this.parentObjectApiName,
                        configKey: this.configKey,
                        childObjectApiName: this.childObjectValue,
                        relationshipFieldApiName: this.childRelationshipFieldValue,
                        grandchildObjectApiName: this.grandchildObjectValue,
                        grandchildRelationshipFieldApiName: this.grandchildRelationshipFieldValue
                    }
                })
            );
        } catch (error) {
            this.handleError(error, 'Error saving rollup relationship configuration.');
        } finally {
            this.isLoading = false;
        }
    }

    // ----------------------------------------------------------
    // UI helpers
    // ----------------------------------------------------------

    async withSpinner(asyncCallback) {
        this.isLoading = true;
        try {
            await asyncCallback();
        } catch (error) {
            this.handleError(error, 'Error refreshing dependent options.');
        } finally {
            this.isLoading = false;
        }
    }

    handleError(error, fallbackMessage) {
        let message = fallbackMessage;

        if (error) {
            if (Array.isArray(error.body) && error.body.length > 0) {
                message = error.body[0].message;
            } else if (error.body && error.body.message) {
                message = error.body.message;
            } else if (error.message) {
                message = error.message;
            }
        }

        this.loadError = message;

        this.dispatchEvent(
            new ShowToastEvent({
                title: 'Error',
                message,
                variant: 'error'
            })
        );
    }

    // Computed properties for template
    get isChildRelationshipDisabled() {
        return !this.childObjectValue;
    }

    get isGrandchildPicklistsDisabled() {
        return !this.childObjectValue;
    }

    get isGrandchildRelationshipDisabled() {
        return !this.grandchildObjectValue;
    }

    get hasError() {
        return !!this.loadError;
    }
}
