# Custom Rule How-To

This guide explains how to create, test, and maintain custom discrepancy rules in IMT.

## What Custom Rules Do

Custom rules let you add simple discrepancy checks from the UI without changing Apps Script code.

A custom rule can:

- Check one required condition (Condition 1)
- Optionally check a second condition (Condition 2)
- Create a discrepancy row when both conditions pass
- Be enabled or disabled
- Be scoped to specific carriers

Custom rules are best for field-based checks.

For advanced logic (multi-step calculations, special carrier workflows), keep using built-in SYSTEM rules.

## Where To Configure Rules

1. Open the IMT web UI.
2. Go to Configuration.
3. In Live Configuration Editor, click Load Config.
4. Open the Rule Config tab.
5. Use Create Custom Rule.

## Rule Fields

When creating a custom rule, you will set:

- Rule Key: Unique ID, for example CUSTOM_CNC_WITH_SHIFT
- Rule Type: Use CUSTOM
- Enabled: YES or NO
- Carrier Scope: Optional list like HB, WERNER. Leave blank for all carriers.
- Condition 1: Required rule expression
- Condition 2: Optional second expression
- Issue Label: Discrepancy text shown in tracker
- Expected: Guidance text shown as expected value
- Actual Source: Field(s) to show in actual value
- Description: Internal note for admins

## Condition Format

Use this format:

LEFT OPERATOR RIGHT

Or for blank checks:

LEFT BLANK
LEFT NOT_BLANK

### Supported Operators

- EQUALS
- NOT_EQUALS
- CONTAINS
- NOT_CONTAINS
- STARTS_WITH
- ENDS_WITH
- BLANK
- NOT_BLANK

## Supported Tokens

Use these in LEFT or RIGHT:

- invoice.shift
- invoice.store
- invoice.tu
- invoice.baseTu
- invoice.totalCost
- invoice.number
- haulier.shift
- haulier.store
- haulier.tour
- haulier.deliveryType
- haulier.amount
- calc.invoiceTour
- calc.storeMatch
- calc.shiftMatch
- calc.tourMatch
- calc.match
- calc.type

## Actual Source Format

Actual Source controls what appears in the discrepancy Actual column.

- Single token: invoice.shift
- Multiple tokens: haulier.shift|haulier.tour

Use a pipe character to combine values.

## Example Rules

### Example 1: CNC route still has invoice shift

- Rule Key: CUSTOM_CNC_HAS_SHIFT
- Rule Type: CUSTOM
- Enabled: YES
- Carrier Scope: (blank)
- Condition 1: haulier.deliveryType EQUALS CNC
- Condition 2: invoice.shift NOT_BLANK
- Issue Label: Route recorded as Carrier Cancel
- Expected: No assigned carrier shift when delivery type is CNC
- Actual Source: invoice.shift
- Description: Custom CNC safety check

### Example 2: Tour 2 but shift mismatch marker

- Rule Key: CUSTOM_TOUR2_SHIFT_MISMATCH
- Rule Type: CUSTOM
- Enabled: YES
- Carrier Scope: HB, WERNER
- Condition 1: calc.invoiceTour EQUALS 2
- Condition 2: calc.shiftMatch EQUALS UNMATCHED
- Issue Label: Tour 2 Shift Mismatch
- Expected: Invoice and haulier shift should match
- Actual Source: invoice.shift|haulier.shift
- Description: Flags shift mismatch on second tour invoices

### Example 3: Missing store on invoice

- Rule Key: CUSTOM_STORE_MISSING
- Rule Type: CUSTOM
- Enabled: YES
- Carrier Scope: (blank)
- Condition 1: invoice.store BLANK
- Condition 2: (blank)
- Issue Label: Missing Store
- Expected: Invoice store value required
- Actual Source: invoice.store
- Description: Data quality check

## Save And Validate

1. Click Save Changes in the Rule Config panel.
2. Click Validate Config in the Configuration Validation card.
3. Review warnings or errors.

Validation warns you if:

- Rule Type is not SYSTEM or CUSTOM
- Condition text is malformed
- Custom rule has no Condition 1
- Rule keys are duplicated

## Testing Workflow

Use this process for safe rollout:

1. Create rule with Enabled = NO.
2. Save and validate config.
3. Review rule text for typos.
4. Set Enabled = YES.
5. Process a small invoice batch first.
6. Confirm expected discrepancies in Discrepancy Tracker and Invoice Results.

## Common Mistakes

- Typo in token name, for example invoice.shfit instead of invoice.shift
- Wrong operator spelling, for example NOT_EQUAL instead of NOT_EQUALS
- Forgetting Condition 1 (required)
- Reusing the same Rule Key in multiple rows
- Using custom rules for logic that should remain a system rule

## Naming Conventions

Recommended Rule Key format:

- Prefix with CUSTOM_
- Use uppercase letters and underscores
- Keep keys short and descriptive

Examples:

- CUSTOM_CNC_HAS_SHIFT
- CUSTOM_STORE_MISSING
- CUSTOM_TOUR2_SHIFT_MISMATCH

## Governance Tips

- Add clear Description text for every rule.
- Use Carrier Scope when a rule is lane-specific.
- Keep one business intent per rule.
- Prefer multiple small rules over one overloaded rule.
- Review and prune old rules monthly.

## Quick Reference

- Rule triggers when Condition 1 and Condition 2 both pass.
- Condition 2 can be blank.
- BLANK and NOT_BLANK do not need a RIGHT value.
- Actual Source can include one or many tokens separated by |.
- If Rule Config changes, save then validate before processing.
