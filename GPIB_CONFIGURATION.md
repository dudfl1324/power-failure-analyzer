# GPIB Multi-Device Configuration Guide

## Problem Solved
Previously, the GPIB address was hardcoded to `GPIB0::6::INSTR`, which prevented controlling multiple GPIB devices connected to the same GPIB interface. This update adds support for configurable GPIB addresses per model.

## What Changed

### 1. Configuration File (Model_Setting.json)
Added a new `MODEL_GPIB_ADDRESS` section that maps each phone model to its GPIB address:

```json
{
  "MODEL_VOLTAGE_MAP": { ... },
  "MODEL_CRITERIA": { ... },
  "MODEL_GPIB_ADDRESS": {
    "F966": 6,
    "S931": 6,
    "S936": 6,
    ...
  }
}
```

### 2. Code Changes (AutoPowerTest.ipynb)
- **load_config()**: Now loads GPIB address mappings
- **save_config()**: Persists GPIB address settings
- **measure_current_and_get_avg()**: Accepts `gpib_address` parameter
- **measure_current_and_get_avg_with_progress()**: Accepts `gpib_address` parameter and displays it in status messages
- **Model Settings Dialog**: Added GPIB address field (admin only)

## How to Use

### For Administrators
1. Login as admin
2. Click "Edit Model Settings"
3. Select a model or enter a new model name
4. Set the **GPIB Address** field (0-30)
5. Click "Save"

### Configuring Multiple Devices
If you have multiple power supplies connected via GPIB:
1. Assign different GPIB addresses to each power supply (using the power supply's front panel)
2. In the application, configure each phone model to use the appropriate GPIB address
3. For example:
   - Power Supply 1 at GPIB address 6 → Models F966, S931
   - Power Supply 2 at GPIB address 7 → Models S936, S938

### Default Behavior
- All models default to GPIB address 6 (maintaining backward compatibility)
- Valid GPIB addresses: 0-30

## Technical Details

### GPIB Resource String Format
The application now constructs GPIB resource strings dynamically:
```python
resource_str = f'GPIB0::{gpib_address}::INSTR'
```

### Example Usage
When measuring a phone model:
1. The application looks up the model's voltage: `MODEL_VOLTAGE_MAP[model]`
2. The application looks up the model's GPIB address: `MODEL_GPIB_ADDRESS[model]`
3. It connects to the power supply at that specific GPIB address
4. The status message shows: "Connecting to power supply at GPIB address {N}..."

## Benefits
- ✅ Support for multiple GPIB devices simultaneously
- ✅ Per-model GPIB address configuration
- ✅ Backward compatible (defaults to address 6)
- ✅ Admin-controlled settings prevent accidental changes
- ✅ Clear status messages showing which GPIB address is being used
