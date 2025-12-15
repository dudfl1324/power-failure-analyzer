# power-failure-analyzer

## Features
- Automated power failure testing for mobile devices
- Support for multiple device models with configurable settings
- **Multi-GPIB device support** - Control multiple power supplies via different GPIB addresses
- Excel logging and export functionality
- User authentication with admin privileges

## GPIB Multi-Device Support
The application now supports controlling multiple GPIB devices simultaneously. Each phone model can be configured with its own GPIB address.

📖 **See [GPIB_CONFIGURATION.md](GPIB_CONFIGURATION.md) for detailed setup instructions**

## Quick Start
1. Run the Jupyter notebook `AutoPowerTest.ipynb`
2. Login (admin/worker accounts available)
3. Configure model settings (admin only)
4. Run measurements on your devices

## Configuration
Model settings including voltage, GPIB address, and pass/fail criteria are stored in `Model_Setting.json` or can be edited through the admin interface.
