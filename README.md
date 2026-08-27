# corneal-report-extractor
Python tool for extracting structured ophthalmic and corneal measurement data from PDF reports into Excel and CSV files.

## Overview

This project extracts clinical measurement data from corneal topography / corneal mapping PDF reports and converts the information into a structured format suitable for analysis and review.

The extractor is designed to process measurements for both:

- **OD** — Right eye
- **OS** — Left eye

The extracted information can be exported to:

- Excel (`.xlsx`)
- CSV (`.csv`)

The program also provides validation and error reporting to identify missing or incomplete fields.

---

## Features

### Patient Information

Extracts available patient and examination information such as:

- Patient ID
- Examination date
- Examination time
- Device information
- Other report metadata

### Sim-K Measurements

The Sim-K section is extracted separately for OD and OS.

For each eye, the program extracts:

- K1 power
- K1 axis
- K2 power
- K2 axis
- CYL power
- CYL axis

The standardized Sim-K fields are:

```text
OD_SimK_K1_D
OD_SimK_K1_axis_deg
OD_SimK_K2_D
OD_SimK_K2_axis_deg
OD_SimK_CYL_D
OD_SimK_CYL_axis_deg

OS_SimK_K1_D
OS_SimK_K1_axis_deg
OS_SimK_K2_D
OS_SimK_K2_axis_deg
OS_SimK_CYL_D
OS_SimK_CYL_axis_deg
