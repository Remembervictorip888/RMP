# RMP Excel VBA Processing Engine - Helper Functions Structure

## Overview

The RMP Excel VBA Processing Engine uses a modular architecture where individual helper functions perform specific data processing tasks. The main processing flow calls these helpers sequentially through the [ExecuteBatchHelpers](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1496-L1504) function.

## Main Processing Flow

1. [ProcessSingleBatch](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1425-L1482) - Processes a single batch of data
2. [ExecuteBatchHelpers](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1496-L1504) - Calls [RunAllHelpers](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1507-L1547) to execute all helper functions
3. [RunAllHelpers](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1507-L1547) - Executes all helper functions in a predefined sequence
4. [ExecuteHelper](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1559-L1570) - Calls individual helper functions using `Application.Run`

## Helper Functions Execution

The helper functions are stored in an array within [RunAllHelpers](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1507-L1547):

```vb
Dim helpers As Variant
helpers = Array("Helper_01_SubGroupMapping", "Helper_02_Source", "Helper_03_NB_and_RI", _
               "Helper_04_DefaultRate", "Helper_05_LTVPropValue", "Helper_06_AssumptionValuation", _
               "Helper_07_Policies_Date", "Helper_08_ScenarioMerge", "Helper_09_ReinsurerMapping", _
               "Helper_10_AddGMR_MTHLY", "Helper_11_AddMaturity_Term", "Helper_13_Projection_Expand", _
               "Helper_14_AddRemaining_Term", "Helper_15_YrMth_Indicator", "Helper_16_DefaultPattern", _
               "Helper_17_ClaimRate", "Helper_18_CalculateOMCR", "Helper_20_OMV_and_PctOMV_Combined", _
               "Helper_21_CPR_SMM_Combined", "Helper_23_AcquisitionExpense", "Helper_24_PolicyForceAndOPB_Combined", _
               "Helper_24A_Duration_Plus_and_Min", "Helper_24B_Next_Policy_In_Force", "Helper_24C_PreviousValues", _
               "Helper_25_FixedAssumedSeverity", "Helper_26_InflationFactor", "Helper_27_MaintenanceExpense", _
               "Helper_28_Commission_and_Commission_Recovery", "Helper_29_RI_Policies_Count", "Helper_30_RI_Premium", _
               "Helper_31_RiskInForce_and_DefaultClaimOutgo", "Helper_32_RI_Collateral", "Helper_33_RI_NPR")
```

Each helper function is called using `Application.Run procName, ws` in the [ExecuteHelper](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1559-L1570) subroutine.

## Where Are the Helper Implementations?

Based on the code structure, the individual helper functions (e.g., [Helper_01_SubGroupMapping](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/Old_Version/RMP_Excel_VBA_Code.bas#L1835-L1856), [Helper_02_Source](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/Old_Version/RMP_Excel_VBA_Code.bas#L1858-L1882), etc.) are likely implemented in one of the following locations:

1. **Embedded in the Excel workbook (.xlsm file)** - The most likely location since VBA modules can be embedded in Excel workbooks
2. **In separate .bas files** - Possibly in the main .bas file or other module files not examined
3. **In class modules** - May be implemented as methods in class modules

## How to Locate Helper Implementations

To find the actual implementation of the helper functions:

1. Open the `RMP_Excel_VBA_V1.xlsm` file in Excel
2. Press `Alt + F11` to open the VBA editor
3. In the Project Explorer, look for modules containing the helper functions
4. Search for procedures with names matching those in the helpers array

## Adding New Helpers

To add a new helper function:

1. Create a new subroutine with a unique name following the pattern `Helper_##_Description`
2. Add the function name to the helpers array in [RunAllHelpers](file:///c%3A/Users/Vip_Minibook/OneDrive/HKMC_BUDGET_ENHANCEMENT/MIP/Excel_VBA/RMP/RMP_Excel_VBA_Code_OPTIMIZED.bas#L1507-L1547)
3. Ensure the function accepts a Worksheet parameter: `Sub Helper_##_Description(ws As Worksheet)`

## Best Practices for Helper Functions

1. Follow the template provided in the comments section of the main module
2. Include proper error handling with meaningful log messages
3. Use the available utility functions (NzLong, NzDouble, NzString, etc.)
4. Validate input data before processing
5. Log the start and completion of each helper
6. Use appropriate data types (Currency for monetary values, LongLong for large counts on 64-bit)