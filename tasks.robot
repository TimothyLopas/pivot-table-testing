*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Excel.Application


*** Variables ***
${ActiveFilePath}       pivotTableTest.xlsx


*** Tasks ***
Minimal task
    Open Application
    # This will open the workbook to the last saved worksheet.
    # In our case Sheet1
    Open Workbook    filename=${ActiveFilePath}
    # Add New Sheet    TestPivot
    @{rows}=    Create List    Buyer
    # All of the fields must be created individually
    ${amount_sum_field}=    Create Pivot Field    Amount    sum    \#,\#0
    # Once the fields are cretaed individually they should be added 
    # (in ther order necessary for the Pivot Table) to a List
    @{fields}=    Create List    ${amount_sum_field}
    # Why is the below error being thrown with every run of the Create Pivot Table keyword?
    # com_error: (-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2147024809), None)
    ${pivot_table}=    Create Pivot Table    source_worksheet=Sheet1    pivot_worksheet=TestPivot    rows=${rows}    fields=${fields}
    Sleep    10
