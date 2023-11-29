*** Settings ***
Documentation    Download and send Payslip
Library    BuiltIn
Library    String
Library    pyautogui
Library    XML
Library    OperatingSystem
Library    RPA.Windows
Library    RPA.Desktop
Library    RPA.HTTP
Library    PayslipEmailBlast



*** Variables ***

*** Keywords ***

Running_Download
    run_download

Converting_to_PDF
    convert_sheetPdf    ExcelFile

*** Test Cases ***

Run Python
    Running_Download

Test Function
    Converting_to_PDF

    
    
