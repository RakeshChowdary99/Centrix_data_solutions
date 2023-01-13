*** Settings ***
Documentation       Template robot main suite.

Library    RPA.Desktop
Library    RPA.Desktop.Windows
Library    RPA.Excel.Files
Library    RPA.Tables
Library    RPA.Images
Library    RPA.Browser.Selenium    auto_close=${False}
Library    RPA.Windows
Library    String
Library    RPA.Dialogs
Library    XML
Library    RPA.JSON
            

*** Tasks ***

Centrix_application
    ${config_values}=    reading config
    TRY
        ${table}=    reading data    ${config_values}[input_path]
        ${length}=    Get Length    ${table}
        IF    '${length}' != '1'
            opening Application    ${config_values}[application_path]
            ${logi}=    log in to Centrix_application          
            ${count}=    set variable    0         
            FOR    ${element}    IN    @{table}
                Log    ${element}
                ${or}=    navigating to orders window
                IF    ${or}
                    ${ref_num}=    getting reference number    ${element}
                        ${ref_num}=    Convert To String    ${ref_num}
                        Log    ${element}[Reference Number]
                        ${row}=    Convert To String    ${count}
                        Set Table Cell    ${table}    ${row}    Reference Number    ${ref_num}
                        ${count}=    Evaluate    ${count}+1
                    
                ELSE
                    Add text    unable to navigate to order
                    ${dai}=    Show dialog    
                    Wait dialog    ${dai}
                    BREAK
                END
            END
        Write table to CSV    ${table}    result.csv
        ELSE
            Add text    No data available to process
            ${dai}=    Show dialog    
            Wait dialog    ${dai}
        END
        
    
    EXCEPT    message
        Log    unable to process 
    END


*** Keywords ***
opening Application
    [Arguments]    ${path}
    RPA.Desktop.Open Application    ${path}
log in to Centrix_application
    TRY
        Sleep    7s
        Set Value    id:txt_StaffNumber    BP
        Set Value    id:txt_Password    password
        RPA.Windows.Click    id:btn_Login
        Sleep    10s
        RPA.Windows.Click    name:ORDERS
        RETURN     1
    EXCEPT    message
        RETURN    0
    
    END
navigating to orders window
    TRY
        Sleep    10s
        Set Value    id:txt_OrdersMain_Option    1
        RPA.Windows.Click    name:Go
        RETURN    1
    EXCEPT
        RETURN    0
    END
reading data
    [Arguments]    ${in_path}
    TRY
        ${table}=    Read table from CSV    ${in_path}
        RETURN    ${table}
    EXCEPT
        RETURN    0
    END
getting reference number
    [Arguments]    ${ele}
    TRY
        Set Value    id:txt_OrdersNew_ProductCode    ${ele}[Product Code]
        Set Value    id:txt_OrdersNew_UnitPrice    ${ele}[Unit Price]
        Set Value    id:txt_OrdersNew_CustomerAccount    ${ele}[Customer Acct Number]
        IF    '${ele}[Priority Order]' == 'TRUE' 
            RPA.Windows.Click    id:chk_OrdersNew_Priority
        END
        select    id:cmb_OrdersNew_Quantity    ${ele}[Quantity]
        RPA.Windows.Click    name:Submit
        Sleep    5s
        ${data}=    RPA.Windows.Get Text    id:65535
        Log    ${data}
        RPA.Windows.Click    name:OK
        ${re}=    Split String    ${data}    :
        RETURN    ${re}[1]
    EXCEPT
        RETURN    0
    END      
reading config
    
    ${config_details}=    Parse Xml    source=cofig.xml     strip_namespaces=False
    ${application_path}=    XML.Get Element     ${config_details}    applications_path
    ${input_path}=    XML.Get Element     ${config_details}    input_path
    ${output_path}=    XML.Get Element     ${config_details}    output_path
    ${details}=    Create Dictionary
                ...    application_path=${application_path.text}
                ...    input_path=${input_path.text}
                ...    output_path=${output_path.text}
    RETURN    ${details}
