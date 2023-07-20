*** Settings ***
Documentation       Template robot main suite.

Library             NAKISA 
Library             sharepoint
Library             function

*** Variables ***
${run}

*** Tasks ***
R2R
    Download Input File
    ${run}=    Generate
    IF    $run == $True
        Upload Status
        Lease Recon Send Mail 
    END