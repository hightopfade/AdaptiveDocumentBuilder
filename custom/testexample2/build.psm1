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

    $xlApp = New-Object -COMObject "Excel.Application"
    $xlApp.Visible = $false

    $fname = "$PSScriptRoot\Book1.xlsx"

    $xlBook = $xlApp.workbooks.open($fname)

    $vbaModule = $xlBook.VBProject.VBComponents.Add(1)

    $macro = BuildMacro
    $vbaModule.CodeModule.AddFromString($macro) | Out-Null

    # 52 is needed for xlsm
    # https://docs.microsoft.com/en-us/office/vba/api/Excel.XlFileFormat
    if (!$pass) {
        $xlBook.SaveAs($out, 52)
        $xlBook.Close()
        $xlApp.Quit()
    } else {
        $xlBook.SaveAs($out, 52, $pass)
        $xlBook.Close()
        $xlApp.Quit()
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlBook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Remove-Variable xlBook
    Remove-Variable xlApp

}

function BuildMacro() {

    $macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macro.txt")

    $var_counter = 1
    do {
        $this_var_name = "aVAR" + $var_counter + "a"
        $randName = RandTextAlphanumeric -len 15
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
        $randName = RandTextAlphanumeric -len 15
        if ($macro.contains($this_var_name)) {
            $macro = $macro -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    return $macro

}