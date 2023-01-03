Import-Module $PSScriptRoot\utils.psm1

function CreateFreshDocument() {
    param(
        [Parameter(mandatory=$true)]
        [string]
        $out
    )

    SetAuthor

    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false
    $wdDoc = $wdApp.Documents.Add()
    $wdDoc.SaveAs("$out\Doc1.docm", 13)
    $wdDoc.Close()
    $wdApp.Quit()
}

function CreateFreshWorkbook() {
    param(
        [Parameter(mandatory=$true)]
        [string]
        $out
    )

    SetAuthor

    $xlApp = New-Object -COMObject "Excel.Application"
    $xlApp.Visible = $false
    $xlBook = $xlApp.Workbooks.Add()
    $xlBook.SaveAs("$out\Book1.xlsm", 52)
    $xlBook.Close()
    $xlApp.Quit()

}

function EditMetaData() {

    <#

    To use this with a JSON file, use the 'example.json' found within the
    'internals' directory

    Ensure that the '$in' parameter takes in a FULL PATH

    #>

    param(
        [Parameter(mandatory=$true)]
        [string]
        $in,

        [Parameter(mandatory=$false)]
        [string]
        $json
    )

    $fname = $in.split("\\")[-1]
    $fnameNoExtension = $fname.split(".")[0]
    $ext = $fname.split(".")[1]

    # really hacky
    # doing it this way because powershell arrays are created
    # of a fixed size
    $arr = $in.split("\\")
    $cnt = 0
    do {
        $path += $arr[$cnt]
        $path += "\\"
        $cnt++
    } while ($cnt -lt ($arr.Length-1))

    Copy-Item -LiteralPath $in -Destination "$path$fnameNoExtension.zip"
    Expand-Archive -LiteralPath "$path$fnameNoExtension.zip" -DestinationPath "$path$fnameNoExtension"
    Remove-Item -Path "$path$fnameNoExtension.zip"

    [xml]$coreProperties = Get-Content -Path "$path$fnameNoExtension\docProps\core.xml"

    if ($json) {
        $jsn = Get-Content -Path $json | ConvertFrom-Json
        $coreProperties.coreProperties.title = $jsn.title
        $coreProperties.coreProperties.subject = $jsn.subject
        $coreProperties.coreProperties.creator = $jsn.creator
        $coreProperties.coreProperties.keywords = $jsn.keywords
        $coreProperties.coreProperties.description = $jsn.description
        $coreProperties.coreProperties.lastModifiedBy = $jsn.lastModifiedBy
        $coreProperties.coreProperties.created.innertext = $jsn.created
        $coreProperties.coreProperties.modified.innertext = $jsn.modified
    } else {
        # modify the author
        $author, $initials = GenerateRandomAuthor
        $coreProperties.coreProperties.creator = $author

        # modify the 'Last Modified By' name
        $author, $initials = GenerateRandomAuthor
        $coreProperties.coreProperties.lastModifiedBy = $author

        # modify creation time
        $date = GenerateRandomDateTime -format "yyyy-MM-ddTHH:mm:ssZ"
        $coreProperties.coreProperties.created.innertext = $date

        # modify 'Last Modified By' time
        $date = GenerateRandomDateTime -format "yyyy-MM-ddTHH:mm:ssZ"
        $coreProperties.coreProperties.modified.innertext = $date

        # modify Comments field
        $coreProperties.coreProperties.description = ""

    }

    $coreProperties.Save("$path$fnameNoExtension\docProps\core.xml")

    Compress-Archive -LiteralPath "$path$fnameNoExtension\[Content_Types].xml",
        "$path$fnameNoExtension\word\",
        "$path$fnameNoExtension\docProps\",
        "$path$fnameNoExtension\_rels\" -DestinationPath "$path$fnameNoExtension.zip"

    Remove-Item -Path "$path$fnameNoExtension" -Recurse
    Rename-Item -Path "$path$fnameNoExtension.zip" -NewName $path$fnameNoExtension"_modified.$ext"


}

function CleanUp() {
    param(
        [Parameter(mandatory=$false)]
        [Switch] $all
    )

    $word = Get-Process "WINWORD" -ErrorAction SilentlyContinue
    $excel = Get-Process "EXCEL" -ErrorAction SilentlyContinue
    if ($word) {
        $word | Stop-Process -Force
        Remove-Variable word
    }

    if ($excel) {
        $excel | Stop-Process -Force
        Remove-Variable excel
    }

    if ($all){
        Write-Host "test"
        Get-ChildItem -Path "$PSScriptRoot\..\" -Include *.doc*,*.xls* | ForEach-Object {
            Write-Host $_.FullName
        }
    }

    Get-ChildItem -Path "$PSScriptRoot\..\" -Include *.doc*,*.xls* -Recurse -Force | ForEach-Object {
        if ($_.Name.StartsWith("~")) {
            Remove-Item $_.FullName -Force
        }
    }
    Remove-Module adb
}

function Get-DocumentProperties() {
    param(
        [Parameter(mandatory=$true)]
        [string]
        $file
    )

    <#
    Currently supporting word, need to make it more generic to support both word and excel

    potentially a quick loop to determine ofApp and ofDoc based on the file extension
    
    #>

    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $true
    $wdDoc = $wdApp.Documents.open("$file")
    $binding = "System.Reflection.BindingFlags" -as [type];
    Write-Host "[+] Dumping BuiltInDocumentProperties"
    $properties = $wdDoc.BuiltInDocumentProperties
    foreach($property in $properties)
    {
        $pn = [System.__ComObject].invokemember("name", $binding::GetProperty, $null, $property, $null)
            trap [system.exception]
            {
                write-host -foreground blue "Value not found for $pn"
                continue
            }
        "$pn`: " +
        [System.__ComObject].invokemember("value", $binding::GetProperty, $null, $property, $null)
    }

    $properties = $wdDoc.CustomDocumentProperties
    If ($properties.Length -gt 0) {
        Write-Host "`r`n[+] Dumping CustomDocumentProperties"
        foreach($property in $properties)
        {
            $pn = [System.__ComObject].invokemember("name", $binding::GetProperty, $null, $property, $null)
                trap [system.exception]
                {
                    write-host -foreground blue "Value not found for $pn"
                    continue
                }
            "$pn`: " +
            [System.__ComObject].invokemember("value", $binding::GetProperty, $null, $property, $null)
        }
    }


    $wdDoc.Close($false)
    $wdApp.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdDoc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Remove-Variable wdDoc
    Remove-Variable wdApp

}