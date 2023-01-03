function writelog() {
    param(
        [Parameter(mandatory=$true)]
        [string]
        $log
    )

    if (!(Test-Path "$PSScriptRoot\..\output_log.txt")) {
        Write-Host "[!] Log does not exist"
        Write-Host "[+] Creating log file now..."
        New-Item -Path "$PSScriptRoot\..\output_log.txt" -ItemType File | Out-Null
    }

    $currTime = get-date -format "dd.MMM.yyyy HH:mm:ss" 
    "$currTime $log" | out-file -FilePath "$PSScriptRoot\..\output_log.txt" -append
}