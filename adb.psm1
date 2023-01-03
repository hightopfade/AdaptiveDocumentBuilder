Import-Module $PSScriptRoot\internals\utils.psm1
Import-Module $PSScriptRoot\internals\logger.psm1
Import-Module $PSScriptRoot\internals\tools.psm1

function ListAdversaries {
    Write-Host "[+] Adversary Templates"
    $adversaries = Get-Childitem "$PSScriptRoot\adversaries" -directory | Select-Object Name
    foreach ($adversary in $adversaries) {
        Write-Host $adversary.Name
    }

    Write-Host "`r`n[+] Custom/One Off Templates"
    $custom = Get-Childitem "$PSScriptRoot\custom" -directory | Select-Object Name
    foreach ($itr in $custom) {
        Write-Host $itr.Name
    }
}
function CreateDocument {
    param(
        [Parameter(mandatory=$false)]
        [int]
        $count,

        [Parameter(mandatory=$false)]
        [string]
        $out,

        [Parameter(mandatory=$false)]
        [string]
        $adversary,

        [Parameter(mandatory=$false)]
        [string]
        $custom,

        [Parameter(mandatory=$false)]
        [String]
        $pass
    )

    $itr = 1
    if ($count -eq 0) { $count = 1 } 

    Write-Host "[*] Creating '$count' document(s)..."
    writelog -log "Starting job"
    writelog -log "Adversary: '$adversary'"
    writelog -log "Number of documents: '$count'"

    Do {
        
        $name = GenerateDocumentName

        if ($out) {
            if (Test-Path $out) {} else {
                Write-Host "[*] Specified output directory does NOT exist"
                Write-Host "[+] Creating '$out' now..."
                New-Item -ItemType Directory -Force -Path $out >$null 2>&1
            }
            $outPath = "$out"
        } else {
            $outPath = "$PSScriptRoot"
        }

        if ($adversary) {
            if (Test-Path "$PSScriptRoot\adversaries\$adversary") {
                Import-Module $PSScriptRoot\adversaries\$adversary\build.psm1
            } else {
                Write-Host "[-] Adversary template does not exist..."
                Write-Host "[+] Run 'ListAdversaries' to see what is available."
                writelog -log "Adversary '$adversary' does NOT exist..."
                break
            }
        } elseif ($custom) {
            if (Test-Path "$PSScriptRoot\custom\$custom") {
                Import-Module $PSScriptRoot\custom\$custom\build.psm1
            } else {
                Write-Host "[-] Custom/One Off template does not exist..."
                Write-Host "[+] Run 'ListAdversaries' to see what is available."
                writelog -log "Adversary '$adversary' does NOT exist..."
                break
            }
        }

        if ($pass) {
            Build -out "$outPath\$name" -pass $pass
            writelog -log "Password: $pass"
        } else {
            Build -out "$outPath\$name"
        }

        Write-Host "[*] Created '$outPath\$name'..."
        $itr++
    } While ($itr -le $count)

    writelog -log "=============================="

}