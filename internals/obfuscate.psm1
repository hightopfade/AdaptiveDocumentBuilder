$workdir = "$PSScriptRoot\donut"
$payloaddir = "$PSScriptRoot\donut\payload"

#Create donut directory and install donut,yasm,golink tools
if (!(Test-Path $workdir)) {
    Write-Host "[!] Donut folder does not exist"
    Write-Host "[+] Creating donut folder now..."
    Write-Host "[+] Installing tools..."
    New-Item -Path "$PSScriptRoot\donut" -ItemType Directory | Out-Null
    New-Item -Path "$PSScriptRoot\donut\payload" -ItemType Directory | Out-Null
    $ProgressPreference = 'Continue'
    Invoke-WebRequest "https://github.com/TheWover/donut/releases/download/v0.9.3/donut_v0.9.3.zip" -OutFile $workdir\donut.zip
    Invoke-WebRequest "https://github.com/yasm/yasm/releases/download/v1.3.0/yasm-1.3.0-win64.exe" -OutFile $workdir\yasm.exe
    Invoke-WebRequest "http://www.godevtool.com/Golink.zip" -OutFile $workdir\golink.zip
    Expand-Archive -Path $workdir\donut.zip -DestinationPath $workdir
    Expand-Archive -Path $workdir\golink.zip $workdir
    Write-Host "[+] Tool installation complete"
    Write-Host "[+] To run: Obfuscate -binary C:\PATH\TO\BINARY.EXE"
    }

#Create payload directory
if (!(Test-Path $payloaddir)) {
    New-Item -Path "$PSScriptRoot\donut\payload" -ItemType Directory | Out-Null
    }
function Obfuscate {
    param(
        [Parameter(mandatory=$true)]
        [string]
        $binary,

        [Parameter(mandatory=$false)]
        [int]
        $OSarch = 64
    )

Start-Process -NoNewWindow -FilePath "$workdir\donut.exe" -ArgumentList "-f 1 $binary -o $payloaddir\temp.bin" | Out-Null
Start-Sleep -Seconds 1.0
#Extra spaced in heredoc below MUST be maintained and not trimmed
$SCAsmBuilder = @"
Global Start
SECTION 'foo' write, execute,read
Start:
incbin "temp.bin"  
"@
$SCAsmBuilder | Set-Content $payloaddir\shellcode.asm
Start-Process -NoNewWindow -FilePath "$workdir\yasm.exe" -ArgumentList "-f win$OSarch -o $payloaddir\shellcode.obj $payloaddir\shellcode.asm" | Out-Null
Start-Sleep -Seconds 1.0
Start-Process -NoNewWindow -FilePath "$workdir\golink.exe" -ArgumentList "/ni /entry Start $payloaddir\shellcode.obj" | Out-Null
Start-Sleep -Seconds 1.0
Rename-Item -Path "$payloaddir\shellcode.exe" -NewName "$payloaddir\obfuscated.exe"
Write-Host "[+] Complete! $binary has been obfuscated and renamed to $payloaddir\obfuscated.exe"
}
