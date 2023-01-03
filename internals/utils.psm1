function GenerateRandomAuthor() {
    $firstNames = Get-Content $PSScriptRoot\lists\first-names.txt
    $middleNames = Get-Content $PSScriptRoot\lists\middle-names.txt
    $lastNames = Get-Content $PSScriptRoot\lists\last-names.txt

    $num = Get-Random -Minimum 0 -Maximum $firstNames.Length
    $fname = $firstNames[$num]

    $num = Get-Random -Minimum 0 -Maximum $lastNames.Length
    $lname = $lastNames[$num]

    $authorName = "$fname $lname"
    $initials = $fname[0], $lname[0] -join ""
    return $authorName, $initials.ToLower()
}

function SetAuthor() {
    $authorName, $initials = GenerateRandomAuthor

    $regPath = "HKCU:\SOFTWARE\Microsoft\Office\Common\UserInfo"
    New-ItemProperty -Path $regPath -Name Username -Value $authorName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name UserInitials -Value $initials -PropertyType String -Force | Out-Null
}

function GenerateDocumentName() {
    $word1 = @('Final', 'Outstanding', 'Next', 'Selected', 'Partial', 'Needed', 'Sales', 'Past_Due', 'Overdue', 'Paid', 'Incorrect', 'New', 'Your')
    $word2 = @("Invoice", "Bill", "Ticket", "Payment", "Receipt", "Invoices")

    $num = Get-Random -Minimum 0 -Maximum $word1.Length
    $num1 = Get-Random -Minimum 0 -Maximum $word2.Length
    $num2 = Get-Random -Minimum 100 -Maximum 10000
    $name = $word1[$num], $word2[$num1], $num2 -join "_"
    return $name
}

function RandTextAlphaUpper() {
    param(
        [Parameter(mandatory=$true)]
        [int]
        $len
    )

    return -join ((65..90) | Get-Random -Count $len | % {[char]$_})
}

function RandTextAlphaLower() {
    param(
        [Parameter(mandatory=$true)]
        [int]
        $len
    )

    return -join ((97..122) | Get-Random -Count $len | % {[char]$_})
}

function RandTextAlpha() {
    param(
        [Parameter(mandatory=$true)]
        [int]
        $len
    )

    return -join ((65..90) + (97..122) | Get-Random -Count $len | % {[char]$_})
}

function RandTextAlphanumeric() {
    param(
        [Parameter(mandatory=$true)]
        [int]
        $len
    )

    return -join ((65..90) + (97..122) + (48..57) | Get-Random -Count $len | % {[char]$_})
}

function RenameVBAModule() {
    param(
        [Parameter(mandatory=$true)]
        [System.Object]
        $officeObject,

        [Parameter(mandatory=$true)]
        [string]
        $from,

        [Parameter(mandatory=$true)]
        [string]
        $to
    )

    $rename = $officeObject.VBProject.VBComponents.Item($from)
    $rename.Name = $to
}

function GenerateRandomDateTime() {
    <#

    The $format parameter will decide how to output the date/time

    PS C:\Users\ebfe\Desktop\adb> GenerateRandomDateTime -format "dd.MMM.yyyy HH:mm:ss"

    29.Oct.2010 19:38:20
    PS C:\Users\ebfe\Desktop\adb> GenerateRandomDateTime -format "dd.MMM.yyyy"

    17.Jul.2007
    PS C:\Users\ebfe\Desktop\adb> GenerateRandomDateTime -format "dd-MMM-yyyy"

    10-Jan-2010
    PS C:\Users\ebfe\Desktop\adb> GenerateRandomDateTime -format "HH:mm:ss"

    #>

    param(
        [Parameter(mandatory=$false)]
        [string]
        $format
    )

    $hours = Get-Random -Minimum 0 -Maximum 12
    $days = Get-Random -Minimum 0 -Maximum 30
    $months = Get-Random -Minimum 0 -Maximum 12
    $years = Get-Random -Minimum 0 -Maximum 15

    $date = (Get-Date).addHours($hours).addDays($days).addMonths($months).addYears(-$years)
    if ($format) {
        $date = $date.ToString($format)
    }
    return $date
}

function GenerateJunkWord() {
    # param(
    #     [Parameter(mandatory=$true)]
    #     [System.Object]
    #     $docObject,
    #
    #     [Parameter(mandatory=$true)]
    #     [System.Object]
    #     $officeObject,
    #
    #     [Parameter(mandatory=$true)]
    #     [string]
    #     $type,
    #
    #     [Parameter(mandatory=$false)]
    #     [String]
    #     $pass,
    #
    #     [Parameter(mandatory=$true)]
    #     [string]
    #     $out
    # )

    #need args here for sentence structure, 3x3 for 3 word 3 deep
    #need random function for random length
    $words = 6
    $sentence = ""

    $wordlist = Get-Content $PSScriptRoot\lists\words.txt
    $word = $wordlist[$(Get-Random -Minimum 0 -Maximum $wordlist.Length)]
    Write-Host $word
    #Build sentence
    # $sentence = ""
    # $count = 5
    # Do {
    #   $sentence = $sentence + $word
    #   $count--
    # } While ($count -gt 0)
}
