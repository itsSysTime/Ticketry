param(
    [string]$edition
)

# === Edition Mapping ===
$editions = @{
    "Enterprise" = @{
        Pfn = "Microsoft.Windows.4.X19-99683_8wekyb3d8bbwe"
        Key = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
    }

    "Enterprise N" = @{
        Pfn = "Microsoft.Windows.27.X19-98746_8wekyb3d8bbwe"
        Key = "3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT"
    }

    "Professional" = @{
        Pfn = "Microsoft.Windows.48.X19-98841_8wekyb3d8bbwe"
        Key = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"
    }

    "Professional N" = @{
        Pfn = "Microsoft.Windows.49.X19-98859_8wekyb3d8bbwe"
        Key = "2B87N-8KFHP-DKV6R-Y2C8J-PKCKT"
    }

    "Home N" = @{
        Pfn = "Microsoft.Windows.98.X19-98877_8wekyb3d8bbwe"
        Key = "4CPRK-NM3K3-X6XXQ-RXX86-WXCHW"
    }

    "Home Country Specific" = @{
        Pfn = "Microsoft.Windows.99.X19-99652_8wekyb3d8bbwe"
        Key = "N2434-X9D7W-8PF6X-8DV9T-8TYMD"
    }

    "Home Single Language" = @{
        Pfn = "Microsoft.Windows.100.X19-99661_8wekyb3d8bbwe"
        Key = "BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"
    }

    "Home" = @{
        Pfn = "Microsoft.Windows.101.X19-98868_8wekyb3d8bbwe"
        Key = "YTMG3-N6DKC-DKB77-7M9GH-8HVX7"
    }

    "PPI Pro" = @{
        Pfn = "Microsoft.Windows.119.X19-99606_8wekyb3d8bbwe"
        Key = "XKCNC-J26Q9-KFHD2-FKTHY-KD72Y"
    }

    "Education" = @{
        Pfn = "Microsoft.Windows.121.X19-98886_8wekyb3d8bbwe"
        Key = "YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"
    }

    "Education N" = @{
        Pfn = "Microsoft.Windows.122.X19-98892_8wekyb3d8bbwe"
        Key = "84NGF-MHBT6-FXBX8-QWJK7-DRR8H"
    }

    "Enterprise LTSC 2019" = @{
        Pfn = "Microsoft.Windows.125.X21-83233_8wekyb3d8bbwe"
        Key = "43TBQ-NH92J-XKTM7-KT3KK-P39PB"
    }

    "Enterprise LTSC 2016" = @{
        Pfn = "Microsoft.Windows.125.X21-05035_8wekyb3d8bbwe"
        Key = "NK96Y-D9CD8-W44CQ-R8YTK-DYJWX"
    }

    "Enterprise LTSC 2015" = @{
        Pfn = "Microsoft.Windows.125.X19-99617_8wekyb3d8bbwe"
        Key = "FWN7H-PF93Q-4GGP8-M8RF3-MDWWW"
    }

    "Enterprise N LTSC 2019" = @{
        Pfn = "Microsoft.Windows.126.X21-83264_8wekyb3d8bbwe"
        Key = "M33WV-NHY3C-R7FPM-BQGPT-239PG"
    }

    "Enterprise N LTSC 2016" = @{
        Pfn = "Microsoft.Windows.126.X21-04921_8wekyb3d8bbwe"
        Key = "2DBW3-N2PJG-MVHW3-G7TDK-9HKR4"
    }

    "Enterprise N LTSC 2015" = @{
        Pfn = "Microsoft.Windows.126.X19-98770_8wekyb3d8bbwe"
        Key = "NTX6B-BRYC2-K6786-F6MVQ-M7V2X"
    }

    "Professional for Workstations" = @{
        Pfn = "Microsoft.Windows.161.X21-43626_8wekyb3d8bbwe"
        Key = "DXG7C-N36C4-C4HTG-X4T3X-2YV77"
    }

    "Professional for Workstations N" = @{
        Pfn = "Microsoft.Windows.162.X21-43644_8wekyb3d8bbwe"
        Key = "WYPNQ-8C467-V2W6J-TX4WX-WT2RQ"
    }

    "Professional Education" = @{
        Pfn = "Microsoft.Windows.164.X21-04955_8wekyb3d8bbwe"
        Key = "8PTT6-RNW4C-6V7J2-C2D3X-MHBPB"
    }

    "Professional Education N" = @{
        Pfn = "Microsoft.Windows.165.X21-04956_8wekyb3d8bbwe"
        Key = "GJTYN-HDMQY-FRR76-HVGC7-QPF8P"
    }

    "Enterprise G" = @{
        Pfn = "Microsoft.Windows.171.X21-24727_8wekyb3d8bbwe"
        Key = "FV469-WGNG4-YQP66-2B2HY-KD8YX"
    }

    "Enterprise G N" = @{
        Pfn = "Microsoft.Windows.172.X21-24709_8wekyb3d8bbwe"
        Key = "FW7NV-4T673-HF4VX-9X4MM-B4H4T"
    }

    "Cloud/S" = @{
        Pfn = "Microsoft.Windows.178.X21-32983_8wekyb3d8bbwe"
        Key = "V3WVW-N2PV2-CGWC3-34QGF-VMJ2C"
    }

    "Cloud/S N" = @{
        Pfn = "Microsoft.Windows.179.X21-32987_8wekyb3d8bbwe"
        Key = "NH9J3-68WK7-6FB93-4K3DF-DJ4F6"
    }

    "Cloud/S E" = @{
        Pfn = "Microsoft.Windows.183.X21-76403_8wekyb3d8bbwe"
        Key = "2HN6V-HGTM8-6C97C-RK67V-JQPFD"
    }

    "IoT Enterprise" = @{
        Pfn = "Microsoft.Windows.188.X21-99378_8wekyb3d8bbwe"
        Key = "XQQYW-NFFMW-XJPBH-K8732-CKFFD"
    }

    "IoT Enterprise LTSC 2021" = @{
        Pfn = "Microsoft.Windows.191.X21-99682_8wekyb3d8bbwe"
        Key = "QPM6N-7J2WJ-P88HH-P3YRH-YY74H"
    }

    "CloudEdition/SE N" = @{
        Pfn = "Microsoft.Windows.202.X22-53884_8wekyb3d8bbwe"
        Key = "K9VKN-3BGWV-Y624W-MCRMQ-BHDCD"
    }

    "CloudEdition/SE" = @{
        Pfn = "Microsoft.Windows.203.X22-53847_8wekyb3d8bbwe"
        Key = "KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W"
    }

    "IoT Enterprise LTSC Subscription" = @{
        Pfn = "Microsoft.Windows.205.X23-15027_8wekyb3d8bbwe"
        Key = "J7NJW-V6KBM-CC8RW-Y29Y4-HQ2MJ"
    }
    
    "IoT Enterprise LTSC 2024" = @{
        Pfn = "Microsoft.Windows.191.X23-12617_8wekyb3d8bbwe"
        Key = "CGK42-GYN6Y-VD22B-BX98W-J8JXD"
    }
}

$aliases = @{
    "edu"  = "Education"
    "edun" = "Education N"
    "pro"  = "Professional"
    "pron" = "Professional N"
    "wrk"  = "Professional for Workstations"
    "wrkn" = "Professional for Workstations N"
    "ent"  = "Enterprise"
    "entn" = "Enterprise N"
    "iot"  = "IoT Enterprise"
    "iot21" = "IoT Enterprise LTSC 2021"
    "iot24" = "IoT Enterprise LTSC 2024"
    "home" = "Home"
}


# === Step 0: Show Available Editions and Aliases ===
if ($edition) {
    # Handle alias support and normalize input
    $editionInput = $edition.ToLower()
    if ($aliases.ContainsKey($editionInput)) {
        $selectedEdition = $aliases[$editionInput]
    } elseif ($editions.ContainsKey($edition)) {
        $selectedEdition = $edition
    } else {
        Write-Host "Invalid edition input via -edition. Available options:"
        $editions.Keys | Sort-Object
        exit
    }
} else {
    # Fallback to interactive mode if -edition wasn't supplied
    Write-Host "`nAvailable editions:" -ForegroundColor Blue
    $editionList = $editions.Keys | Sort-Object
    $i = 1
    $editionList | ForEach-Object { Write-Host "$i. $_"; $i++ }
    $userInput = Read-Host "`nYou can enter a number, full edition name, alias, or current"
    if ($userInput -match '^\d+$') {
    $selectedEdition = $editionList[$userInput - 1]
} elseif ($aliases.ContainsKey($userInput.ToLower())) {
    $selectedEdition = $aliases[$userInput.ToLower()]
} elseif ($editionList -contains $userInput) {
    $selectedEdition = $userInput
} elseif ($userInput.ToLower() -eq 'current') {
    $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions'
    $pfn = (Get-ItemProperty -Path $regPath -Name OSProductPfn).OSProductPfn
    $selectedEdition = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").EditionID
    $productKey = $editions[$selectedEdition]["Key"]
    Write-Host "`nDetected PFN from registry: $pfn" -ForegroundColor Yellow
    Write-Host "Proceeding using this PFN." -ForegroundColor Cyan
} else {
    Write-Host "Invalid selection. Try again." -ForegroundColor Red
    exit
}
}

# === Step 1: Prompt Edition ===
$editionList = $editions.Keys | Sort-Object

if ($userInput -match '^\d+$') {
    $selectedEdition = $editionList[$userInput - 1]
} elseif ($aliases.ContainsKey($userInput.ToLower())) {
    $selectedEdition = $aliases[$userInput.ToLower()]
} elseif ($editionList -contains $userInput) {
    $selectedEdition = $userInput
} elseif ($userInput.ToLower() -eq 'current') {
    $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions'
    $pfn = (Get-ItemProperty -Path $regPath -Name OSProductPfn).OSProductPfn
    $selectedEdition = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").EditionID
    $productKey = $editions[$selectedEdition][$Key]
    Write-Host "`nDetected PFN from registry: $pfn" -ForegroundColor Yellow
    Write-Host "Proceeding using this PFN." -ForegroundColor Cyan
} else {
    Write-Host "Invalid selection. Try again." -ForegroundColor Red
    exit
}


# Assign PFN and product key now
$pfn = $editions[$selectedEdition].Pfn
$productKey = $editions[$selectedEdition].Key

Write-Host "`nSelected Edition: $selectedEdition"
Write-Host "Using PFN: $pfn"
Write-Host "Using RTM Key: $productKey"


# === Step 2: Prepare Files Directory ===
$filesPath = "C:\Files"
if (!(Test-Path $filesPath)) {
    New-Item -Path $filesPath -ItemType Directory
}

# === Step 3: Create Ticket Function
function SPPTicket {
    param (
        [string]$Pfn,
        [string]$IID = "465145217131314304264339481117862266242033457260311819664735280",
        [string]$Timestamp = "2022-10-11T12:00:00Z",
        [string]$OutPath = "C:\ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket\genuineticket.xml"
    )

    function SignProperties {
        param ($Properties, $rsa)
        $hash = [Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($Properties))
        return [Convert]::ToBase64String($rsa.SignHash($hash, [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1))
    }

    [byte[]] $key = 0x07,0x02,0x00,0x00,0x00,0xA4,0x00,0x00,0x52,0x53,0x41,0x32,0x00,0x04,0x00,0x00,
                0x01,0x00,0x01,0x00,0x29,0x87,0xBA,0x3F,0x52,0x90,0x57,0xD8,0x12,0x26,0x6B,0x38,
                0xB2,0x3B,0xF9,0x67,0x08,0x4F,0xDD,0x8B,0xF5,0xE3,0x11,0xB8,0x61,0x3A,0x33,0x42,
                0x51,0x65,0x05,0x86,0x1E,0x00,0x41,0xDE,0xC5,0xDD,0x44,0x60,0x56,0x3D,0x14,0x39,
                0xB7,0x43,0x65,0xE9,0xF7,0x2B,0xA5,0xF0,0xA3,0x65,0x68,0xE9,0xE4,0x8B,0x5C,0x03,
                0x2D,0x36,0xFE,0x28,0x4C,0xD1,0x3C,0x3D,0xC1,0x90,0x75,0xF9,0x6E,0x02,0xE0,0x58,
                0x97,0x6A,0xCA,0x80,0x02,0x42,0x3F,0x6C,0x15,0x85,0x4D,0x83,0x23,0x6A,0x95,0x9E,
                0x38,0x52,0x59,0x38,0x6A,0x99,0xF0,0xB5,0xCD,0x53,0x7E,0x08,0x7C,0xB5,0x51,0xD3,
                0x8F,0xA3,0x0D,0xA0,0xFA,0x8D,0x87,0x3C,0xFC,0x59,0x21,0xD8,0x2E,0xD9,0x97,0x8B,
                0x40,0x60,0xB1,0xD7,0x2B,0x0A,0x6E,0x60,0xB5,0x50,0xCC,0x3C,0xB1,0x57,0xE4,0xB7,
                0xDC,0x5A,0x4D,0xE1,0x5C,0xE0,0x94,0x4C,0x5E,0x28,0xFF,0xFA,0x80,0x6A,0x13,0x53,
                0x52,0xDB,0xF3,0x04,0x92,0x43,0x38,0xB9,0x1B,0xD9,0x85,0x54,0x7B,0x14,0xC7,0x89,
                0x16,0x8A,0x4B,0x82,0xA1,0x08,0x02,0x99,0x23,0x48,0xDD,0x75,0x9C,0xC8,0xC1,0xCE,
                0xB0,0xD7,0x1B,0xD8,0xFB,0x2D,0xA7,0x2E,0x47,0xA7,0x18,0x4B,0xF6,0x29,0x69,0x44,
                0x30,0x33,0xBA,0xA7,0x1F,0xCE,0x96,0x9E,0x40,0xE1,0x43,0xF0,0xE0,0x0D,0x0A,0x32,
                0xB4,0xEE,0xA1,0xC3,0x5E,0x9B,0xC7,0x7F,0xF5,0x9D,0xD8,0xF2,0x0F,0xD9,0x8F,0xAD,
                0x75,0x0A,0x00,0xD5,0x25,0x43,0xF7,0xAE,0x51,0x7F,0xB7,0xDE,0xB7,0xAD,0xFB,0xCE,
                0x83,0xE1,0x81,0xFF,0xDD,0xA2,0x77,0xFE,0xEB,0x27,0x1F,0x10,0xFA,0x82,0x37,0xF4,
                0x7E,0xCC,0xE2,0xA1,0x58,0xC8,0xAF,0x1D,0x1A,0x81,0x31,0x6E,0xF4,0x8B,0x63,0x34,
                0xF3,0x05,0x0F,0xE1,0xCC,0x15,0xDC,0xA4,0x28,0x7A,0x9E,0xEB,0x62,0xD8,0xD8,0x8C,
                0x85,0xD7,0x07,0x87,0x90,0x2F,0xF7,0x1C,0x56,0x85,0x2F,0xEF,0x32,0x37,0x07,0xAB,
                0xB0,0xE6,0xB5,0x02,0x19,0x35,0xAF,0xDB,0xD4,0xA2,0x9C,0x36,0x80,0xC6,0xDC,0x82,
                0x08,0xE0,0xC0,0x5F,0x3C,0x59,0xAA,0x4E,0x26,0x03,0x29,0xB3,0x62,0x58,0x41,0x59,
                0x3A,0x37,0x43,0x35,0xE3,0x9F,0x34,0xE2,0xA1,0x04,0x97,0x12,0x9D,0x8C,0xAD,0xF7,
                0xFB,0x8C,0xA1,0xA2,0xE9,0xE4,0xEF,0xD9,0xC5,0xE5,0xDF,0x0E,0xBF,0x4A,0xE0,0x7A,
                0x1E,0x10,0x50,0x58,0x63,0x51,0xE1,0xD4,0xFE,0x57,0xB0,0x9E,0xD7,0xDA,0x8C,0xED,
                0x7D,0x82,0xAC,0x2F,0x25,0x58,0x0A,0x58,0xE6,0xA4,0xF4,0x57,0x4B,0xA4,0x1B,0x65,
                0xB9,0x4A,0x87,0x46,0xEB,0x8C,0x0F,0x9A,0x48,0x90,0xF9,0x9F,0x76,0x69,0x03,0x72,
                0x77,0xEC,0xC1,0x42,0x4C,0x87,0xDB,0x0B,0x3C,0xD4,0x74,0xEF,0xE5,0x34,0xE0,0x32,
                0x45,0xB0,0xF8,0xAB,0xD5,0x26,0x21,0xD7,0xD2,0x98,0x54,0x8F,0x64,0x88,0x20,0x2B,
                0x14,0xE3,0x82,0xD5,0x2A,0x4B,0x8F,0x4E,0x35,0x20,0x82,0x7E,0x1B,0xFE,0xFA,0x2C,
                0x79,0x6C,0x6E,0x66,0x94,0xBB,0x0A,0xEB,0xBA,0xD9,0x70,0x61,0xE9,0x47,0xB5,0x82,
                0xFC,0x18,0x3C,0x66,0x3A,0x09,0x2E,0x1F,0x61,0x74,0xCA,0xCB,0xF6,0x7A,0x52,0x37,
                0x1D,0xAC,0x8D,0x63,0x69,0x84,0x8E,0xC7,0x70,0x59,0xDD,0x2D,0x91,0x1E,0xF7,0xB1,
                0x56,0xED,0x7A,0x06,0x9D,0x5B,0x33,0x15,0xDD,0x31,0xD0,0xE6,0x16,0x07,0x9B,0xA5,
                0x94,0x06,0x7D,0xC1,0xE9,0xD6,0xC8,0xAF,0xB4,0x1E,0x2D,0x88,0x06,0xA7,0x63,0xB8,
                0xCF,0xC8,0xA2,0x6E,0x84,0xB3,0x8D,0xE5,0x47,0xE6,0x13,0x63,0x8E,0xD1,0x7F,0xD4,
                0x81,0x44,0x38,0xBF

    $rsa = New-Object Security.Cryptography.RSACryptoServiceProvider
    $rsa.ImportCspBlob($key)

    $SessionId = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes("OSMajorVersion=5;OSMinorVersion=1;OSPlatformId=2;PP=0;Pfn=$Pfn;PKeyIID=$IID;" + [char]0))
    $PropertiesStr = "OA3xOriginalProductId=;OA3xOriginalProductKey=;SessionId=$SessionId;TimeStampClient=$Timestamp"
    $SignatureStr = SignProperties $PropertiesStr $rsa

    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0">
  <version>1.0</version>
  <genuineProperties origin="sppclient">
    <properties>$PropertiesStr</properties>
    <signatures>
      <signature name="clientLockboxKey" method="rsa-sha256">$SignatureStr</signature>
    </signatures>
  </genuineProperties>
</genuineAuthorization>
"@

    [System.IO.File]::WriteAllText($OutPath, $xml, [System.Text.Encoding]::ASCII)
}

# === Step 5: Generate Ticket ===
SPPTicket -Pfn $pfn

try {
    Stop-Service -Name clipsvc -Force
    Start-Service -Name clipsvc
} catch {
    Write-Host "ClipSVC service restart failed: $_"
}
clipup.exe -v -o > $null 2>&1

# === Step 6: Change Product Key ===
cls
$keyChange = Read-Host "Install product key for $selectedEdition and activate? (Y/N)"
if ($keyChange.ToUpper() -eq "Y") {
    Write-Host "Uninstalling product key from registry and licensing store..." -ForegroundColor Yellow
    cscript.exe //B $env:windir\System32\slmgr.vbs /cpky
    cscript.exe //B $env:windir\System32\slmgr.vbs /upk
    Write-Host "Installing product key..." -ForegroundColor Cyan
    cscript.exe //B $env:windir\System32\slmgr.vbs /ipk $productKey
    cls
    $verify = & cscript.exe $env:windir\System32\slmgr.vbs /ato
	
	Write-Host "`nActivation attempted. Confirming status..."
	
	if ($verify -like "*activated successfully*") {
		Write-Host "Windows has been activated successfully. You may now close this window." -ForegroundColor Green
	} else {
		Write-Host "Please check your activation status. Opening Settings to confirm.." -ForegroundColor Cyan
		Start-Process -FilePath "ms-settings:activation"
    }
} else {
	Write-Host "Product key not installed for ticket. Please do this later manually, ignore it, or run the script again." -ForegroundColor Yellow
}

# === Step 7: Pause
pause
