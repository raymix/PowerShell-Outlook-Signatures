# Written by Raimonds Virtoss @ https://github.com/raymix

### Signature location ###
$UNC = "\\FILESERVER\Outlook Signatures"

# Define DATA folder on network share
$DATA = Join-Path $UNC "DATA"
# Define local path to Outlook signatures
$signaturePath = Join-Path $env:APPDATA "Microsoft\Signatures"
# Create signatures folder if doesn't exist yet
if (!(Test-Path $signaturePath)) {New-Item -Path $signaturePath -ItemType Directory}

### AD STUFF ###

# Get current user's details from Active Directory using ADSI
$adsi = [adsisearcher]"(samaccountname=$env:USERNAME)"
$userADProperties = $adsi.FindOne().Properties

$fullName = $userADProperties.displayname
$jobTitle = $userADProperties.title
$mobile = $userADProperties.mobile
$telephone = $userADProperties.telephonenumber
$email = $userADProperties.mail
$SAM = $userADProperties.samaccountname
$DN = $userADProperties.distinguishedname

# You can override above here if you need to do some testing
#$DN = "CN=USERNAME,OU=Users,OU=Location1,OU=Site1,DC=DOMAIN,DC=local"
#$SAM = "USERNAME"

#Allows more dynamic choice by using $phones, best used for users with no phones as it will return empty.
$phones = ""
if ($telephone) { $phones += "<span style='color:#215732;'><strong>T:</strong></span>&nbsp;<a href='tel:#telephone' style='color:#215732;'>$telephone</a>" }
if ($telephone -and $mobile) {$phones +=  "&nbsp;<span style='color: #BA0C2F;'>|</span>&nbsp;"} 
if ($mobile) {$phones += "<span style='color:#215732;'><strong>M:</strong></span>&nbsp;<a href='tel:#mobile' style='color:#215732;'>$mobile</a>"} #"<strong>M:</strong>&nbsp;$($mobile)" }

$templatePath = $UNC

# Find which OU user belongs to and which template folder to read from on FILESERVER
# This is where you will need to do some modifications depending on your environment

# This part extracts LocationX and SiteX from Distinguished Name
$OU = $DN -replace "CN=$($SAM),","" -replace ",DC=DOMAIN,DC=local","" -replace "OU=Users,","" -replace "OU=","" -split ","
# Now reverse the resulted array
$OU = ($OU[($OU.Length-1)..0])
# Take UNC path and add above reversed result to form full path to locate signature files on network share
$($OU | % {$templatePath = Join-Path $templatePath $_})

### LOCAL STUFF ###

#Load existing templates and read hashes from first line
$localHTM = Get-ChildItem $signaturePath -Filter "*.htm"

$localSignatures = $null
$localSignatures = @()
foreach ($htm in $localHTM) {
    $lastLine = ((Get-Content $htm.FullName -Last 1) -replace "<!-- ","" -replace " -->","").Split(",") #<!-- MD5=XXXXXXXXX -->

    $localSignatures +=, [pscustomobject]@{
        Name=$htm.Name
        Base=$htm.BaseName
        Path=$htm.Directory
        SAM=$lastLine[0]
        jobTitle=$lastLine[1]
        mobile=$lastLine[2]
        telephone=$lastLine[3]
        email=$lastLine[4]
        MD5=$(if ($lastLine -match $SAM) { $lastLine[5] } else { "N/A" })
    }
}

#Windows bubble message
function Send-Notification($title, $msg) {
    Add-Type -AssemblyName System.Windows.Forms 
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = $(Join-Path $UNC "DATA\balloon.ico")
    $balloon.BalloonTipText = $msg
    $balloon.BalloonTipTitle = $title 
    $balloon.Visible = $true 
    $balloon.ShowBalloonTip(10000)
    $balloon.Dispose()
}

function Write-Signature($md5, $template) {
    
    $HTM_files = $($template.BaseName + "_files")
    $HTM_filesNoSpace = $HTM_files -replace " ","%20"
    $localHTM_files = Join-Path $signaturePath $HTM_files
    $localHTM_tmp = Join-Path $signaturePath $("tmp_" + $template.Name)
    $localHTM_tmp2 = Join-Path $signaturePath $("tmp2_" + $template.Name)
    $localHTM = Join-Path $signaturePath $template.Name
    $localRTF = "$(Join-Path $signaturePath $template.BaseName).rtf"
    $localTXT = "$(Join-Path $signaturePath $template.BaseName).txt"
    $image = "$(Join-Path $template.Directory $template.BaseName).jpg"

    if (Test-Path $localHTM_files) {
        Remove-Item $localHTM_files -Recurse -Force
    }

    New-Item -Path $localHTM_files -ItemType Directory

    if (Test-Path $image) {
        Copy-Item -Path $image -Destination (Join-Path $localHTM_files "image001.jpg")    
    }

    # Write HTM file, keep the string indented to left
"$((Get-Content $template.FullName -Raw) -replace "#fullName",$fullName -replace "#jobTitle",$jobTitle -replace "#phones",$phones -replace "#mobile",$mobile -replace "#telephone",$telephone -replace "#email",$email -replace "#head",'' -replace "#folder",$HTM_filesNoSpace)" | Out-File $localHTM_tmp -Encoding utf8

    # Convert HTM to RTF locally
    $wrd = new-object -com word.application 
    $wrd.visible = $false 
    $doc = $wrd.documents.open($localHTM_tmp) # needs unused var defined
    $opt = 6
    $wrd.ActiveDocument.Saveas([ref]$localRTF,[ref]$opt)
    $wrd.Quit()

    # Convert HTM to TXT and strip all html stuff, tabulators and empty lines
    $txt = Get-Content $localHTM_tmp
    $txt = $txt -replace "<style.+\/style>","" # Remove styles
    $txt = $txt -replace "<[^>]*>","" # HTML tags and comments
    $txt = $txt -replace "&nbsp;"," " # HTML strong spaces character
    $txt = $txt -replace "&#173;"," " # HTML decimal char?
    $txt = $txt.trim() # Tabulators
    $txt = $txt | ? {$_.trim() -ne "" }  # Empty line breaks

    $txt | Out-File $localTXT
    
    # Write HTM file, keep the string indented to left
"$((Get-Content $template.FullName -Raw) -replace "#fullName",$fullName -replace "#jobTitle",$jobTitle -replace "#phones",$phones -replace "#mobile",$mobile -replace "#telephone",$telephone -replace "#email",$email -replace "#folder",$HTM_filesNoSpace)" | Out-File $localHTM_tmp2 -Encoding utf8

    $localHTMFromTMP = Get-Content $localHTM_tmp2 -Raw
    $xmlns = Get-Content (Join-Path $DATA "xmlns.htm") -Raw
    $head = ((Get-Content (Join-Path $DATA "head.htm") -Raw) -replace "#folder",$HTM_filesNoSpace)
    $filelist = ((Get-Content (Join-Path $DATA "filelist.xml") -Raw) -replace "#folder",$HTM_filesNoSpace -replace "#baseName","$($template.BaseName -replace " ","%20").htm")

    $localHTMFromTMP = $localHTMFromTMP -replace "<!DOCTYPE html>",$xmlns
    $localHTMFromTMP = $localHTMFromTMP -replace "#head",$head

    $localHTMFromTMP | Out-File $localHTM
    # Add n HTML comment at the last line of htm file, this is used to check for changes in AD and template file (md5 file checksum) later
    Add-Content $localHTM "<!-- $SAM,$jobTitle,$mobile,$telephone,$email,$md5 -->"

    $filelist | Out-File (Join-Path $localHTM_files "filelist.xml")

    Copy-Item -Path (Join-Path $DATA "colorschememapping.xml") -Destination $localHTM_files
    Copy-Item -Path (Join-Path $DATA "themedata.thmx") -Destination $localHTM_files

    Remove-Item $localHTM_tmp -Force
    Remove-Item $localHTM_tmp2 -Force

}

function Commit-Signatures($templates) {
    foreach ($template in $templates) {
        $md5 = (Get-FileHash $template.FullName -Algorithm MD5).Hash
        if (Test-Path (Join-Path $signaturePath $template.Name)) { # If template file exists
            
            # Check md5 in hashtables to see if signature is updated and needs replacing
            if ($md5 -notin $localSignatures.md5) {
                Write-Signature $md5 $template
                Send-Notification "Company Signature Updated" "Outlook signature [$($template.BaseName)] has been updated"
            } else {
                $findChanged = $localSignatures | Where-Object {$_.Name -match $template.Name}
                if (($findChanged.SAM -ne $SAM) -or ($findChanged.jobTitle -ne $jobTitle) -or ($findChanged.mobile -ne $mobile) -or ($findChanged.telephone -ne $telephone) -or ($findChanged.email -ne $email)) {
                    Write-Signature $md5 $template
                    Send-Notification "Company Signature Details Updated" "Outlook signature [$($template.BaseName)] has been updated to reflect changes of your profile in Active Directory (such as name, email, phone or mobile number)"
                }
            }

        } else { # Template file does not exist
            Write-Signature $md5 $template
            Send-Notification "Company Signature Added" "A new Outlook signature [$($template.BaseName)] has been added to your Outlook."
        }
   
    }
}

# I named variable country templates because our OUs are based on countries. In this public version this variable is for LocationX
$countryTemplates = Get-ChildItem $templatePath -Filter "*.htm"
$defaultTemplates = Get-ChildItem $UNC -Filter "*.htm"

foreach ($ls in $localSignatures) {
    if ($ls.MD5 -ne "N/A") {
        if ((($ls.Base -notin $countryTemplates.BaseName) -and $countryTemplates.count -gt 0) -or (($ls.Base -in $defaultTemplates.BaseName)-and $countryTemplates.count -gt 0) -or (($ls.Base -notin $defaultTemplates.BaseName)-and $countryTemplates.count -eq 0)){ 
            Remove-Item "$(Join-Path $ls.Path $ls.Base).htm" -Force -ErrorAction SilentlyContinue
            Remove-Item "$(Join-Path $ls.Path $ls.Base).rtf" -Force -ErrorAction SilentlyContinue
            Remove-Item "$(Join-Path $ls.Path $ls.Base).txt" -Force -ErrorAction SilentlyContinue
            Remove-Item "$(Join-Path $ls.Path $ls.Base)_files" -Recurse -Force -ErrorAction SilentlyContinue
            Send-Notification "Company Signature Removed" "An Outlook signature [$($ls.Base)] has been removed from your Outlook."
        }
    }
}

if ($countryTemplates.count -eq 0) {
    if ($defaultTemplates.count -gt 0) {
        Commit-Signatures $defaultTemplates
    }
} else {
    Commit-Signatures $countryTemplates
}




