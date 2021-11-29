#Clear-Host

#Create Variables
$OfficeRegPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

#Connect to AD, get information
$UserName = $env:username
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()
$Cname = $ADUser.DisplayName
$Mail = $ADUser.mail

#Set folder for pfr
$AppData=(Get-Item env:appdata).value
$PRFPath = "\Microsoft\"
$LocalPRFPath = $AppData+$PRFPath

#Set share path for PRF
$NetPRFFile = "\\net_share\outlook.prf"

#Copy from net_share to local path
Copy-Item -Path $NetPRFFile $LocalPRFPath -Force
#Change data
((Get-Content -path $LocalPRFPath"outlook.prf" -Raw) -replace '%cname%',$Cname) | Set-Content -Path $LocalPRFPath"outlook.prf"
((Get-Content -path $LocalPRFPath"outlook.prf" -Raw) -replace '%mail%',$Mail) | Set-Content -Path $LocalPRFPath"outlook.prf"



#Check reg path
if (Test-Path $OfficeRegPath) {
Exit}
Else {
New-Item -path "HKCU:\Software\Microsoft" -name "Office"
New-Item -path "HKCU:\Software\Microsoft\Office" -name "16.0"
New-Item -path "HKCU:\Software\Microsoft\Office\16.0" -name "Outlook"
New-Item -path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -name "Setup"
Set-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup" -name ImportPRF -Value $LocalPRFPath"outlook.prf"
}
