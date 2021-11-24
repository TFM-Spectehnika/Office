#Clear-Host

#Set variables
$OfficeRegPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

#Get data forom AD
$UserName = $env:username
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()

#Set system enviroment
[Environment]::SetEnvironmentVariable("mail", $ADUser.mail , "User")
[Environment]::SetEnvironmentVariable("cname", $ADUser.DisplayName , "User")

#Check gegistry for value, if not set PRF
if (Test-Path $OfficeRegPath) {
Exit}
Else {
New-Item -path "HKCU:\Software\Microsoft" -name "Office"
New-Item -path "HKCU:\Software\Microsoft\Office" -name "16.0"
New-Item -path "HKCU:\Software\Microsoft\Office\16.0" -name "Outlook"
New-Item -path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -name "Setup"
Set-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup" -name ImportPRF -Value "\\srv-db1\Config.1C$\Outlook\outlook.prf"
}
