# Office
Auto Install and Costomise Profile Outlook. I create for 2021&amp;2019 Office, but may be changed for any MS Office version.
Costomise Outlook are personal gpo, you mast put in User OU like a StartScript (In costomise you have configure adressbook and set profile). In first run you mast set two password (pass for address book and user mail). 
For costomize use *.prf file and automail.ps1
Automail.ps1 - configure regestry for prf file and set variables for user and adbook.
For auto install ms office you mast configure xml file (i use https://config.office.com/) and you need install Office Deployment Tool (https://www.microsoft.com/en-us/download/details.aspx?id=49117).
After it you mast install ODT and configure path for office install (you may want to use microsoft servers for deploy application, i use local souce for it).
Use comand for prepair (setup /download {your config.xml}).
Use comand for install (setupo /configuration {your config.xml}).
Use script install.bat for gpo deploy.
Script mast put into computers ou, and set lite autostart script.
