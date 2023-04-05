$MainMenu = {
Write-Host " ******************************************************"
Write-Host " * Microsoft Office 365 for Windows 7 Installation    *" 
Write-Host " * Author: bonguides.com                              *" 
Write-Host " * Updated on: 10/16/2022                             *" 
Write-Host " ******************************************************" 
Write-Host 
Write-Host " 1. Install Office 32-bit" 
Write-Host " 2. Install Office 64-bit"
Write-Host " 3. Quit or Press Ctrl + C"
Write-Host 
Write-Host " Select an option and press Enter: "  -nonewline
}
cls

$install = {
   New-Item -Path $env:temp\temp -ItemType Directory -Force | Out-Null
   cd $env:temp\temp
   Write-Host 'Downloading........'
   (New-Object System.Net.WebClient).DownloadFile("https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/bonguides.com/Scripts/O365_Windows7_$($arch)bit/setup.exe","$env:temp\temp\setup.exe")
   (New-Object System.Net.WebClient).DownloadFile("https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/bonguides.com/Scripts/O365_Windows7_$($arch)bit/configuration.xml","$env:temp\temp\configuration.xml")
   Write-Host 'Installing........'
   .\setup.exe /configure .\configuration.xml
   Write-Host 'Done........'
}

Do { 
cls
Invoke-Command $MainMenu
$Select = Read-Host
if ($select -eq 1) {$arch= '32'}
if ($select -eq 1) {$arch= '64'}
Switch ($Select)
   {
      #1. Install Office 32-bit
       1 {
            $365SubMenu = {
            Write-Host " *************************************************"
            Write-Host " * Select a Microsoft Office 365 Product         *" 
            Write-Host " *************************************************" 
            Write-Host 
            Write-Host " 1. Office 365 Home" 
            Write-Host " 2. Office 365 Personal"
            Write-Host " 3. Microsoft 365 Apps for Business" 
            Write-Host " 4. Microsoft 365 Apps for Enterprise" 
            Write-Host " 5. Go Back"
            Write-Host 
            Write-Host " Select an option and press Enter: "  -nonewline
            } 
            cls
                     
         Do { 
         cls
         Invoke-Command $365SubMenu
         $365SubSelect = Read-Host
         Switch ($365SubSelect)
                  {
                     1 {Invoke-Command $install}
                     2 {Invoke-Command $install}
                     3 {Invoke-Command $install}
                     4 {Invoke-Command $install}
                     }
            }
         While ($365SubSelect -ne 5)
         cls
         }  
   }
}
While ($Select -ne 3 )
