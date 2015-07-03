# IEFavorites
A module that allows you to add/remove IE favorites using PowerShell

````PowerShell
Get-IEFavorite

Name     : The WSUS Support Team Blog - Site Home - TechNet Blogs.url
IsFolder : False
IsLink   : True
Url      : http://blogs.technet.com/b/sus/
Path     : C:\Users\PROXB\Favorites\Links\The WSUS Support Team Blog - Site Home - TechNet Blogs.url
````

````PowerShell
Add-IEFavorite -Name Bing -Url http://Bing.com
````

````PowerShell
Get-IEFavorite -Name Google | Set-IEFavorite -NewName Bing -NewUrl http://Bing.com
````
