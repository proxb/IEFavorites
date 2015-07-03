Function Get-IEFavorite {
    <#
        .SYNOPSIS
            Display a list of Internet Shortcuts and folders

        .DESCRIPTION
            Display a list of Internet Shortcuts and folders

        .PARAMETER Name
            Name of the links you wish to find.

        .NOTES
            Author: Boe Prox
            Name: Get-IEFavorite
            Created: 24 Dec 2013
            Version History 
                1.0 -- 24 Dec 2013
                    -Initial Version
                1.1 -- 26 Dec 2013
                    -Directory and File parameters for better filtering
                    -Try/Catch to handle errors with attempting to parse .lnk files
            
        .EXAMPLE
            Get-IEFavorite

            Name     : The WSUS Support Team Blog - Site Home - TechNet Blogs.url
            IsFolder : False
            IsLink   : True
            Url      : http://blogs.technet.com/b/sus/
            Path     : C:\Users\PROXB\Favorites\Links\The WSUS Support Team Blog - Site Home - TechNet Blogs.url

            Name     : WSUS Product Team Blog - Site Home - TechNet Blogs.url
            IsFolder : False
            IsLink   : True
            Url      : http://blogs.technet.com/b/wsus/
            Path     : C:\Users\PROXB\Favorites\Links\WSUS Product Team Blog - Site Home - TechNet Blogs.url 

            Description
            -----------
            Displays all of the favorites
        
        .EXAMPLE
            Get-IEFavorite -Directory

            Description
            -----------
            Displays all folders in Favorites

        .EXAMPLE
            Get-IEFavorite -Name WSUS*

            Description
            -----------
            Displays all Favorites with a name beginning with WSUS
    #>
    #Requires -Version 3.0
    [OutputType('System.IO.InternetShortcutFile','System.IO.InternetShortcutFolder')]
    [cmdletbinding(
        DefaultParameterSetName = 'All'
    )]
    Param (
        [parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string[]]$Name="*",
        [parameter(ParameterSetName= 'Directory')]
        [switch]$Directory,
        [parameter(ParameterSetName= 'File')]
        [switch]$File
    )
    Begin {
        $IEFav =  [Environment]::GetFolderPath('Favorites','None')
        $params = @{
            Recurse = $True
            Path = $IEFav
        }
        If ($PSBoundParameters.ContainsKey('Directory')) {
            $params['Directory'] = $True
        }
        If ($PSBoundParameters.ContainsKey('File')) {
            $params['File'] = $True
        }
    }
    Process {
        ForEach ($item in $Name) {
            $params['Filter'] = $item
            Get-ChildItem @params | ForEach {
                $object = $_
                Try {
                    If ($object.PSIsContainer) {
                        $Object = [pscustomobject]@{
                            Name = $object.Name
                            IsFolder = [bool]$object.PSIsContainer
                            IsLink = [bool]$False 
                            Url = $Null
                            Path = $Object.FullName
                        }  
                        $Object.pstypenames.insert(0,'System.IO.InternetShortcutFolder')
                    } Else {
                        $Object = [pscustomobject]@{
                            Name = $object.Name
                            IsFolder = [bool]$object.PSIsContainer
                            IsLink = [bool]$True 
                            Url = ($object | Select-String "^URL").Line.Trim("URL=")
                            Path = $Object.FullName
                        }  
                        $Object.pstypenames.insert(0,'System.IO.InternetShortcutFile')            
                    }  
                    $Object
                } Catch {}
            }
        }
    }
    End {}
}

Function Add-IEFavorite {
    <#
        .SYNOPSIS
            Adds a Favorite to IE

        .DESCRIPTION
            Adds a Favorite to IE

        .PARAMETER Name
            Name of the Favorite to add

        .PARAMETER Folder
            Place where Favorite will reside

        .PARAMETER Url
            Url of the Favorite

        .PARAMETER Force
            Force creation of folder path if it does not exist.

        .NOTES
            Author: Boe Prox
            Name: Add-IEFavorite
            Created: 26 Dec 2013
            Version 1.0 -- 26 Dec 2013
                -Initial Version    

        .EXAMPLE
            Add-IEFavorite -Name Bing -Url http://Bing.com

            Description
            -----------
            Adds favorite called Bing with url of http://bing.com

        .EXAMPLE
            Add-IEFavorite -Name Bing -Url http://Bing.com

            Description
            -----------
            Adds favorite called Bing with url of http://bing.com
    #>
    #Requires -Version 3.0
    [cmdletbinding(
        SupportsShouldProcess
    )]
    Param (
        [parameter(Mandatory)]
        [string]$Name,
        [parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Alias('FullName','Path')]
        [string]$Folder,
        [parameter(Mandatory)]
        [string]$Url
    )
    $IEFav =  [Environment]::GetFolderPath('Favorites') 
    $Shell = New-Object -ComObject WScript.Shell
    If ($PSBoundParameters.ContainsKey('Folder')) {
        $IEFav = Join-Path -Path $IEfav -ChildPath $Folder
        If (-Not (Test-Path $IEFav -PathType Container)) {
            Write-Verbose "Creating folder path: $($IEFav)"
            $null = New-Item -ItemType Directory -Path $IEfav -Force
        }
    }
    
    $FullPath = Join-Path -Path $IEFav -ChildPath "$($Name).url"

    If ($PSCmdlet.ShouldProcess("$Name -> $url ($Fullpath)",'Create Favorite')) {
        $shortcut = $Shell.CreateShortcut($FullPath)
        $shortcut.TargetPath = $Url
        $shortcut.Save()
    }
       
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Shell)        
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable Shell -WhatIf:$False    
}

Function Set-IEFavorite {
    <#
        .SYNOPSIS
            Sets an IE Favorite to a different url or name

        .DESCRIPTION
            Sets an IE Favorite to a different url or name

        .PARAMETER NewName
            New name to change the favorite to

        .PARAMETER Favorite
            Favorite being updated

        .PARAMETER NewUrl
            New url of the favorite

        .NOTES
            Author: Boe Prox
            Name: Set-IEFavorite
            Created: 26 Dec 2013
            Version 1.0 -- 26 Dec 2013
                -Initial Version    

        .EXAMPLE
            Get-IEFavorite -Name Google | Set-IEFavorite -NewName Bing -NewUrl http://Bing.com

            Description
            -----------
            Updates the Google favorite to now use Bing instead
    #>
    #Requires -Version 3.0
    [cmdletbinding(
        SupportsShouldProcess
    )]
    Param(
        [parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$NewName,
        [parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Alias('FullName','Path')]
        [string]$Favorite,
        [parameter()]
        [string]$NewUrl
    )
    Process {
        If ($PSBoundParameters['NewUrl']) {
            If ($PSCmdlet.ShouldProcess("$Favorite",'Set Favorite URL')) {
                $lines = Get-Content $Favorite
                ForEach ($line in $lines) {
                    If ($line.StartsWith("BASEURL")) {
                        $lines[$i] = "BASEURL=$NewUrl"
                    }
                    If ($line.StartsWith("URL")) {
                        $lines[$i] = "URL=$NewUrl"
                    }            
                    If ($line.StartsWith("IconFile")) {
                        $lines[$i] = "IconFile=$NewUrl/favicon.ico"
                    }            
                    $i++  
                }
                Set-Content -Value $lines -Path $Favorite
            }
        }
        If ($PSBoundParameters.ContainsKey('NewName')) {
            If ($PSCmdlet.ShouldProcess("$Favorite",'Set Favorite Name')) {
                Rename-Item -Path $Favorite -NewName "$NewName.url"
            }
        } 
    }
}

Set-Alias -Name gief -Value Get-IEFavorite
Set-Alias -Name aief -Value Add-IEFavorite
Set-Alias -Name rief -Value Remove-Item
Set-Alias -Name sief -Value Set-IEFavorite

Export-ModuleMember -Function * -Alias *