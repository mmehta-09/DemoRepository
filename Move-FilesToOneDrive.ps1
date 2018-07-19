$OneDrivePath = $env:USERPROFILE + "\OneDrive - Blackbaud, Inc"
$OldOneDrivePath = $env:USERPROFILE + "\OneDrive for Business"

$WaitTime = 180

if((Get-Process OneDrive -ErrorAction SilentlyContinue) -and (Test-Path $OneDrivePath))
{
     
    if(Test-Path $OldOneDrivePath)
    {
        throw "the script doesn't work for old one drive"
    }

    $UserProfileBackupPath = $OneDrivePath + "\" + $env:UserName 

    if(-not(Test-Path $UserProfileBackupPath))
    {
        New-Item -Path $UserProfileBackupPath -ItemType Directory
    }
    
    Write-Host "Moving contents of desktop folder"
    $Desktop = $env:USERPROFILE + "\Desktop"
    $DesktopDest = $UserProfileBackupPath + "\Desktop"
    if(-not(Test-Path $DesktopDest))
    {
        New-Item -Path $DesktopDest -ItemType Directory
    }

    $Contents = Get-ChildItem $Desktop -Recurse
    $Contents = $Contents | where {$_.Extension -ne ".lnk"}

    if($Contents)
    {
    foreach($item in $Contents)
    {
        Move-Item $item.FullName -Destination $DesktopDest -Force
    }
    }
    Write-Host "Contents of desktop folder moved"

    $Documents = $env:USERPROFILE + "\Documents"
    if(-not(Test-Path $Documents))
    {
        throw "Documents folder doesn't exist"
    }
    $DocumentsDest = $UserProfileBackupPath + "\Documents"
    if(-not(Test-Path $DocumentsDest))
    {
        New-Item -Path $DocumentsDest -ItemType Directory
    }

    Write-Host "Moving contents of documents folder"

    $Contents = Get-ChildItem $Documents -Recurse
    
    if($Contents)
    {
    foreach($item in $Contents)
    {
        Move-Item $item.FullName -Destination $DocumentsDest -Force
    }
    }
    
    write-Host "Contents of documents folder moved"

}

else
{
        if(-not(Test-Path "$env:LocalAppData\Microsoft\OneDrive\OneDrive.exe"))
        {
            throw "One drive is not installed in the system"
            
        }
        else
        {
            start-process "$env:LocalAppData\Microsoft\OneDrive\OneDrive.exe"   
            Start-Sleep -Seconds $WaitTime 
        }     
    
    

    if(-not(Test-Path $OneDrivePath))
    {
        $wshell = new-object -comobject wscript.shell

        $wshell.popup(“Unsuccessful Login to Onedrive“,0,"OneDrive Error”)
        Exit
    }
        
    if(Test-Path $OldOneDrivePath)
    {
        throw "the script doesn't work for old one drive"
    }

    $UserProfileBackupPath = $OneDrivePath + "\" + $env:UserName 

    if(-not(Test-Path $UserProfileBackupPath))
    {
        New-Item -Path $UserProfileBackupPath -ItemType Directory
    }

    $OtherBackupPath = $UserProfileBackupPath + "\Other"

    if(-not(Test-Path $OtherBackupPath))
    {
        New-Item -Path $OtherBackupPath -ItemType Directory
    }

    #Backing up signature folder
    Write-Host "Backing up signatures folder"
    $SignaturesSource = $env:APPDATA + "\Microsoft\Signatures"
    $SignaturesDest = $OtherBackupPath + "\SignaturesBackup"

    if(-not(Test-Path $SignaturesDest))
    {
        New-Item -Path $SignaturesDest -ItemType Directory
    }
    if(Test-Path $SignaturesSource)
    {
        Copy-Item -Path $SignaturesSource -Destination $SignaturesDest -Recurse -Force
    }

    Write-Host "Backup completed"


    #Moving all folders from UserProfile to One drive folder
    $items = Get-ChildItem $env:UserProfile
    foreach($item in $items)
                                                                                                                {
    if(-not($item.Name -eq "OneDrive - Blackbaud, Inc"))
    {
        if($item.Name -eq "Desktop")
        {
            Write-Host "Moving desktop folder"
            $Desktop = $env:USERPROFILE + "\Desktop"
            $DesktopDest = $UserProfileBackupPath + "\Desktop"
            if(-not(Test-Path $DesktopDest))
            {
                New-Item -Path $DesktopDest -ItemType Directory
            }

            $Contents = Get-ChildItem $Desktop -Recurse
            $Contents = $Contents | where {$_.Extension -ne ".lnk"}
            if($Contents)
            {
            foreach($item in $Contents)
            {
                Move-Item $item.FullName -Destination $DesktopDest
            }
            }

            Write-Host "Desktop folder moved"
        }
        else
        {
            $folder = $item.Name
            Write-Host "Moving " $folder " folder"
            Move-Item $item.FullName $UserProfileBackupPath -Force
            write-Host $folder "folder moved"
        }
    }
    }

    #Copy chrome bookmarks
    Write-Host "Copying chrome bookmarks"
    $Bookmarks = $env:LocalAppData + "\Google\Chrome\User Data\Default\Bookmarks"
    $ChromePath = $OtherBackupPath + "\Chrome"
    if(-not(Test-Path $ChromePath))
        {
    New-Item -Path $ChromePath -ItemType Directory
    }

    if(Test-Path $Bookmarks)
        {
    Copy-Item $Bookmarks $ChromePath
    }
    Write-Host "Files copied"


    #Copy firefox profile
    Write-Host "Copying firefox profile folder"
    $Firefox = $OtherBackupPath + "\Firefox"
    $Profiles = $env:AppData + "\Mozilla\Firefox\Profiles"
    if(-not(Test-Path $Firefox))
        {
    New-Item -Path $Firefox -ItemType Directory
    }
    if(Test-Path $Profiles)
        {
    Copy-Item $Profiles $Firefox -Recurse -Force
    }
    Write-Host "Folder copied"

    #Copy all .pst file from UserProfile
    Write-Host "Copying .pst files"
    $AppData = Get-ChildItem $env:UserProfile -Recurse 
    $Files = $AppData | where {$_.extension -eq ".pst"}
    if($Files)
                    {
    foreach($file in $Files)
    {
        Copy-Item $file.FullName $OtherBackupPath 
    }
    }
    Write-Host "Files copied"

}

#Set-ExecutionPolicy restricted
