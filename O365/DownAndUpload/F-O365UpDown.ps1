<# 
.SYNOPSIS 
    Example on how to Download and Upload Files to and from a SharePoint O365 Lib
.DESCRIPTION 
    This script uses the CSOM functionality 
    You have to Download "Microsoft.SharePoint.Client.dll" and "Microsoft.SharePoint.Client.Runtime.dll"
.NOTES 
    Author: TimTimBS
    Date: 05.01.2018
#>  

function Download-Files()
{

    Param(
      [Parameter(Mandatory=$True)]
      [String]$Url,

      [Parameter(Mandatory=$True)]
      [String]$UserName,

      [Parameter(Mandatory=$False)]
      [String]$Password, 

      [Parameter(Mandatory=$True)]
      [String]$SourceListTitle,

      [Parameter(Mandatory=$True)]
      [String]$TargetFolderPath

    )


    if($Password) {
       $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    }
    else {
      $SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,$SecurePassword)    
    $Context.Credentials = $Credentials

    $web = $Context.Web 
    $Context.Load($web)
    $Context.ExecuteQuery()
    $list = $web.Lists.GetByTitle($SourceListTitle);
    $rootFolder = $list.RootFolder
    $filesInRootFolder = $rootFolder.Files
    $Context.Load($filesInRootFolder)
    $Context.ExecuteQuery()

    foreach($file in $filesInRootFolder)
    {
        Write-Host $file.Name
        if ($Context.HasPendingRequest)
        {
            $Context.ExecuteQuery()
        }

        $fileReference = $file.ServerRelativeUrl
        $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context,$fileReference)
        $filePath = $TargetFolderPath + $file.Name
        $fileStream = [System.IO.File]::Create($filePath)
        $fileInfo.Stream.CopyTo($fileStream);
        $fileStream.Close()

    }
}

Function Ensure-Folder()
{
Param(
  [Parameter(Mandatory=$True)]
  [Microsoft.SharePoint.Client.Web]$Web,

  [Parameter(Mandatory=$True)]
  [Microsoft.SharePoint.Client.Folder]$ParentFolder, 

  [Parameter(Mandatory=$True)]
  [String]$FolderUrl

)

    $folderNames = $FolderUrl.Trim().Split("/",[System.StringSplitOptions]::RemoveEmptyEntries)
    $folderName = $folderNames[0]
    Write-Host "Creating folder [$folderName] ..."
    $curFolder = $ParentFolder.Folders.Add($folderName)
    $Web.Context.Load($curFolder)
    $web.Context.ExecuteQuery()
    Write-Host "Folder [$folderName] has been created succesfully. Url: $($curFolder.ServerRelativeUrl)"

    if ($folderNames.Length -gt 1)
    {
        $curFolderUrl = [System.String]::Join("/", $folderNames, 1, $folderNames.Length - 1)
        Ensure-Folder -Web $Web -ParentFolder $curFolder -FolderUrl $curFolderUrl
    }
}

Function Upload-File() 
{
Param(
  [Parameter(Mandatory=$True)]
  [Microsoft.SharePoint.Client.Web]$Web,

  [Parameter(Mandatory=$True)]
  [Microsoft.SharePoint.Client.List]$List,

  [Parameter(Mandatory=$True)]
  [String]$FolderRelativeUrl, 

  [Parameter(Mandatory=$True)]
  [System.IO.FileInfo]$LocalFile

)

    try 
    {
       $fileUrl = $FolderRelativeUrl + "/" + $LocalFile.Name
       Write-Host "Uploading file [$($LocalFile.FullName)] ..."       

        $FileFullName = $LocalFile.FullName
        $FileStream = New-Object IO.FileStream($FileFullName, [System.IO.FileMode]::Open)
        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $FileCreationInfo.Overwrite = $true
        $FileCreationInfo.ContentStream = $FileStream
        $FileCreationInfo.URL = $LocalFile.Name
        $FileUpload = $List.RootFolder.Files.Add($FileCreationInfo)
        $context.Load($FileUpload)
        $context.ExecuteQuery()
        
       Write-Host "File [$($LocalFile.FullName)] has been uploaded succesfully. Url: $fileUrl"
    }    
    catch 
    {
       write-host "An error occured while uploading file [$($LocalFile.FullName)]"
       
       Write-host $_

    }
    finally
    {
        $FileStream.Close()
    }
}

function Upload-Files()
{

Param(
  [Parameter(Mandatory=$True)]
  [String]$Url,

  [Parameter(Mandatory=$True)]
  [String]$UserName,

  [Parameter(Mandatory=$False)]
  [String]$Password, 

  [Parameter(Mandatory=$True)]
  [String]$TargetListTitle,

  [Parameter(Mandatory=$True)]
  [String]$SourceFolderPath

)

    if($Password) {
       $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    }
    else {
      $SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,$SecurePassword)    
    $Context.Credentials = $Credentials

    $web = $Context.Web 
    $Context.Load($web)
    $Context.ExecuteQuery()
    $list = $web.Lists.GetByTitle($TargetListTitle);
    $Context.Load($list.RootFolder)
    $Context.ExecuteQuery()

    
    Get-ChildItem $SourceFolderPath -Recurse | % {
       if ($_.PSIsContainer -eq $True) {
          $folderUrl = $_.FullName.Replace($SourceFolderPath,"").Replace("\","/")   
          if($folderUrl) {
             Ensure-Folder -Web $web -ParentFolder $list.RootFolder -FolderUrl $folderUrl
          }  
       }
       else{
          #$folderRelativeUrl = $list.RootFolder.ServerRelativeUrl + $_.DirectoryName.Replace($SourceFolderPath,"").Replace("\","/")  
          $folderRelativeUrl = $list.RootFolder.ServerRelativeUrl + $_.Name
          Upload-File -Web $web -List $list -FolderRelativeUrl $folderRelativeUrl -LocalFile $_ 
       }
    }
}

#Usage

cd $PSScriptRoot
Add-Type -Path "C:\Temp\O365\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Temp\O365\Microsoft.SharePoint.Client.Runtime.dll"

#Upload-Files -Url $Url -UserName $UserName -Password $Password -TargetListTitle $TargetListTitle -SourceFolderPath $SourceFolderPath  #-UrlAdminCenter $UrlAdminCenter 

Download-Files -Url "https://example.sharepoint.com/" -UserName user.name@example.onmicrosoft.com -Password "Passw0rd!" -SourceListTitle  "Dokumente" -TargetFolderPath "C:\Temp\O365\"