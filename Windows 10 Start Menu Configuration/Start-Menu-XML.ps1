Param(
    [String]$XMLFileToParse = "$PSScriptRoot\StartMenuSettings.xml"
)

##### Variables #####

$script:LogFile = "$ENV:Temp\StartMenuLayout.log"
$script:MaxLogSizeInKB = 2048
$script:ScriptsFolderPath = "$PSScriptRoot\Scripts"
$script:Settings

#region Functions & Classes

function Log
{
  param (
  [Parameter(Mandatory=$true)]
  $message,
  [Parameter(Mandatory=$true)]
  $component,
  [Parameter(Mandatory=$true)]
  $type )

  switch ($type)
  {
    1 { $type = "Info" }
    2 { $type = "Warning" }
    3 { $type = "Error" }
    4 { $type = "Verbose" }
  }

  if (($type -eq "Verbose") -and ($script:Verbose))
  {
    $toLog = "{0} `$$<{1}><{2} {3}><thread={4}>" -f ($type + ":" + $message), ($script:ScriptName + ":" + $component), (Get-Date -Format "MM-dd-yyyy"), (Get-Date -Format "HH:mm:ss.ffffff"), $pid
    $toLog | Out-File -Append -Encoding UTF8 -FilePath ("filesystem::{0}" -f $script:LogFile)
    Write-Host $message
  }
  elseif ($type -ne "Verbose")
  {
    $toLog = "{0} `$$<{1}><{2} {3}><thread={4}>" -f ($type + ":" + $message), ($script:ScriptName + ":" + $component), (Get-Date -Format "MM-dd-yyyy"), (Get-Date -Format "HH:mm:ss.ffffff"), $pid
    $toLog | Out-File -Append -Encoding UTF8 -FilePath ("filesystem::{0}" -f $script:LogFile)
    Write-Host $message
  }
  if (($type -eq 'Warning') -and ($script:ScriptStatus -ne 'Error')) { $script:ScriptStatus = $type }
  if ($type -eq 'Error') { $script:ScriptStatus = $type }

  if ((Get-Item $script:LogFile).Length/1KB -gt $script:MaxLogSizeInKB)
  {
    $log = $script:LogFile
    Remove-Item ($log.Replace(".log", ".lo_"))
    Rename-Item $script:LogFile ($log.Replace(".log", ".lo_")) -Force
  }
} 

class CustomTaskbarLayoutCollection
{
    $PinListPlacement = "Replace"
    $Header = '<CustomTaskbarLayoutCollection PinListPlacement="' + $this.PinListPlacement + '">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>'
    $Footer = ' </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>'
  [string]$XML
  [string]$DesktopApps = @()

  Add($XML)
  {
    $this.DesktopApps += $XML
    $this.XML = $This.Header + $this.DesktopApps + $this.Footer
  }
}

class TaskBarDesktopApp
{
    $Header = '<taskbar:DesktopApp DesktopApplicationLinkPath="'
    $Footer = '"/>'
    $XML = ""

    New($Shortcut){
         $This.XML = $This.header + $Shortcut + $This.Footer
    }
}

class DesktopApplicationTile
{
    [String]$DATHeader = '<start:DesktopApplicationTile '
    [String]$Size
    [String]$Column
    [String]$Row
    [String]$DATLinkPath
    [String]$DATFooter = "/>"
    [String]$DATCompiledXML

    New($Size, $Column, $Row, $DatLinkPath)
    {
        $This.Size = $size
        $This.Column = $Column
        $This.Row = $Row
        $this.DATLinkPath = $DatLinkPath
        $this.DATCompiledXML = $This.DATHeader + 'Size="' + $this.Size + '" Column="' + $this.Column + '" Row="' + $this.Row + `
            '" DesktopApplicationLinkPath="' + $this.DATLinkPath + '" ' + $this.DATFooter
    }

}

Class StartMenuGroup
{
    [String]$GroupHeader = '<start:Group Name="'
    [string]$GroupFooter = '</start:Group>'
    [string]$DATS
    [String]$SMGXML
    [String]$Name
    [Int]$NumberOfApps = 0

    Add($DatToAdd)
    {
        $this.DATS += $DatToAdd
    }

    New($Name)
    {
        $this.Name = $Name
    }

    GenXML()
    {
        $this.SMGXML = $this.GroupHeader + $This.Name + '">' + $this.DATS + $this.GroupFooter
    }
}

Class LayoutModification
{
    [string]$Header = '<LayoutModificationTemplate 
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout" 
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout" 
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <LayoutOptions StartTileGroupCellWidth="6" />
  <DefaultLayoutOverride>
    <StartLayoutCollection>
      <defaultlayout:StartLayout GroupCellWidth="6">'
      [string]$SMGFooter = '</defaultlayout:StartLayout>
        </StartLayoutCollection>
      </DefaultLayoutOverride>'

    [string]$Footer = '
    </LayoutModificationTemplate>'
    [String]$SMGS
    [string]$TaskBarItems
    
    Add($SMG)
    {
        $This.SMGS += $SMG
    } 

    AddTBI($TBI)
    {
        Foreach($Item in $TBI)
        {
            $this.TaskBarItems += $Item
        }
    }

    GenXML($FilePath)
    {
        $XML = [XML]$($This.Header + $This.SMGS + $this.SMGFooter + $This.TaskBarItems + $This.Footer)
        $XML.Save($FilePath)
    }
    GenCleanXML($FilePath)
    {
        $Inputxml = ""
        if(!$this.TaskBarItems)
        {
        $Inputxml = '<LayoutModificationTemplate xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout" xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout" xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout" Version="1" xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification">
          <LayoutOptions StartTileGroupCellWidth="6" />
          <DefaultLayoutOverride>
            <StartLayoutCollection>
              <defaultlayout:StartLayout GroupCellWidth="6" />
            </StartLayoutCollection>
          </DefaultLayoutOverride>
            <CustomTaskbarLayoutCollection PinListPlacement="Replace">
            <defaultlayout:TaskbarLayout>
              <taskbar:TaskbarPinList>
                <taskbar:DesktopApp DesktopApplicationLinkPath="#leaveempty"/>
              </taskbar:TaskbarPinList>
            </defaultlayout:TaskbarLayout>
          </CustomTaskbarLayoutCollection>
        </LayoutModificationTemplate>'
           
        }
        else
        {
            #We have task bar items.. our xml just got more complex!
            $InputHeader = '<LayoutModificationTemplate 
	            xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout" 
	            xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout" 
	            xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout" 
	            Version="1" 
	            xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification">
                <LayoutOptions StartTileGroupCellWidth="6" />
		            <DefaultLayoutOverride>
			            <StartLayoutCollection>
				            <defaultlayout:StartLayout GroupCellWidth="6" />
                        </StartLayoutCollection>
                      </DefaultLayoutOverride>'

          $InputFooter = '</LayoutModificationTemplate>'

          $Inputxml = $InputHeader + $this.TaskBarItems + $InputFooter
        }

        $XML = [xml]$Inputxml
        $XML.Save($FilePath)
    }

}

Class Settings
{
    [String]$CompanyName
    [boolean]$CleanStartLayout
    [String]$AppsToPin
    [boolean]$ShowHiddenFiles
    [boolean]$ShowFileExt
    [System.Xml.XmlElement]$StartMenuGroups
    
    New([String]$XMLFile)
    {
        [xml]$XmlReader = Get-Content -Path $XMLFile
        $this.CompanyName = $XmlReader.StartMenuSettings.CompanyName.CompanyName
        $this.CleanStartLayout = [System.Convert]::ToBoolean($XmlReader.StartMenuSettings.CleanStartLayout.CleanStartLayout)
        $this.AppsToPin = $XmlReader.StartMenuSettings.AppsToPin.AppsToPin
        #If we aren't doing a clean startlayout load in our start menu groups
        if($this.CleanStartLayout -eq $false)
        {
            $this.StartMenuGroups = $XmlReader.StartMenuSettings.StartMenuGroups
        }
               
    }
}

Function Get-TaskBarItems($AppNames)
{
    $StartMenuLocation = "$ENV:ProgramData\Microsoft\Windows\Start Menu\Programs"
    $StartMenuLNKS = GCI -LiteralPath $StartMenuLocation -Recurse | Where-Object {$_.Extension -like ".lnk"}
    $ReturnTBC = New-Object CustomTaskbarLayoutCollection
    Foreach($App in $AppNames)
    {
        Log -message "Attempting to find $App" -component "Get-TaskBarItems" -type Info
        Write-Debug -Message "Looking for $App LNK file.."
        if($App -eq "File Explorer")
        {
            $LNKPath = New-Object psobject
            $LNKPath | Add-Member -MemberType NoteProperty -Name "FullName" -Value '%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\File Explorer.lnk'
        }
        elseif($App -eq "Internet Explorer")
        {
            $LNKPath = New-Object PSObject
            $LNKPath | Add-Member -MemberType NoteProperty -Name "FullName" -Value '%APPDATA%\Microsoft\Windows\Start Menu\Programs\Accessories\Internet Explorer.lnk'
        }
        elseif($App -eq "Onedrive")
        {
            $LNKPath = New-Object PSObject
            $LNKPath | Add-Member -MemberType NoteProperty -Name "FullName" -Value '%APPDATA%\Microsoft\Windows\Start Menu\Programs\Onedrive.lnk'
        }
        else
        {
            $LNKPath = $StartMenuLNKS | Where-Object {$_.Name -like "*$App*"} | Select -first 1
        }
        if($LNKPath)
        {
            log -message ("Found $App at: " + $LNKPath.FullName) -component "Get-TaskBarItems" -type Info
            #Configure OfficeApplicationsCTLC
            $TBI = New-Object TaskBarDesktopApp
            $TBI.New($LNKPath.FullName)
            $ReturnTBC.Add($TBI.XML)
        }
        else
        {
            log -message ("Unable to find $App") -component "Get-TaskBarItems" -type Error
        }
    }
    
    Return $ReturnTBC
}

Function Get-StartMenuGroups($SMGXML)
{
    Log -message ("Attempting to create start menu groups for: " + ($SMGXML.SMG.Name -join ","))  -component "Get-StartMenuGroups" -type 1
    $SMGArray = @()
    $StartMenuLocation = "$ENV:ProgramData\Microsoft\Windows\Start Menu\Programs"
    $StartMenuLNKS = GCI -LiteralPath $StartMenuLocation -Recurse | Where-Object {$_.Extension -like ".lnk"}
    Try
    {
        
        Foreach($SMG in $SMGXML.SMG)
        {
            #Write-Debug "Attempting to create SMG" -Debug
            $MaxCol = 4
            $CurCol = 0
            $CurRow = 0
            Log -message ("Creating Start menu Group:" + $SMG.Name) -component "Get-StartMenuGroups" -type 1
            $CurrentSMG = New-Object StartMenuGroup
            $CurrentSMG.Name = $SMG.Name
            Foreach($App in $($SMG.Apps -split ','))
            {
                Log -message ($CurrentSMG.Name + " - Looking for: $App") -component "Get-StartMenuGroups" -type 1
                $LNKPath = $StartMenuLNKS | Where-Object {$_.name -like "*$App*"} | Select -first 1
                if($LNKPath)
                {
                    Log -message ($CurrentSMG.Name + " - Found: $App at: " + $LNKPath.FullName) -component "Get-StartMenuGroups" -type 1
                    $DAT = New-Object DesktopApplicationTile
                    $Dat.New("2x2", $CurCol, $CurRow, $LNKPath.FullName)
                    $CurrentSMG.Add($Dat.DATCompiledXML)

                    if($CurCol -lt $MaxCol)
                    {
                        $CurCol += 2
                    }
                    else
                    {
                        $CurCol = 0
                        $CurRow += 2
                    }
                    $CurrentSMG.NumberOfApps++
                }
            }
            if($CurrentSMG.NumberOfApps -gt 0)
            {
                $CurrentSMG.GenXML()
                $SMGArray += $CurrentSMG
            }
            else
            {
                Log -message ($SMG.Name + " contains no applications, skipping it!") -component "Get-StartMenuGroups" -type 3
            }
        }
    }
    catch{
        $line = $_.Exception.InvocationInfo.ScriptLineNumber
        Log -message ("Error on line $Line : " + $_.Exception.Message) -component "Get-StartMenuGroups" -type 3
    }
    Return $SMGArray
}


#endregion

#region MainScript

#### Log out our current settings ####

$Script:Settings = New-Object Settings
$Script:Settings.New($xmlFiletoParse)    

Log -message $("Company Name set to:" + $script:Settings.CompanyName) -component "Base Script" -type 1
Log -message $("Clean Start Layout set to: " + $script:Settings.CleanStartLayout) -component "Base Script" -type 1
Log -message $("Apps To Pin: " + $script:Settings.AppsToPin) -component "Base Script" -type 1
$SMGLog = Foreach($SMG in $script:Settings.StartMenuGroups.SMG){("SMGName: " + $SMG.Name); Foreach($App in $SMG.Apps -split ","){"AppName: $App"}}
Log -message $SMGLog -component "Base Script" -type 1

$OutputLocation = $("$env:SystemDrive\Windows\" + $Script:Settings.CompanyName)

if(!(Test-Path $OutputLocation))
{
    New-Item -Path "$env:SystemDrive\Windows" -Name $Script:Settings.CompanyName -ItemType Directory
}

#region Setup our scripts folder under our company folder
    Copy-item $script:ScriptsFolderPath -Destination $OutputLocation -Recurse
#endregion

$OutputLocation = $OutputLocation + "\LayoutModification.xml"

if($Script:Settings.CleanStartLayout -eq $False){
    Log -message "Attempting to create our SMGs" -component "Base Script" -type 1
    #### Generate our SMGS based on our xml config settings in Default_OS_Configure ####
    $SMGS = Get-StartMenuGroups -SMGXML $($Script:Settings.StartMenuGroups)
    
    #### Create our Start Layout #####
    $LayoutModification = New-Object LayoutModification
    
    Foreach($SMG in $SMGS)
    {
        $LayoutModification.Add($SMG.SMGXML)
    }

    $LayoutModification.GenXML($OutputLocation)
}
else
{
    $LayoutModification = New-Object LayoutModification
    Write-Output "Saving XML to $OutputLocation"
    $LayoutModification.GenCleanXML($OutputLocation)
    
}


#### Configure Default Layout Script

if($Script:Settings.AppsToPin)
{
    
    ### Create Second Layout ###

    $ReturnedTBI = Get-TaskBarItems -AppNames $($script:Settings.AppsToPin -split ',')

    $LayoutModification.AddTBI($ReturnedTBI.XML)
    $LayoutModification.GenXML($OutputLocation.Replace(".xml", "_TaskBar.xml"))

    Import-StartLayout -LayoutPath $($OutputLocation.Replace(".xml", "_TaskBar.xml")) -MountPath $($env:SystemDrive + "\")

    
    #region Create Logon Script to disable locked down taskbar
    ##### Check to see if runonce key exists #####
    
    reg load hklm\test C:\users\default\NTUSER.DAT

    $ScriptBlock = {
    param(
        $Script:CompanyName
        )
    $CreateRegKey = 0
    $CorrectTaskBarCommand = "Powershell.exe -executionpolicy bypass -file C:\Windows\$Script:CompanyName\Scripts\CorrectTaskBar.ps1 -companyname $CompanyName"
    $DefaultRunOncePath = "Registry::HKEY_Local_Machine\Test\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    $DefaultRunOnceKey = Get-Item -LiteralPath $DefaultRunOncePath -ErrorAction SilentlyContinue
    if(!$DefaultRunOnceKey)
    {
        New-Item -Path "Registry::HKEY_Local_Machine\Test\Software\Microsoft\Windows\CurrentVersion\RunOnce"
        $CreateRegKey = 1
    }
    else
    {
        ### Path already exists so does our run once for our powershell exist? ###
        $Property = Get-ItemProperty -Path $DefaultRunOncePath
        if($Property.Count -gt 0){
            if(!($Property -match $RunCommand))
            {
                $CreateRegKey = 1
            }
        }
        else
        {
            $CreateRegKey = 1
        }
    }

    if($CreateRegKey -eq 1)
    {
        New-ItemProperty -Path $DefaultRunOncePath -Name "CorrectTaskbar" -PropertyType String -Value $CorrectTaskBarCommand
    }
    }
    
    #Using start job here with the script block to stop the registry for not being able to unload hklm\test

    Start-Job -ScriptBlock $ScriptBlock -Name "Update Registry" -ArgumentList $script:Settings.CompanyName | Wait-Job

    Sleep -Seconds 5

    [gc]::Collect()

    reg unload hklm\test
    
    #endregion
}
else
{
    Import-StartLayout -LayoutLiteralPath $OutputLocation -MountLiteralPath "$env:SystemDrive\"
}

#endregion