<#	
	.NOTES
	===========================================================================

	 Created on:   	1/27/2016 
	 Created by:   	Vadim Smal
  	
	===========================================================================
	.DESCRIPTION
		Modify Folder Structure inorder to meet existing folder structure and permissions
#>


function init
{	
	Clear-Host
	
	$db = 1 #debug off flag
	
	#Prompt for year
	$year = Read-Host 'Please Enter Year [2016]'
	#prompt for month
	$month = Read-Host 'Please Enter Month [2016-01-01]'
	
	#prompt for appended test
	#prompt for location
	$tempFolderLocation = Read-Host "Folder Path of Temp Location of Reports. Don't put in Mexico's SAP or XFER location."
	
	
	$log = $(get-date -f yyyy-MM-dd)
	$log = "$tempFolderLocation\LOG $year-$month.txt"
	

	
	#get Region folders
	$regionFolders = Get-RegionFolders -SharePath $tempFolderLocation
	
	foreach ($region in $regionFolders)
	{
		#get Region number
		$regionNum = Get-RegionNumber -RegionName $region
		$regionReportsLocation = $region.FullName + "\Region reports"
		$regionTotReportsLocation = $region.FullName + "\Total Region " + [string]$regionNum
		$regionTotYear = $region.FullName + "\Region reports" + "\$year"
		$regionTotMonth = $region.FullName + "\Region reports" + "\$year\$month"

		
		
		## Do Total Region Work
		#get Children of Region reports		
		if ((Test-Path $regionReportsLocation) -and (!(Test-Path $regionTotReportsLocation)))
		{
			$regionReportsContents = Get-ChildItem $regionReportsLocation
			Set-Logging -type 1 -m "`tFound Region Reports contents: $regionReportsContents" -deb $db
			
			#create year folder
			New-FolderCreator -path $regionTotYear
			#create month folder
			New-FolderCreator -path $regionTotMonth
			#move contents of  Region Reports to $month folder
			foreach ($folder in $regionReportsContents)
			{				
				$source = $folder.FullName
				$dest = "$regionTotMonth\$folder"
				Set-Robocopy -source $source -dest $dest -log "$log"			
			}
		}		
		if (!(Test-path $regionTotReportsLocation))
		{
			#rename	$regionReportsLocation to $regionTotReportsLocation
			Set-Logging -type 1 -m "`tRenaming : $regionReportsLocation to Total Region $regionNum" 
			rename-item -path $regionReportsLocation -newname $regionTotReportsLocation
		}
		
		
		
		#Do Division Work		
		#Get Divisions Folder
		$divisionFolders = Get-DivisionFolders -RegionPath $region.FullName
		
		foreach ($division in $divisionFolders)
		{
			
			$divFull = $division.Fullname
			Set-Logging -type 1 -m "`tWorking on Division : $divFull" -deb $db
			
			$divYear = $division.Fullname + "\$year"
			$divMonth = $division.Fullname + "\$year\$month"
			$divNum = Get-DivisionNumber -DivisionName $division
			
			$divisionContents = Get-ChildItem $divFull
			
			
			
			if (!(Test-Path $divMonth))
			{
				New-FolderCreator -path $divYear
				New-FolderCreator -path $divMonth
				
				foreach ($folder in $divisionContents)
				{
					$source = $folder.FullName
					$dest = "$divMonth\$folder"
					Set-Logging -type 1 -m "`* Moving : $source *TO* Total Region $dest"
					Set-Robocopy -source $source -dest $dest -log "$log"
				}
			}
			
			#Start Doing Division Reports moves
			$divReport = $region.FullName + "\Division reports"
			$divReportMov = $divMonth + "\Division reports"
			if ((Test-Path $divReport) -and(!(Test-Path $divReportMov)))
			{
				Set-Logging -type 1 -m "`t`t`t`tStarting to work on Division Reports  : $divReport" -deb $db
				$divisionReportsContents = Get-ChildItem $divReport
				Set-Logging -type 1 -m "`tFound following contents in Division Reports  : $divisionReportsContents" -deb $db
				
				foreach ($folder in $divisionReportsContents)
				{
					$source = Find-DivReport -DivisionNumber $divNum -Path $folder.Fullname
					$dest = "$divMonth\Division Reports\$folder"
					Set-Logging -type 1 -m "`* Moving : $source **TO** Total Region $dest"
					Set-Robocopycopy -source $source -dest $dest -log "$log"
				}
			}
			
			
		}
		
		
		#rename and move Division Reports
		$divReport = $region.FullName + "\Division reports"
		
		$regionDivYear = $region.FullName + "\$year"
		$regionDivMonth = $region.FullName + "\$year\$month"

		
		New-FolderCreator -path $regionDivYear
		New-FolderCreator -path $regionDivMonth	
		
		Set-Robocopy -source $divReport -dest $regionDivMonth -log "$log"
		
		

	}
	
	$wshell = New-Object -ComObject Wscript.Shell	
	$wshell.Popup("Operation Completed", 0, "Done", 0x1)
}

## Find appropriate Div folder to move from 
function Find-DivReport
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$DivisionNumber,
		[Parameter(Mandatory = $true)]
		[string]$Path
	)
	
	Set-Logging -type 1 -m "`tPath  : $Path" -deb $db
	Set-Logging -type 1 -m "`tDiv Number  : $DivisionNumber" -deb $db
	
	$searchTag = "D" + $DivisionNumber
	Set-Logging -type 1 -m "`tSearching For  : $searchTag" -deb $db
	
	$centerRpts = Get-ChildItem $Path
	foreach ($folder in $centerRpts)
	{
		$name = [string]$folder
		if ($name.length -ge ([int]$searchTag.length))
		{
			$sub3 = $name.subString($name.length - [int]$searchTag.length, [int]$searchTag.length)
			if ($sub3 -like "*$searchTag*")
			{
				Set-Logging -type 1 -m "`tFound Tag  : $sub3" -deb $db
				$Path = $Path + "\$name"
				Set-Logging -type 1 -m "`tFound Folder to Move  : $Path" -deb $db
			}
		}
		
	}
	
	return $path
}

## Copy with deleting source
function Set-Robocopy
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$source,
		[Parameter(Mandatory = $true)]
		[string]$dest,
		[Parameter(Mandatory = $true)]
		[string]$log
	)
	
	Set-Logging -type 1 -m "`tSource moving: $source" -deb $db
	Set-Logging -type 1 -m "`tDestination Move: $dest" -deb $db
	$flags = '/E /MOVE /R:3 /W:5 /log+:"' + $log + '"'
	$cmd = 'ROBOCOPY "' + $source + '" "' + $dest + '" ' + $flags
	invoke-expression $cmd
}

## Copy without deleting source
function Set-Robocopycopy
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$source,
		[Parameter(Mandatory = $true)]
		[string]$dest,
		[Parameter(Mandatory = $true)]
		[string]$log
	)
	
	Set-Logging -type 1 -m "`tSource moving: $source" -deb $db
	Set-Logging -type 1 -m "`tDestination Move: $dest" -deb $db
	$flags = '/E /R:3 /W:5 /log+:"' + $log + '"'
	$cmd = 'ROBOCOPY "' + $source + '" "' + $dest + '" ' + $flags
	invoke-expression $cmd
}

## create folder if not exist, given fullpath
function New-FolderCreator
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$path
	)
	
	$c = 0
	if (Test-Path $path)
	{
		Set-Logging -type 1 -m "`tSkipping, folder already exists: $path" -deb $db
	}
	else
	{		
		Set-Logging -type 1 -m "`tCreating folder: $path"
		$ret = New-Item -Path $path -ItemType directory
		$c = 1
	}
	
	#if running creator, recheck to make sure folder was created successfully.
	if ($c)
	{
		if ((Test-Path $path))
		{
			Set-Logging -type 1 -m "`tChecking, folder creation successfull" -deb $db
		}
		else
		{
			Set-Logging -type 1 -m "`tError, unable to successfully create folder - EXITING"
			Exit
		}
	}
	
}


## Given sharePath, return folders that Start with Division but exclude Division reports
function Get-DivisionFolders
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$RegionPath
	)
	
	Set-Logging -type 1 -m "Getting Divisions from: $RegionPath"
	
	$folderList = Get-ChildItem $RegionPath -attributes D | ?{ $_.fullname -notmatch "\\Division reports\\?" }
	$returnList = @()
	foreach ($folder in $folderList)
	{
		if ($folder.Name -like "Division*")
		{
			$returnList = $returnList += $folder
			Set-Logging -type 1 -m "`tFound Division folder called: $folder" -deb $db
		}
	}
	return $returnList
}

## Given Region name, make sure matches 'Region ####' format and then return ###
function Get-DivisionNumber
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$DivisionName
	)
	
	$num = 0
	
	Set-Logging -type 1 -m "Starting Get-DivisionNumber Function" -Debug $db
	
	if ($DivisionName -like "*Division*")
	{
		Set-Logging -type 1 -m "`tGetting Division Number from: $RegionName" -Debug $db
		$num = $DivisionName.substring(9, $DivisionName.length - 9)
		Set-Logging -type 1 -m "`tDivision Number Found: $num" -Debug $db
	}
	
	return $num
}

## Given sharePath, return folders that Start with Region
function Get-RegionFolders
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$SharePath
	)

	Set-Logging -type 1 -m "Getting Regions from folder path: $SharePath" 
	
	$folderList = Get-ChildItem $SharePath -attributes D	
	$returnList = @()	
	foreach ($folder in $folderList)
	{
		if ($folder.Name -like "Region*")
		{
			$returnList = $returnList += $folder
			Set-Logging -type 1 -m "`tFound Region folder called: $folder" 
		}		
	}	
	return $returnList
}

## Given Region name, make sure matches 'Region ####' format and then return ###
function Get-RegionNumber
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$RegionName
	)
	
	$num = 0
	
	Set-Logging -type 1 -m "Starting Get-RegionNumber Function" -Debug $db
	
	if ($RegionName -like "*Region*")
	{
		Set-Logging -type 1 -m "`tGetting Region Number from: $RegionName" -Debug $db
		$num = $RegionName.substring(7, $RegionName.length - 7)
		Set-Logging -type 1 -m "`tRegion Number Found: $num" -Debug $db
	}
	
	return $num	
}

## Logging options
function Set-Logging
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$m,
		[int]$type = 0,
		[int]$deb = 0
	)
	
	$msg = ''
	if ($deb -eq 0)
	{
		if ($type -eq 0)
		{
			$msg = $m
		}
		if ($type -eq 1)
		{
			$msg = "`t	" + $m
		}
		
		
		Write-Host ($msg)
	}
	
	
}



#main runner
init