<#
SMTPReport.ps1
Created By - Kristopher Roy
Ceated On - 11/29/2017
Last Modified - 12/5/2017
Purpose - The purpose of this script is to loop through a list of users and create a SMTP report for them, initially based upon csv list, but eventually will include options to manually input users, or grab and report on all!
#>

#import the AD Module
Import-Module ActiveDirectory

#Function to select a file
function Get-FileName
{
  param(
      [Parameter(Mandatory=$false)]
      [string] $Filter,
      [Parameter(Mandatory=$false)]
      [switch]$Obj,
      [Parameter(Mandatory=$False)]
      [string]$Title = "Select A File"
    )
 
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.FileName = $Title
  #can be set to filter file types
  IF($Filter -ne $null){
  $FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
	$OpenFileDialog.filter = $FilterString}
  if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
  $OpenFileDialog.filter = $Filter
  }
  $OpenFileDialog.ShowDialog() | Out-Null
  ## dont bother asking, just give back the object
  IF($OBJ){
  $fileobject = GI -Path $OpenFileDialog.FileName.tostring()
  Return $fileObject
  }
  else{Return $OpenFileDialog.FileName}
}

<#
	This Function creates a dialogue to return a Folder Path
#>
function Get-Folder {
    param([string]$Description="Select Folder to place results in",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

$file = get-filename -Filter "csv" -Title "Select User List" -obj
$userlist = import-csv $file.FullName

#Create the Array to store the data
$smtparray = @()

$progresscount = 0

#loop through each user in the list
FOREACH($user in $userlist)
{
	$progresscount++
	Write-Progress -Activity ("Gathering SMTP Addresses . . ."+$user.EmailAddress) -Status "Scanned: $progresscount of $($userlist.emailaddress.Count)" -PercentComplete ($progresscount/$userlist.emailaddress.Count*100)
	
	#convert email to filterable data
	$upnpre = (($user.emailaddress).split('@'))[0]	
	#$username = $user.emailaddress.split("@")[0]
	#$email = $userlist.emailaddress
	
	#get ADUser Object
	$aduser = Get-ADUser -Filter "UserPrincipalName -like '$upnpre*'" -Properties proxyaddresses|Select Name,sAMAccountName,ProxyAddresses,UserPrincipalName
	
	#Create SMTP details to add to report
	$smtpobject = new-object PSObject
	
	#add username to the smtp object report
	$smtpobject|add-member -membertype NoteProperty -name "user" -value $aduser.name
	$smtpobject|add-member -membertype NoteProperty -name "sAMAccountName" -value $aduser.sAMAccountName
	$smtpobject|add-member -membertype NoteProperty -name "UserPrincipalName" -value $aduser.UserPrincipalName
	$count = 0
	
	#loop through each smtp and add them to the report
	FOREACH($address in $aduser.proxyaddresses)
	{
		$count++
		$smtpobject|add-member -membertype NoteProperty -name "smtp0$count" -value $address
	}
	$smtparray += $smtpobject
	$smtpobject = $null
	$count = $null
	$user = $null
}
#Date for timestamp
$date = Get-Date -Format yyyy_MM_dd-HHmm

#runs the get-folder function to grab location for output then names the output file and exports it 
$smtparray|export-csv ((Get-Folder)+"\"+$date+"_SMTPReport_"+($file.Name)) -NoTypeInformation