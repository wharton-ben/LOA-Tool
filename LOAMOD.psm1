<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.187
	 Created on:   	6/6/2021 11:14 PM
	 Created by:   	Ben
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>


<#
.Synopsis
   This function will set a user to LOA status. 
.DESCRIPTION
    Set-LOAUser will set the user's AD account to LOA status. This includes modifying the user's AD description, logon workstations and group memberships.
    Because the user is only allowed to access the mail servers while on LOA, they will need to have the mail servers added explicitly to their 
    "logonworkstation" property. The user will also be added to the AD group "LOA Accounts". This will manage the users permissions futher. 
#>

function Set-LOAUser
{
	# This will open a file dialog window so the user can select the list of users. 
	
	$initialdirectory = 'C:\'
	$openfiledialog = new-object system.windows.forms.openfiledialog
	$openfiledialog.initialdirectory = $initialdirectory
	$openfiledialog.filter = "CSV (*.csv)| *.csv"
	$openfiledialog.showdialog() | out-null
	$openfiledialog.filename
	
	# Import the CSV that the user selected into the $users variable
	
	$users = Import-Csv -path $openfiledialog.filename
	
	# Count the amount of users to be set to LOA status
	
	$count = ($users | Measure-Object).Count
	
	$note = New-Object -ComObject wscript.Shell
	
	if ($users -eq $null)
	{
		$nulluser = $note.popup("You did not select a csv file, no users will be created. Please select 'okay' to select a file.", 0, "Notice", 1 + 32)
		if ($nulluser -eq 1)
		{
			Set-LOAUser
			break
		}
		else
		{
			break
		}
	}
	
	$selection = $note.popup("Are you sure you want to set $($count) users to LOA status?", 0, "Notification", 1 + 32)
	
	if ($selection -eq 1)
	{
		
            <# A message will be displayed for the user asking to confirm if they would like to modify the account.
            If they select yes, the users will be run through the script block below, otherwise the script will stop.
            #>
		
		foreach ($user in $users)
		{
			
			# This will set the users description and logon workstations to the organizations LOA specifications 
			
			Set-ADUser -identity $user.'employee' -Description "LOA" -LogonWorkstations #Enter the name of the workstations you would like the user to have access to here.
			
			# Next the user will need to be removed from the LOA Accounts group
			
			Add-ADGroupMember -Identity "LOA Accounts" -Members $user.'employee' -Confirm:$false
			
			# Gather details and export to csv
			
			$details = [ordered]@{
				"Action"   = "Set user to LOA";
				"Username" = $user.employee;
				"Description" = $(Get-ADUser $user.'employee' -Properties * | Select-Object description).description;
				"LogonWorkstations" = $(Get-ADUser $user.'employee' -Properties * | Select-Object Logonworkstations).logonworkstations;
				"Groups"   = $(Get-ADPrincipalGroupMembership $user.'employee' | Where-Object { $_.name -eq "LOA Accounts" }).name
			}
			
			$resultdate = Get-Date -Format "hh-mm-MM-dd-yyyy"
			$resultfile = "C:\LOAresults$($resultdate).csv"
			$results += New-Object System.Management.Automation.PSObject -property $details | Export-Csv $resultfile -NoTypeInformation -force -Append
		}
		
	}
	else
	{
		break
	}
	$note.popup("You have successfully set $count accounts to LOA status.")
}

<#
.Synopsis
   This function will restore a user from LOA status. 
.DESCRIPTION
    Restore-LOAUser will restore the user's AD account from LOA status. This includes modifying the user's AD description, logon workstations and group memberships. 
#>

function Restore-LOAUser
{
	# This will open a file dialog window so the user can select the list of users. 
	
	$initialdirectory = 'C:\'
	$openfiledialog = new-object system.windows.forms.openfiledialog
	$openfiledialog.initialdirectory = $initialdirectory
	$openfiledialog.filter = "CSV (*.csv)| *.csv"
	$openfiledialog.showdialog() | out-null
	$openfiledialog.filename
	
	# Import the CSV that the user selected into the $users variable
	
	$users = Import-Csv -path $openfiledialog.filename
	$count = ($users | Measure-Object).Count
	
	# Create a pop up window to send feedback to users while running the script
	
	$note = New-Object -ComObject wscript.Shell
	
	# If the user runs the program, but does not select a file, the user will be notified and the script will stop running. 
	
	if ($users -eq $null)
	{
		$nulluser = $note.popup("You did not select a csv file, no users will be created. Please select 'okay' to select a file.", 0, "Notice", 1 + 32)
		if ($nulluser -eq 1)
		{
			Restore-LOAUser
			break
		}
		else
		{
			break
		}
	}
	
	# The user will be notified that x amount of accounts are about to be restored from LOA status.
	
	$selection = $note.popup("Are you sure you want to restore $($count) users from LOA status?", 0, "Notification", 1 + 32)
	if ($selection -eq 1)
	{
		
        <# A message will be displayed for the user asking to confirm if they would like to modify the account.
        If they select yes, the users will be run through the script block below, otherwise the script will stop.
        #>
		
		foreach ($user in $users)
		{
			
			# This will set the users description and logon workstations to null. This will allow the user to 
			# log in to any device and remove the LOA status from the description. 
			
			Set-ADUser -identity $user.'employee' -Description $null -LogonWorkstations $null
			
			# Next the user will need to be removed from the LOA Accounts group
			
			Remove-ADGroupMember -Identity 'LOA Accounts' -Members $user.'employee' -Confirm:$false
			
			$details = [ordered]@{
				"Action"   = "Restore from LOA";
				"Username" = $user.employee;
				"Description" = $(Get-ADUser $user.'employee' -Properties * | Select-Object description).description;
				"LogonWorkstations" = $(Get-ADUser $user.'employee' -Properties * | Select-Object Logonworkstations).logonworkstations;
				"Groups"   = $(Get-ADPrincipalGroupMembership $user.'employee' | Where-Object { $_.name -eq "LOA Accounts" }).name
			}
			$resultdate = Get-Date -Format "hh-mm-MM-dd-yyyy"
			$resultfile = "C:\LOAresults$($resultdate).csv"
			$results += New-Object System.Management.Automation.PSObject -property $details | Export-Csv $resultfile -NoTypeInformation -force -Append
		}
		
	}
	else
	{
		$y = $note.popup("Do you want to continue?", 0, "Notification", 1 + 32)
		if ($y -eq 1)
		{
			Restore-LOAUser
		}
		break
	}
	$note.popup("You have successfully restored $($count) accounts from LOA status.")
	
}

<#
.Synopsis
   This function will either set a user to LOA status or restore their account from LOA status. 
.DESCRIPTION
   The Modify-LOAStatus function aims to simplify the tedious task of moving users to and from LOA status. Each day at the end of the day, 
   a report is run for all users who are either returning or going on leave of absense. Each user moving to LOA status will need to have the description on 
   their account moved to "LOA", the logon workstations will need to be set to only the mail servers, and they will need to be added to the "LOA Accounts" 
   group. The users who are returning will need their description set to $null, they will need to be set to be able to log on to any workstation and they 
   will need to be removed from the "LOA Accounts" group. 

   When the function is run, the user will be prompted to choose between setting user to LOA status or setting the user back to normal after returning from 
   LOA status. Depending on their choice, either the Restore-LOAUser or the Set-LOAUser function will be run. 

   The Restore-LOAUser function will set the users' description property to $null. It will also set the logon workstations to $null (this means the user is able
   to log on to all workstations) and it will remove the users' from the "LOA Accounts" group. 

   The Set-LOAUser function will set the users' description to "LOA". It will set the logon workstations to only include the mail servers (LOA users are only 
   allowed to access email). Finally the users will be added to the "LOA Accounts" group. 

.EXAMPLE
   Modify-LOAUser -action restore: 
    - This will restore a list of users from LOA status. 
.EXAMPLE
   Modify-LOAUser -action set: 
    - This will set a list of users to LOA status
#>

function Modify-LOAUser
{
	param ([parameter(mandatory)]
		[ValidateSet("restore", "set")]
		[string]$action
	)
	
	#create shell for pop up window
	
	#$popup = New-Object -ComObject WScript.Shell
	
	Switch ($action)
	{
		"restore"{
			Restore-LOAUser
		}
		"set"{
			Set-LOAUser
		}
	}
}

