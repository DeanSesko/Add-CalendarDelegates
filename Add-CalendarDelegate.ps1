
#########################################################################################
# COMPANY: CDW								                                            #
# NAME: Add-CalendarDelegates.ps1                                                       #
#                                                                                       #
# AUTHOR:  Dean Sesko                                                                   #
#                                                                                       #
# DATE:  03/01/2015                                                                     #
# EMAIL: Dean.Sesko@S3.CDW.com                                                          #
#                                                                                       #
# COMMENT:  Script to Add Delegates to all memebers of a group   	                    #
#                                                                                       #
# VERSION HISTORY                                                                       #
# 1.0 03/01/2015  Initial Version.                                                      #
#                                                                                       #
#########################################################################################
$Revieweruser = "Test@Contoso.com"
$DLGroup = "MyGroup"
$Directors = Get-DistributionGroup $DLGroup | Get-DistributionGroupMember | Select SamAccountName

function Validate-IsEmail {
	[OutputType([Boolean])]
	param ([string]$Email)
	
	return $Email -match "^(?("")("".+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$"
}


if (Validate-IsEmail $Revieweruser) {
	foreach ($user in $Directors.SamAccountName) {
		if ($user -ne "") {
			
			[String]$MailboxFolder = $user + ":\Calendar"
			
			Try {
				add-MailboxFolderPermission  $MailboxFolder -AccessRights Reviewer -User $Revieweruser  -ErrorAction 'Continue'
			}
			catch {
				[String]$ErrorMessage = $_.Exception
				
				return $false
			}
			
			Finally {
				switch -regex ($ErrorMessage) {
					"An existing permission entry was found for user:" { Write-host " $Revieweruser Already has Access o $MailboxFolder " }
					
					
				}
				
			}
		}
	}
}
Else {
	Write-Host "Please Enter a Vailid SMTP Address"
}




