<# 
IDFix-style tool for Public Folder MailNicknames

THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT WARRANTY 
OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE 
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF 
THIS CODE REMAINS WITH THE USER.

Comments or suggestions to aaron.guilmette@microsoft.com
#>
 
<#
.SYNOPSIS
 Checks Public Folder nicknames for invalid characters prior to migration to Office 365.
.EXAMPLE
 .\Check-PublicFolderMailandAliases.ps1
.DESCRIPTION
 Generate input file from Exchange Management Shell using command:
 
 Get-Recipient -ResultSize Unlimited -RecipientTypeDetails PublicFolder | Select Alias,DisplayName,Name,RecipientType,RecipientTypeDetails,@{Name='EmailAddresses';
 Expression={[string]::join(";",($_.EmailAddresses))}} | Export-Csv -NoType PublicFolders.csv -Encoding UTF8
 
 Output file will have columns labelled:
 OriginalAlias,SuggestedAlias,SMTP,SmtpIsBad
 
 - Original Alias is the original Public Folder Alias.
 - SuggestedAlias is a suggestion for updating the Alias based on valid characters.
 - SMTP is the primary SMTP address of the public folder.
 - SmtpIsBad indicates that the primary SMTP address failed Net.Mail.MailAddress validation.
.LINK https://gallery.technet.microsoft.com/IDFix-for-Public-Folders-341522d6
#>

# Set Console Size
$pshost = get-host
$pswindow = $pshost.ui.rawui

$newsize = $pswindow.buffersize
$newsize.height = 5000
$newsize.width = 150
$pswindow.buffersize = $newsize

$newsize = $pswindow.windowsize
$newsize.height = 50
$newsize.width = 150
$pswindow.windowsize = $newsize

 
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
Function Get-FileName($initialDirectory)
	{   
     [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	 $OpenFileDialog.initialDirectory = $initialDirectory
	 $OpenFileDialog.filter = "All files (*.*)|*.*"
	 $OpenFileDialog.ShowHelp = $true
	 $OpenFileDialog.ShowDialog() | Out-Null
	 $OpenFileDialog.filename
	} 
 
$inputFile = (Get-FileName -initialDirectory "C:")
$tempFile = Import-Csv $inputFile
$PFs = $tempFile | Where { $_.RecipientTypeDetails -like "PublicFolder" }
$UpdatedPFs = [Microsoft.VisualBasic.Interaction]::InputBox("Enter output filename:", "Output Filename", "SuggestedPFs.csv") 
 
<# Rules for MailNickname per IDFix Guidelines
invalid chars whitespace \ ! # $ % & * + / = ? ^ ` { } | ~ < > ( ) ' ; : , [ ] " @
may not begin or end with a period
less than 64
#>
 
# Build CSV Header
$header = "OriginalAlias,SuggestedAlias,SMTP,SmtpIsBad"
$header | Out-File -Append $UpdatedPFs
 
# Run aliases through rules
foreach ($folder in $PFs) { 
	$SmtpIsBad = $null
	$a = $folder.Alias                                                            													# Set $a to current folder.Alias
	$address = $folder.EmailAddresses.Split(";") | where { $_ -clike "SMTP:*" } 													# Select Primary SMTP Address
	If ($address -eq $null) 
		{ 
		Write-host -ForegroundColor Red "Error: $a does not have a valid primary SMTP address"                   					# Except when it doesn't exist
		$address = $folder.EmailAddresses.Split(";") | where { $_ -like "smtp:*" }                               					# Select another address
		$address = $address[0]                                                                                   					# In case there are more than one
		Write-Host -ForegroundColor DarkYellow "       Using $address as Primary SMTP instead"                   					# Report what we're using to identify the folder
		Write-Host -ForegroundColor DarkGray "--------------------------------------------------------------"    					# Row separator
		} 
	$address = $address.Substring(5)                                                                               					# Remove leading "SMTP:" value
	Write-Progress -Activity "Processing Public Folder Aliases" -Status "Working on $a" -PercentComplete ($i++ / $PFs.Count * 100)
	
	# Check for valid primary SMTP address
	try { 
		$to = New-Object Net.Mail.MailAddress($address)
		} 
	catch [system.exception] {
		write-host -ForegroundColor Red "$address has a malformed SMTP address."
		$SmtpIsBad = "Yes"
		} 
	finally {
		#
		}
	$a = $a.Replace(" ","")                                                        # Remove spaces " "
	$a = $a.Replace("\","")                                                        # Remove backslash "\"
	$a = $a.Replace("!","")                                                        # Remove exclamation points "!"
	$a = $a.Replace("#","")                                                        # Remove number signs "#"
	$a = $a.Replace("$","")                                                        # Remove dollar signs "$"
	$a = $a.Replace("%","")                                                        # Remove percent signs "%"
	$a = $a.Replace("&","")                                                        # Remove ampersands "&"
	$a = $a.Replace("*","")                                                        # Remove asterisks "*"
	$a = $a.Replace("+","")                                                        # Remove plus signs "+"
	$a = $a.Replace("/","")                                                        # Remove forward slashes "/"
	$a = $a.Replace("=","")                                                        # Remove equals signs "="
	$a = $a.Replace("?","")                                                        # Remove question marks "?"
	$a = $a.Replace("^","")                                                        # Remove carats "^"
	$a = $a.Replace("``","")                                                       # Remove back ticks "`"
	$a = $a.Replace("{","")                                                        # Remove left brace "{"
	$a = $a.Replace("}","")                                                        # Remove right brace "}"
	$a = $a.Replace("|","")                                                        # Remove pipe symbol "|"
	$a = $a.Replace("~","")                                                        # Remove tilde "~"
	$a = $a.Replace("<","")                                                        # Remove less than symbol "<"
	$a = $a.Replace(">","")                                                        # Remove greater than symbol ">"
	$a = $a.Replace("(","")                                                        # Remove left parenthesis "("
	$a = $a.Replace(")","")                                                        # Remove right parenthesis ")"
	$a = $a.Replace("'","")                                                        # Remove apostrophe "'"
	$a = $a.Replace(";","")                                                        # Remove semicolon ";"
	$a = $a.Replace(":","")                                                        # Remove colon ":"
	$a = $a.Replace(",","")                                                        # Remove comma ","
	$a = $a.Replace("[","")                                                        # Remove left square bracket "["
	$a = $a.Replace("]","")                                                        # Remove right square bracket "]"
	$a = $a.Replace("""","")                                                       # Remove quotation mark """
	$a = $a.Replace("@","")                                                        # Remove at symbol "@"
	$a = $a.Trim(".")                                                              # Remove leading and trailing period "."
	If ($a.Length -gt 60) {$a = $a.Substring(0,60)}                                # Trim to 60 characters to accommodate random number suffix
	If ($a -notlike $folder.Alias)                                                 # Add PFs that need to be updated to Output file $UpdatedPFs
		{ 
			$a = $a + (Get-Random -Minimum 101 -Maximum 999)                                                                                # Generate Random number to append
			$data = """" + $folder.Alias + """" + "," + """" + $a + """" + "," + """" + $address + """"    + "," + """" + $SmtpIsBad + """" # Build CSV row data
			$data | Out-File -Append $UpdatedPFs                                                                                            # Write output
			$SmtpIsBad = $null
		}
	If ($SmtpIsBad -ne $null)
		{
			$data = """" + $folder.Alias + """" + "," + """" + " " + """" + "," + """" + $address + """" + "," + """" + $SmtpIsBad + """"   # SMTP needs to be fixed
			$data | Out-File -Append $UpdatedPFs
			$SmtpIsBad = $null
		}
	}
 
Write-Host -ForegroundColor Green "All done. Output file is $UpdatedPFs."