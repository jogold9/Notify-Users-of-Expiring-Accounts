####################################################################################################
# Check account expiration date for accounts that will expire soon.  Email user if an email exists. 
# Josh Gold
####################################################################################################

#############
# Functions
#############

# Function checks if email field has values
Function DoesEmailExist
{
   If([string]::IsNullOrEmpty($user.UserPrincipalName)) {            
    Write-Host $User.Name "does not have an email in Active Directory"            
   } else {            
    Write-Host $User.Name "has an email in Active Directory. Let's send an email."
    Send_Email            
   }
        
}

#Function sends emails if needed
$name = $user.Name
$expiring_date = $user.AccountExpirationDate

Function Send_Email{

Write-Host "Sending Email to" $user.Name "whose account expires" $user.AccountExpirationDate "`n"

$subject = "Your account is expiring " + $User.AccountExpirationDate +  ". Please have your manager contact the IT Help Desk"

$body =@"
Hello $name,<br><br>

We noticed your account is expiring soon: $expiring_date <br><br>
 
<b>If your contract is being extended, or you are a current employee, please have your manager call or email the IT Help Desk. </b> <br><br>

<u>Your manager will need to provide your full name <b>AND</b> your account name to the IT Service Desk.</u><br><br>

If applicable, your manager will need to provide the new contract expiration date as well.<br><br>

Thank you, <br><br>

Information Technology Department <br>

__________________________________________________________________ <br>
*This email is intended only for use by addressee(s) named and may contain confidential information. If you are not the intended recipient, or the person responsible for delivering this information to the intended recipient, you are hereby notified that any dissemination, distribution, printing, or copying of this email is strictly prohibited. If you have received this message in error, please notify this office and promptly destroy the original copy of any email.<br>

"@

 $smtpServer = "mailhost.yourserver.or.us"
 $smtp = new-object Net.Mail.SmtpClient($smtpServer)
 $msg = new-object Net.Mail.MailMessage
 $msg.From = ("Support@yourcompany.net")
 #$msg.BCC.Add("itaccounts@yourcompany.net")
 $msg.To.Add("jgold@yourcompany.net")
 #$msg.To.Add($User.UserPrincipalName)
 
 $msg.subject = $subject
 $msg.IsBodyHTML = $true
 $msg.body = $body
 $smtp.Send($msg)

 Write-Host -Foreground Cyan "        Sending email notification to" $User.UserPrincipalName 
 start-sleep 3
 
 Write-Host
 Write-Host  "*****Process complete!" -ForegroundColor Green
}

###########################################
# Functions are done, let's run some code
###########################################

# Get user accounts expiring in the next 30 days, computer accounts are excluded
$expiring_accounts = Search-ADAccount -AccountExpiring -UsersOnly -TimeSpan "30"

# If account is enabled, check if an email address exists, and email user if an address exists
foreach ($user in $expiring_accounts){
    if ($user.Enabled){
    Write-Host $user.Name "is enabled"
    DoesEmailExist
    }
}

