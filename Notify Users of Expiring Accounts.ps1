####################################################
# Check account expiration date for accounts that will expire soon.  Email user if an email exists. 
# Josh Gold
#####################################################

#################################
# Function sends emails if needed
#################################
Function Send_Email{

$username = $user.samaccountname
$expiring_date = $user.AccountExpirationDate

Write-Host "Sending Email to" $user.Name "whose account expires" $user.AccountExpirationDate "`n"

$subject = "Your account will expire on " + $User.AccountExpirationDate 

$body =@"
Hello,<br><br>

We noticed your account is expiring soon: <br>
Username: $username <br>
Expiration date: $expiring_date <br><br>
 
<b>If your contract is being extended, or you are a current employee, please have your manager email support@yourcompany.com or visit https://support.yourcompany.com </b> <br><br>

<u>Your manager will need to provide your full name <b>AND</b> your account name to the IT Service Desk.</u><br><br>

If applicable, your manager will need to provide the new contract expiration date as well.<br><br>

Thank you, <br><br>

Information Technology Department <br>
Your company name <br>
https://support.yourcompany.com <br>
support@yourcompany.com <br><br>

*This email is intended only for use by addressee(s) named and may contain confidential information. If you are not the intended recipient, or the person responsible for delivering this information to the intended recipient, you are hereby notified that any dissemination, distribution, printing, or copying of this email is strictly prohibited. If you have received this message in error, please notify this office and promptly destroy the original copy of any email.<br>
"@

 $smtpServer = "mailhost.yourcompany.com"
 $smtp = new-object Net.Mail.SmtpClient($smtpServer)
 $msg = new-object Net.Mail.MailMessage
 $msg.From = ("accounts@yourcompany.com")
 $msg.To.Add($email)
 #$msg.BCC.Add(""accounts@yourcompany.com")
  
 $msg.subject = $subject
 $msg.IsBodyHTML = $true
 $msg.body = $body
 $smtp.Send($msg)

 Write-Host -Foreground Cyan "        Sending email notification to" $user.userprincipalname 
 start-sleep 3
 
 Write-Host
 Write-Host  "*****Process complete!" -ForegroundColor Green
}

###############################################
# Function checks if an email address is valid
###############################################
Function IsValidEmailAddress {

$IsValidEmailAddress = $True

Try
            {
                New-Object System.Net.Mail.MailAddress($email)
            }
            Catch
            {
                $IsValidEmailAddress = $False
            }
        
        Return $IsValidEmailAddress
}

########################################################
# Function definitions are done now, let's run some code
########################################################

# Get user accounts expiring in the next 30 days, computer accounts and student accounts excluded
$expiring_accounts = Search-ADAccount -AccountExpiring -UsersOnly -TimeSpan "30"

#Run some checks explained in each of the comments below
foreach ($user in $expiring_accounts){
  
  #Let's only worry about enabled accounts
  if ($user.Enabled){
  Write-Host -ForegroundColor Cyan $user.Name "($user.Name) is an enabled account."
     
     #Let's only worry about accounts that are not living in student OUs
     if ($user.DistinguisedName -notcontains "OU=student"){
        Write-Host -ForegroundColor green "Account appears to be employee or contractor, not a student, so let's see if we can send some emails."
            
            #Let's get the email address and make sure it is valid
            $email = (get-aduser -Identity $user.SID -Properties emailaddress).emailaddress  
                if (IsValidEmailAddress){
                  Write-Host -ForegroundColor Yellow $email "appears to be a valid email and we will send a message." `r`n
                  Send_Email
                }
                else {
                  Write-Host -ForegroundColor Red $email " does NOT appear to be a valid email address." `r`n
                }
     }
  }    
}

