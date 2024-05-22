$MailboxOwner = Caroline Wilson
$DelegateUser = 
$ForwardingUser = 
$forwardingmailbox = get-mailbox -identity "Stephanie Couldwell"
$ForwardingAddress = $forwardingmailbox.PrimarySmtpAddress.ToString()
$username = $MailboxOwner -split " "
Add-MailboxPermission -Identity "Caroline Wilson" -User "Stephanie Couldwell" -AccessRights FullAccess
Set-Mailbox -Identity "Caroline Wilson" -ForwardingAddress $ForwardingAddress -DeliverToMailboxAndForward $true
Get-Mailbox "Caroline Wilson" | Format-List ForwardingSMTPAddress,DeliverToMailboxandForward
$oofMessage = "Please be advised Caroline Wilson is no longer at Devon Properties. `nFor assistance, please reach out to administration@devonproperties.com.`nThank you, "
# Set the out-of-office message
Set-MailboxAutoReplyConfiguration -Identity "Caroline Wilson" -AutoReplyState Enabled -ExternalMessage $oofMessage -InternalMessage $oofMessage