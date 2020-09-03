[CmdletBinding()]
param()

Set-ExecutionPolicy RemoteSigned -force

#This script imports a csv file and creates new users in O365, adds licensing (depending on grade level), creates an external contact (parent) for each new user, and sets up email forwarding on each student account (to forward emails to external parent email address).

# verbose and debug settings #
$resetdebugpreference = $debugpreference
$debugpreference = "Inquire"
$resetverbose = $verbosepreference
$verbosepreference = "Continue"

# set folder location for running scripts in this file #
set-location "C:\PathToFolder\Scripts"

# import file for O365 user creation #
$filename = (gci "C:\PathToFolder\Enrollments\" | sort LastWriteTime | select fullname -last 1) 

$newstu = import-csv $filename.fullname
$stuContent = get-content $filename.fullname

# set default dates for O365 export files #
$o365date = (get-date).tostring('MM-dd-yy')

# set default export file to use for re-importing steps later in this script
$o365export = "C:\PathToFolder\Enrollments\exp\O365Export_$o365date.csv"



# Function that creates ALL the export and re-import files for enrollment process #
function create-files {
    # creating O365 export file #
    function test-export {
        $int = $null
        $test = test-path $script:o365export
        If ($test -eq $True) {
            foreach ($t in $test) {
                $int = ++$int
                $script:o365date = (get-date).tostring('MM-dd-yy') + '_' + $int
                $script:o365export = "C:\PathToFolder\Enrollments\exp\O365Export_$o365date.csv"
                $test = test-path $script:o365export
                test-path $script:o365export
            }
        }
    }

    test-export -debug
    write-verbose "Export filename created"
}

# You can use this area to add any dependency scripts; useful if you need to create import files for other systems.
function run-scripts {
    # Load Adobe Connect file import script
    # Uncomment the line below and add your own scripts!
    # invoke-expression -command ./external-sript-sample.ps1

}

create-files
run-scripts

# Set student o365 licenses - all plans are enabled by default, $LO and $LOK5 list the licenses that will NOT be included with eah account.
$LO = New-MsolLicenseOptions -AccountSkuId "ExampleCompany:STANDARDWOFFPACK_IW_STUDENT" -DisabledPlans TEAMS1, Deskless, FLOW_O365_P2, POWERAPPS_O365_P2, OFFICE_FORMS_PLAN_2, PROJECTWORKMANAGEMENT, SWAY, YAMMER_EDU, SHAREPOINTWAC_EDU, SHAREPOINTSTANDARD_EDU, MCOSTANDARD, KAIZALA_O365_P2

# This set of license options disables Exchange email - users can still download Microsoft office by visitng the URL "https://portal.office.com/account/""
$LOK5 = New-MsolLicenseOptions -AccountSkuId "ExampleCompany:STANDARDWOFFPACK_IW_STUDENT" -DisabledPlans TEAMS1, Deskless, FLOW_O365_P2, POWERAPPS_O365_P2, OFFICE_FORMS_PLAN_2, PROJECTWORKMANAGEMENT, SWAY, YAMMER_EDU, SHAREPOINTWAC_EDU, SHAREPOINTSTANDARD_EDU, MCOSTANDARD, KAIZALA_O365_P2, EXCHANGE_S_STANDARD

# create new o365 user accounts #
write-verbose "Process starting, will take a few minutes..."

# Creates O365 users based on the data from the Excel files. You can change the $_. variables to the header row from your csv.
$newstu | ForEach-object {
   
    # If student is grade 6 or above, create account with email access
    if ([int]$_.grade -gt 5) {
        New-MsolUser -UserPrincipalName $_.Stu_Email -password $_.Initial_PW -FirstName $_.First_Name -LastName $_.Last_Name -DisplayName $_.Display_Name -UsageLocation $_.Country -LicenseAssignment "ExampleCompany:STANDARDWOFFPACK_IW_STUDENT" -LicenseOptions $LO -erroraction silentlycontinue
        }

     # If student grades are under 5, create licensing that does not include email access (but they still have an account for other platorms' SSO purposes) 
    if ([int]$_.grade -lt 6) {
        New-MsolUser -UserPrincipalName $_.Stu_Email -password $_.Initial_PW -FirstName $_.First_Name -LastName $_.Last_Name -DisplayName $_.Display_Name -UsageLocation $_.Country -LicenseAssignment "ExampleCompany:STANDARDWOFFPACK_IW_STUDENT" -LicenseOptions $LOK5 -erroraction silentlycontinue
    }
   
} | export-csv $o365export -NoTypeInformation


# Giving time to process users, use at least 60 seconds for each new user. We need to input a number because the new-msoluser command doesn't wait until a user is creaed to try to move to the next set-mailbox step.
$seconds = read-host "Type in the number of seconds to process students (only input a number)."
start-sleep $seconds
write-verbose "Done processing, moving to next step."

# Set student email accounts up for auditing = note that you need speciic priviledges as an Admin to be able to audit accounts in O365.
function setAccounts {
    $newstu | ForEach-Object {
        if ([int]$_.grade -gt 5) {
        Set-Mailbox -identity $_.Stu_Email -AuditEnabled $true -AuditLogAgeLimit 180 -AuditAdmin Update, MoveToDeletedItems, SoftDelete, HardDelete, SendAs, SendOnBehalf, Create, UpdateFolderPermission -AuditDelegate Update, SoftDelete, HardDelete, SendAs, Create, UpdateFolderPermissions, MoveToDeletedItems, SendOnBehalf -AuditOwner UpdateFolderPermission, MailboxLogin, Create, SoftDelete, HardDelete, Update, MoveToDeletedItems | out-null
        }
    }
}

# The below function custom attributes, set parent forwarding, and add to student Address Book Policy. Note that this is a custom policy and yours will be different if you have one at all. 
function processAccounts {
    $newstu | foreach-object {
        if ([int]$_.grade -lt 6) {
            #We just need to set up a password and ensure the user's account isn't blocked if they are in grades K - 5
            set-MsolUserPassword -userprincipalname $_.Stu_Email -newpassword $_.Initial_PW
            set-msoluser -userprincipalname $_.Stu_Email -blockcredential $false
        }

        if ([int]$_.grade -gt 5) {
            #For users in grades 6+ they need to be set to the "Student" Address Book Policy (this is custom and you need to create your own ABP for your org), we set up email forwarding to parent account, then set new password and make sure their accounts are not blocked if the user is not a new account.
        set-mailbox -identity $_.Stu_Email -addressbookpolicy "Student.Abp" -roleassignmentpolicy "Student Role Assignment Policy" -CustomAttribute1 $_.Grade -CustomAttribute2 $_.Par_Email -delivertomailboxandforward $true -forwardingsmtpaddress $_.Par_Email
        set-MsolUserPassword -userprincipalname $_.Stu_Email -newpassword $_.Initial_PW
        set-msoluser -userprincipalname $_.Stu_Email -blockcredential $false
        }
    }   
}

setAccounts
write-verbose "Done creating new accounts."
processAccounts
write-verbose "Done processing accounts."


# notification that O365 mailbox settings are complete #
write-verbose "Parent email forwarding and student Address Book Policy set."

$pcuser = gci "c:\users\" -recurse -directory | where { $_.name -like 'user1' -or $_.name -like 'Administrator' }

if ($pcuser.fullname -eq "C:\users\srowl") {
    $script:creds = import-clixml 'C:\PathToFolder\auto.xml' 
}

elseif ($pcuser.fullname -eq "C:\users\seanr") {
    $script:creds = import-clixml 'C:\PathToFolder\auto2.xml' 
}

elseif ($pcuser.fullname -eq "C:\users\Administrator") {
    try {
        $script:creds = import-clixml 'C:\PathToFolder\auto3.xml'
    } 
    
    catch { 
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        $msg = $e.Message  
    }

    finally {
        $script:creds = import-clixml 'C:\PathToFolder\auto4.xml' 
    }
} 
 

 

# Function to send email notifications to appropriate parties letting them know students have been created in O365
function notify-admin {
    # credentials to send email & smtp email settings #
    $to = "someemail@test.org", "someemail2@test.org"
    $from = "fromemail@test.org"
    $subject = "New O365 Users Added"
    $creds
    $smtpserver = 'smtp.office365.com'
    $smtpport = '587'

    # This sample shows how to link o speciic folder and select the most recently edited file to attach to your email notification
    $SampleFileFolderVariable = (gci "C:\PathToFolder\Enrollments\import-files\" | sort LastWriteTime | select fullname -last 1)

    # This gets the speciic filename from teh above sample folder so you can actually attach the file to the email.
    $SampleFileVariable = $SampleFileFolderVariable.fullname


    # Creating the list of students for the body of the email.
    foreach ($s in $newstu) {

        $script:studentRow = @"
        <tr>
            <td>
                $($s.Last_Name), $($s.First_Name)
            </td>
            <td>
                 $($s.Stu_Email)
            </td>
            <td>
                $($s.Par_First) $($s.Par_Last)
            </td>
            <td>
                $($s.Par_Email)
            </td>
        </tr>
"@

        # Looping through and adding each table row to the body of the email
        $studentData += $studentRow
    }

    # Body of notification, adjust the CSS to your own needs.
    $body = @"
        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
        <style type="text/css">
            @import url("https://use.typekit.net/xyf7vbf.css");
            body {min-width: 100%!important;}
            .content {width: 100%; max-width: 600px;}
            @media only screen and (max-width: 480px){
            .bodyContent{font-size:16px !important;margin-left: 10px;}
        </style>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>New O365 User Added</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
        </head>
        <body>
        <table
    cellpadding="0"
    cellspacing="0"
    height="100%"
    width="100%"
    align="center"
    border="0"
    id="bodytable"
  >
    <tr
      style="background-color:#780000; color:#ffffff; font-size:16pt; font-weight:bold;"
    >
      <td colspan="6">New Student Information</td>
    </tr>
    <tr
      style="background-color:#898585; color:#ffffff; font-size:12pt; font-weight:bold; font-family: sans-serif"
    >
      <td>
        Student Name
      </td>
      <td>
        Email
      </td>
      <td>
        Parent Name
      </td>
      <td>
        Parent Email
      </td>
    $($studentData)
  </table>
        </body>
        </head>
        </html>
"@
    # Uncomment the attachment line and add your own attachment variables if you need.
    $mailparam = @{
        to         = $to
        #cc = $cc
        from       = $from
        Subject    = $subject
        SmtpServer = $smtpserver
        port       = $smtpport
        credential = $creds
        body       = $body
        bodyashtml = $True
        # attachment = $SampleFileVariable
    }
        
    
    # Email error notification in case the email notiication is not successful. Note that that users will still be created even if this email fails.
    send-mailmessage @mailparam -UseSsl -erroraction continue -ErrorVariable whoops
    if ($whoops) { 
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        $msg = $e.Message 

        Write-Host -ForegroundColor Red "caught exception: $e at $line"

    }
    else { write-verbose "Notification email sent to admins." }


    
}


notify-admin @mailparam -UseSsl 

$resetVerboseForegroundColor = $host.privatedata.VerboseForegroundColor

write-verbose "Setting email transfer rules."
#invoke-expression -command $transportScript

write-verbose "Student O365 account creation complete!"