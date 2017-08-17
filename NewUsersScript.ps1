############################
## D Morris and M Whittle ##
## New User Script v1.4   ##
##  12/8/2017 Branch 4    ##
############################


##v1.1 Exchange online mailbox based on profiles location only.  
##     Email address in the email report.
##v1.5 Added Function to allow selection of CSV file
##v1.6 Added Exit sendmail message function and new log file process.

   
######################
## Create functions

##Send Mail function

Function Send-reportmsg($output)
{
        	#get username of logged on person and get email address to send report to
        	$username = [Environment]::UserName
        
        	if ($output_email = get-aduser $username -properties emailaddress|select -ExpandProperty emailaddress)
            		{
                		Write-Host "Logged on user is mail enabled so sending report to $($output_email)"
                		send-MailMessage -From 'itservicedesk@nottinghamcity.gov.uk' -SmtpServer nccexw2k122 -port "25" -To $output_email -Subject "New NCC Accounts"  -BodyAsHtml $output -Attachments "\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log"
            		}
        	else
            	{
               		#logged on user isn't mail enabled so 
                	$output_mail_user = Get-ADGroupMember -identity "IT email list"| Select-Object Name, samaccountname | Out-GridView -OutputMode Single -Title "Select User who will receive email report"
	                $output_mail_address = $output_mail_user|select -ExpandProperty samaccountname |get-aduser -properties mail
        	        Write-Host "Logged on user is not mail enabled so sending report to $($output_mail_address.mail)"
                	send-MailMessage -From 'itservicedesk@nottinghamcity.gov.uk' -SmtpServer nccexw2k122 -port "25" -To $output_mail_address.mail -Subject "New NCC Accounts" -BodyAsHtml $output -Attachments "\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log"
            	}
} #end function

## Log writing function
Function LogWrite
    {
        Param ([string]$logstring)
        $Logstring = "$(Get-Date -Format G): " + $Logstring
        Add-content $Logfilepath -value $logstring -PassThru
    }

## Log writing function
Function ErrorLogWrite
    {
        Param ([string]$errorlogstring)
        $errorlogstring = "$(Get-Date -Format G): " + $errorlogstring
        Add-content $errorLogfilepath -value $errorlogstring -PassThru |write-host -ForegroundColor Yellow -BackgroundColor Black
        $error_count++
    }


##Import CSV Function
Function Get-ImportCSV($initialDirectory, $DialogTitle)
{
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $DialogTitle
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “All files (*.*)| *.*”
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} #end function

#Sleep with progress bar function
function Start-Sleeping($seconds,$sleepmsg) 
    {
        $doneDT = (Get-Date).AddSeconds($seconds)
        while($doneDT -gt (Get-Date)) 
            {
                $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
                $percent = ($seconds - $secondsLeft) / $seconds * 100
                Write-Progress -Activity "Working" -Status $sleepmsg -SecondsRemaining $secondsLeft -PercentComplete $percent
                [System.Threading.Thread]::Sleep(500)
            }
        Write-Progress -Activity "Working" -Status $sleepmsg -SecondsRemaining 0 -Completed
    }

##Date variable
$date=get-date -format "dd-MM-yyyy hh:mm tt"

##Timestamp filter for log entries.
function global:timestamp {"$(Get-Date -Format G): $_"}


#check for log file and create if necessary
$Logfilepath="\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log"
$ErrorLogfilepath="\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\usererror.log"



#Map M drive if not mapped
 if (!(Get-PSDrive -Name "M"))
    {
        New-PSDrive -PSProvider FileSystem -Root "\\nottinghamcity.gov.uk\shd_res" -Name "M" -Persist
    }

## create logfile
if (!(Test-Path $LogFilePath))
    {
        New-Item -path "\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log" -name new_users.log -type "file" -Force
        LogWrite "Created new log file"
    }
else
    {
    #Make a copy of the existing log file
    try{Get-Item -Path "\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log" -ErrorAction Stop | where-object {$_.length -gt 10mb} | ForEach-Object {Rename-Item $_ ($_ -replace "log","old")}}catch{write-host -ForegroundColor red -BackgroundColor Blue "Log file not renamed"}
    }

## create error logfile
if (!(Test-Path $ErrorLogFilePath))
    {
        New-Item -path "\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log" -name usererror.log -type "file" -Force
        LogWrite "Created new error log file"
    }

#clear the output variables
$output = ""
$usererror = ""
$neus=0

#create the output mail header

$Output = $output +"<head> 
        <STYLE TYPE='text/css'>
        <!--
        BODY
           {
  
           font-family:sans-serif;
           }
        --> 
        </STYLE>


        </head>
        <body bgcolor='grey' align='center' width='100%' >
        <table bgcolor='white' align='center' width='99%' height='100%' border='0'>

        <tr><td colspan='6'> Please find the details for new accounts created. Users will be asked to change their password first time they log on </br></br>
        </td></tr>
        <tr height='20'><td bgcolor='#BBD723' align='center' colspan='6'> <b> The following users were created</b></td></tr> 
        <tr><td><b>Given Name</b></td><td><b>Surname</b></td><td><b> Job Title </b><td><b>  Username </b><td><b>Email Address</b></td><td><b>Password </b></td></tr> "



#Now Import the .csv file
$new_users=get-ImportCsv "\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\NewUserScript" "Select User Creation CSV file"|import-csv


#null input data variable
$newuser=$null
##Main loop through user list to create new users

#connect to Exchange onprem
if (get-pssession)
    {get-pssession|Remove-PSSession;$s = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://nccexw2k123/powershell
Import-PSSession -Session $s -AllowClobber}
else
    {$s = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://nccexw2k123/powershell
Import-PSSession -Session $s -AllowClobber}
 
foreach($newuser in $new_users)
    {
        #Form proposed name    
        $GivenName = $newuser.GivenName.Trim()
        $SurName = $newuser.SurName.Trim()
        $name = $GivenName + " " +$SurName

        #region newuser params
        #deal with required and optional parameters of new-aduser cmdlet
        #GivenName,Surname,Title,samaccountname,Department,Company,Manager,Office,OfficePhone,Description,StreetAddress,City,State,PostalCode,profile
        if (($newuser.GivenName) -and ($newuser.Surname) -and ($newuser.Title) -and ($newuser.department)-and ($newuser.samaccountname) -and ($newuser.company) -and ($newuser.office) -and ($newuser.Description) -and ($newuser.streetaddress) -and ($newuser.city))
                {
                    logwrite "--------------------------------------------"
                    logwrite "CSV fields complete for creation of new user $($Name)" -Force
                    
                }
        else
                {
                    errorlogwrite "Creation of new user $($Name) failed because a required field within the CSV is empty"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "csv field empty" -Force
                    Send-reportmsg $output
                    exit
                }

    #endregion
    
    #Create variables as some cmdlets do not allow script block substitution
    $manager = $newuser.Manager
    $newusers = $null
    $new_user=$null
    $country ="GB"
    #Try to get managers AD object
    try
        {
            $manager = get-aduser $newuser.Manager
        }
    catch
        {
            ErrorLogWrite "Creation of new user $($Name) failed because managers AD object canot be resolved"
            continue
        }
        
    #Get Profilers AD details and store them in a variable
    try
        {
            
            $profile = get-aduser $newuser.profile -properties msexchrecipienttypedetails,distinguishedname,samaccountname,mail,company
        }
    catch
        {
            ErrorLogWrite "Creation of new user $($Name) failed because profilers AD object canot be resolved"
            continue
        }
    

    $expiry = $null
    $changepw = $null
    #If ChangePW or expiry aren't set in CSV then do nothing otherwise 
    if ($newuser.changepw -ne "") {$changepw = $newuser.changepw}
    ##Calculate the last day that the AD user account will be allowed to logon before expiring
    if ($newuser.expiry -ne "") {$expiry = $newuser.expiry|get-date|% {$_.adddays(1)}}
    
        
    #Regex Pattern to check name contains only alpha characters
    $regex = "^([a-zA-Z-]+)$"
    #Check name only contains alpha characters
    if (($GivenName -notmatch $regex) -or ($SurName -notmatch $regex))
        {
            ErrorLogWrite "Names must only contain alpha characters a to z and A to Z. Unable to create $($GivenName) $($SurName)"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "Invalid User Name" -Force
            Send-reportmsg $output
            exit
        }
    
    $logonname = $GivenName + "." +$Surname
    
    #If profilers account is not mail enabled, decide how to formulate email domain
    if ($profile.mail.split("@")[1] -eq $null) 
        {
            #Profilers account is not mail enabled so assuming no mailbox required
            #get UPN of profilers AD object and use that to form the UPN suffix
            $profilers_maildomain = $profile.UserPrincipalName.split("@")[1]
            $object = $logonname+"@"+$profilers_maildomain
            
        }
    else
        {
            #split profilers mail string to identify correct mail domain
            $object = $logonname+"@"+$profile.mail.split("@")[1]
        }
    
    $samaccountname = $newuser.samaccountname
    #Get OU name of profilers AD user account
    $profile_oudn=($profile.DistinguishedName) -replace "^CN=[^,]+,"
    
    #Now figure out if mail enabled onprem, online or not mail enabled
    switch ($profile.msexchrecipienttypedetails)
        {
            2147483648 {Add-Member -InputObject $newuser -Name "online" -MemberType NoteProperty -Value "$true" -Force}
            1 {Add-Member -InputObject $newuser -Name "onprem" -MemberType NoteProperty -Value "$true" -Force}
            default {Add-Member -InputObject $newuser -Name "notmailenabled" -MemberType NoteProperty -Value "$true" -Force}
        }
       
   
        ##start writing to log file 
        
        logwrite "Processing started on  $($date): "

        logwrite "--------------------------------------------"
 
 
        #---------Array to generate Password from-------#
        $result = $null
        $set1 = "abcdfghjklmnpqrstvwxyz".ToCharArray()
        $set2 = "aeiou".ToCharArray()
        $result += $set1 | Get-Random
        $result += $set2 | Get-Random
        $result += $set1 | Get-Random
        $result += $set2 | Get-Random
        $Password=$result.Substring(0,1).ToUpper(1)+$result.Substring(1,3)+(Get-Random -Minimum 1000 -Maximum 9999) 
      
               
        ##UPN creation algorithm
        $Count = $null
                    
        #loop to find unique $object
        LogWrite "Searching AD for $($object)...please wait"
           while (Get-ADObject -Properties mail, proxyAddresses , userprincipalname -Filter "mail -like '*$object*' -or proxyAddresses -like '*$object*' -or userprincipalname -like '*$object*'")
            {
                #The uniqueness check failed, add +1 to count and try again
                $Count++
                if ($Count -eq $object.Length)
                    {
                        logwrite "Username uniqueness failure. Please create AD account manually"
                        Send-reportmsg $output
                        break
                    }
                if ($Count -gt 0)
                    {
                        $prefix = $object -split "@"
                        $username=$logonname+$Count
                        $object= $username+"@"+$prefix[1]
                    }
                
            }
          
        #Check proposed name is unique in profilers OU
        if (get-aduser -Filter {name -eq $name} -searchscope Base -SearchBase $profile_oudn)
            {
                ErrorLogWrite "$($name) already exists in $($profile_oudn) so exiting script"
                Send-reportmsg $output
                exit
            }
        else
            {
                LogWrite "$($name) is unique in $($profile_oudn) so continuing with script"
               
            }
                  
            #get email domain and create U drive mapping based on profilers email domain suffix
            switch ($profile.mail.split("@")[1])
                {
                        "nottinghamcity.gov.uk" {$HomeDrive='U:';$UserRoot='\\nccfsw2k122\users4\';$HomeDirectory=$UserRoot+$samaccountname}
                        "collegest.org" {$HomeDrive='U:';$UserRoot='\\nccfsw2k122\users4\';$HomeDirectory=$UserRoot+$samaccountname}
                        "nottinghamcityhomes.org.uk" {$HomeDrive='U:';$UserRoot='\\NCCFSW2K122\NCHUsers\';$HomeDirectory=$UserRoot+$samaccountname} 
                        "robinhoodenergy.co.uk" {$HomeDrive='U:';$UserRoot='\\nccrhw2k121\users\';$HomeDirectory=$UserRoot+$samaccountname; $RHEBat='rhe.bat';$userprincipalname= $logonname+"@nottinghamcity.gov.uk";}
                        default {$HomeDrive='U:';$UserRoot='\\nccfsw2k122\users4\';$HomeDirectory=$UserRoot+$samaccountname}
                }
            
            #Got enough info to create a user
            logwrite "Creating AD user object:"
            logwrite $Name

            #region new-aduser
            #try..catch new user creation
            try
                {
                    $obj_newuser = New-ADUser -Name $Name -SamAccountName $SamAccountName -UserPrincipalName $object -GivenName $GivenName `
                    -Surname $Surname -DisplayName $Name -Description $newuser.Description -Title $newuser.Title -Department $newuser.Department -Company $newuser.Company -Office $newuser.Office`
                    -OfficePhone $newuser.OfficePhone -Manager $Manager -StreetAddress $newuser.StreetAddress -City $newuser.City -State $newuser.state -PostalCode $newuser.PostalCode`
                    -Country $Country -Path $profile_oudn `
                    -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $True -changepasswordatlogon $true -HomeDrive $HomeDrive -HomeDirectory $HomeDirectory -PassThru -erroraction Continue
                }
            catch
                {
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    Errorlogwrite "Unable to create new user $($name) because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "new-aduser fail" -Force
                    Send-reportmsg $output
                    exit
                }
            #endregion
            
            #region Create Home drive for new user
            try
                {
                    NEW-ITEM –path $HomeDirectory -type directory -force
                }
            catch
                {
             
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    Errorlogwrite "Unable to create home drive folder for $($newuser.Name) on $($HomeDirectory) because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "create home drive failure"
                    Send-reportmsg $output
                    exit
                }
            #endregion    
            
            #region obtain current ACL and add new user full control ACL entity to this    
                try
                    {
                        $Acl = (Get-Item $HomeDirectory).GetAccessControl('Access')
                        $Acl.setaccessruleprotection($True, $true)
                        $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($obj_newuser.SID, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
                        $Acl.AddAccessRule($Ar)
                    }
                catch
                    {
                        $ErrorMessage = $_.Exception.Message
                        $FailedItem = $_.Exception.ItemName
                        logwrite "Error modifying ACL for new user folder because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "ACL modification failure"
                        Send-reportmsg $output
                        exit
                    }
                #endregion   
                   
                #region set ACL on home drive
                try
                    {
                        Set-Acl -path $Homedirectory -aclobject $Acl
                        LogWrite "Permissions set for $name on $HomeDirectory `n"
                    }
                catch
                    {
                        $ErrorMessage = $_.Exception.Message
                        $FailedItem = $_.Exception.ItemName
                        ErrorLogWrite "Unable to set premissions for homedrive of user named $($newuser.Name) because $ErrorMessage"; logwrite "Unable to set homedrive on $($newuser.Name) homedrive because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "Home drive permissions failure"
                        Send-reportmsg $output
                        exit
                    }
                #endregion
                
            #Sets attribute for manual pin for followme Printing
           
                    try
                        {
                        $Rand = Get-Random -Minimum 1000 -Maximum 9999
                        Set-ADUser $samaccountname -replace @{'extensionattribute9' = $Rand}
                        LogWrite "Successfully added Manual Konica PIN to $($name)"
                        }
                    catch
                        {
                        $ErrorMessage = $_.Exception.Message
                        $FailedItem = $_.Exception.ItemName
                        ErrorLogWrite "Unable to set manual pin for $($name) because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "Manual PIN set fail" -Force
                        Send-reportmsg $output
                        exit
                        }
            #Assign RHE user logon script
            if ($profile.company -like "*robin*")
                {
                    Set-aduser $SamAccountName -ScriptPath $RHEBat;
                    set-aduser $samaccountname -UserPrincipalName $userprincipalname 
                }  
            #Output password to log file
         
                logwrite "Password for user $($Name) is $Password`n"
            
            if ($expiry)
                {
                    LogWrite "Account $($Name) is a temporary account so setting expiry date to $($Expiry)"
                    set-aduser $samaccountname -AccountExpirationDate $expiry
                }
            #If account is flagged as password must not be changed at first logon then deactivate requirement
            if (($changepw -eq "no") -or ($changepw -eq "n"))
                {
                    LogWrite "Account $($Name) is marked as password must not be changed at first logon so deactivating requirement"
                    set-aduser $samaccountname -ChangePasswordAtLogon $false
                }
        #create array of user profiler group memberships
        $GrpmbmProfile = Get-ADPrincipalGroupMembership $profile
                        
        #Check that AD group members of profiler account do not have data owners in the notes field
        $GrpNoDataOwner = $GrpmbmProfile |Where-Object {(Get-ADGroup -Filter {name -eq $_.Name} -Properties info|select info) -notmatch "owner"}

        #Create list of Security groups that has a data owner so these can be added to the email report for 1st line to contact
        $GrpDataOwnerChecks = $GrpmbmProfile |Where-Object {(Get-ADGroup -Filter {name -eq $_.Name} -Properties info|select info) -match "owner"}


        #create array of sensitive groups referenced by the Service Desk new user document
        $sensitivegroups = "all_localadmin","heads of service","directors","g.xenapp.CAG_Users","Business Support Colleagues","resources Leadership Team",`
        "it email list","business support colleagues","g.xenapp.CAG_Exception","it eitlt","it itlt","BME Staff","LGBT Group","Project Std","Project Pro",`
        "Crystal Reports","GCSX Users","PSN Users","Freja_access_sms_defunct","domain users"
        
        #Filter groups to remove groups starting with the name "mailbox."
        $GrpNotMailbox = $GrpNoDataOwner |Where-Object {$_.name -Notlike "mailbox.*"}
        $GrpNotGDotXen = $GrpNotMailbox |Where-Object {$_.name -Notlike "g.Xen*"}
 
        #loop through each group membership and test for restricted group status, add member if unrestricted 
        foreach ($item_innerloop in $GrpNotGDotXen)
            {
                if ($sensitivegroups -contains $item_innerloop.name)
                    {
                        LogWrite "The user cannot be added to this restricted group: $($item_innerloop.name)"
                    }   
                else
                    {
                        LogWrite "The user can be added to this unrestricted group: $($item_innerloop.name)"
                        try
                            {
                                Add-ADGroupMember $item_innerloop.name $SamAccountName
                            }
                        catch
                            {
                                $ErrorMessage = $_.Exception.Message
                                $FailedItem = $_.Exception.ItemName
                                ErrorLogwrite "Unable to add user named $($name) to AD security group $($item_innerloop.name) because $ErrorMessage"
                            }
                    }
               
             }
      
        #User is online so enable remote mailbox in Exchange online      
        if ($newuser.online)
            {
                LogWrite "Creating online mailbox for $name"
                Start-Sleeping -Seconds 15 "Waiting for AD to replicate the user account"
                enable-remotemailbox -identity $samaccountname -RemoteRoutingAddress $samaccountname@nottinghamcc.mail.onmicrosoft.com -PrimarySmtpAddress $object
                Start-Sleeping -Seconds 15 "Waiting for remote mailbox to replicate"
                $rep_mail=get-remotemailbox $samaccountname|select -ExpandProperty primarysmtpaddress
                Set-RemoteMailbox -Identity $samaccountname -EmailAddressPolicyEnabled $true
                #Add SIP Address for Skype For Business client SSO
                set-aduser $samaccountname -Add @{proxyAddresses = ("SIP:$object")}
            }
        elseif ($newuser.onprem)
            {
                #Onprem mail enable and other mailbox processing functions
                #Add SIP Address for Skype For Business client SSO
                LogWrite "Adding SIP address for $($name)"
                set-aduser $samaccountname -Add @{proxyAddresses = ("SIP:$object")}
                LogWrite "Creating onprem mailbox for $name"
                start-Sleeping -Seconds 20 "Waiting for replication"
                #enable mailbox based on new-aduser
                enable-mailbox -identity $samaccountname -Alias $samaccountname -DomainController nccadw2k122
                start-Sleeping -Seconds 20 "Waiting for replication"
                
                #Make sure onprem Exchange assigned SMTP as per expected behaviour
                $smtptest = get-mailbox $samaccountname |select -expandproperty primarysmtpaddress
                if ($smtptest -ne $object)
                    {
                        try
                        {
                            #make primary email address and upn, mail and SIP matches (except RHE for now)
                            set-mailbox $samaccountname -PrimarySmtpAddress $object -ErrorAction stop
                        }
                        catch
                        {		
                            $ErrorMessage = $_.Exception.Message
                            ErrorLogWrite "Error aligning Primary SMTP address with UPN status for $($name) because $ErrorMessage"
                            Add-Member -InputObject $newuser -Name "error" -MemberType NoteProperty -Value "UPN\SMTP mismatch"
                            Send-reportmsg $output
                            exit
				        }
                                        
                    }
                
                    #region Get email address for output
        
                    try
                        {
                            #wait for RUS to create the mailbox
                            while (!(get-mailbox -Identity $samaccountname -erroraction silentlycontinue)){Write-Host "Waiting for mailbox to exist"; Start-Sleep -Seconds 1}
                            $rep_mail=get-mailbox $samaccountname|select -ExpandProperty windowsemailaddress
                            #check mb has been created on non-GCSX database
                            #Form MBX server from random number
                            $mbxsrv="MBX-DB"+(get-random -Minimum 1 -Maximum 12)
                            $check_mbx = Get-Mailbox $samaccountname|select -ExpandProperty database
                            if ($check_mbx.startswith("G")) {LogWrite "Mailbox has been created on GCSX server so moving to $($mbxsrv)";try {New-MoveRequest $samaccountname -TargetDatabase $mbxsrv} catch {ErrorLogWrite "Error Moving mailbox $($name) because $error[0]"}

                                                            }
                                                    }
                    catch
                        {
                            $ErrorMessage = $_.Exception.Message
                            $FailedItem = $_.Exception.ItemName
                            ErrorLogWrite "Failed to get email address for $($newuser.Name) because $ErrorMessage"
                            Send-reportmsg $output
                            exit
                        }
            }
            else
                {
                    #No mailbox created
                    LogWrite "User $($object) is not mailenabled"
                    $rep_mail = "Not mail enabled"
                }
           
                        
        #endregion
    
                            
                
        #<tr><td><b>Given Name</b></td><td><b>Surname</b></td><td><b> Job Title </b><td><b>  Username </b><td><b>Password </b></td></tr> "
        $output = $output +  "<tr><td><b>$($newuser.Givenname)</b></td><td><b>$($newuser.Surname)</b></td><td><b>$($newuser.Title)</b></td><td><b>$samaccountname</b></td><td><b>$rep_mail</b></td><td><b>$Password</b></td></tr>"
        $neus++
        #endregion
    

        #Include Data owners in final report
        if ($GrpDataOwnerChecks)
            {
                #Setup Dataowner field headers
                $output = $output +  "<tr><td><b>Name of Group</b></td><td colspan='5'><b>Data Owner Field</b></td></tr>"
                foreach ($grpdataownercheck in $grpdataownerchecks)
                    {
                    #Get the contents of the notes field for each AD Group with a data owner
                    $rep_grp_info=get-adgroup -identity $grpdataownercheck.name -Properties info|select -ExpandProperty info
                    $output = $output +  "<tr><td>$($grpdataownercheck.name)</td><td colspan='5'>$rep_grp_info</td></tr>"
                    LogWrite "Unable to add $($name) to $($grpdataownercheck.name) because it has a data owner"
                    }
            }

        }
    
    
#If any errors then create error log and attach to email
if ($newuser.error)
    {
        #record any creation errors and outout to error log file
        $usererror = $usererror + "Failed to complete user creation for $name because " +$newuser.error + " `n"
        $usererror |Out-File "\\nottinghamcity.gov.uk\shd_res\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\usererror.log"
    }

#send email if any users were created
If($neus -gt 0) 
    {
     try
        {Send-reportmsg $output}
     catch
        {ErrorLogWrite "Failed to send report email for $($neus) users because $Error[0]"}
    }
    
$output=$null
