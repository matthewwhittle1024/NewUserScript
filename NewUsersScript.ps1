 ############################
## D Morris and M Whittle ##
## New User Script v1.3   ##
##  2/6/2017 Branch 3     ##
############################

　
##v1.1 Exchange online mailbox based on profiles location only.  
##     Email address in the email report.
#region

##Date variable
$date=get-date -format "dd-MM-yyyy hh:mm tt"

#check for log file and create if necessary
$Logfilepath="M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log"
$ErrorLogfilepath="M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\usererror.log"

　
#Map M drive if not mapped
 if (!(Get-PSDrive -Name "M"))
    {
        New-PSDrive -PSProvider FileSystem -Root "\\nottinghamcity.gov.uk\shd_res" -Name "M" -Persist
    }

## create logfile
if (!(Test-Path $LogFilePath))
    {
        New-Item -path "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log" -name new_users.log -type "file" -Force
        Write-Host "Created new log file"
    }
else
    {
    #Make a copy of the existing log file
    copy-Item "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log"
    "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users_$date.log"
    }
## create error logfile
if (!(Test-Path $ErrorLogFilePath))
    {
        New-Item -path "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log" -name usererror.log -type "file" -Force
        Write-Host "Created new error log file"
    }

## Create functions
## Log writing function
Function LogWrite
    {
        Param ([string]$logstring)
        Add-content $Logfilepath -value $logstring -PassThru
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

        <tr><td colspan='6'> Please find the details for new accounts created. Users will be asked to change their password first time they log on </br></br>If you have any problems please see the script authors
        </td></tr>
        <tr height='20'><td bgcolor='#BBD723' align='center' colspan='6'> <b> The following users were created</b></td></tr> 
        <tr><td><b>Given Name</b></td><td><b>Surname</b></td><td><b> Job Title </b><td><b>  Username </b><td><b>Email Address</b></td><td><b>Password </b></td></tr> "

　
　
#function to test for email address existence in nottinghamcity.gov.uk as a proxy address in AD
function Get-EmailAddressTest
    {
        [CmdletBinding()]
        param
       (
              [Parameter(Mandatory = $True,
                              ValueFromPipeline = $True,
                              ValueFromPipelineByPropertyName = $True,
                              HelpMessage = 'What e-mail address would you like to find?')]
              [string[]]$EmailAddress
       )
      
       process
       {            
              foreach ($address in $EmailAddress)
              {
                     Get-ADObject -Properties mail, proxyAddresses -Filter "mail -like '*$address*' -or proxyAddresses -like '*$address*'"
              }
       }
    }

#Import .csv
$new_users=Import-Csv "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\NewUserScript\usercreationfile.csv"

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
                    logwrite "CSV fields complete for creation of new user $($Name)" -Force
                    
                }
        else
                {
                    logwrite "Creation of new user $($Name) failed because a required field within the CSV is empty"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "csv field empty" -Force
                    exit
                }

    #endregion
    
    #Create variables as some cmdlets do not allow script block substitution
    $manager = $newuser.Manager 
    $country ="GB"
    $manager = get-aduser $newuser.Manager
    $profile = get-aduser $newuser.profile -properties *
    $expiry = $null
    $changepw = $null
    #If ChangePW or expiry aren't set in CSV then do nothing otherwise 
    if ($newuser.changepw -ne "") {$changepw = $newuser.changepw}
    if ($newuser.expiry -ne "") {$expiry = $newuser.expiry}

    
    #Regex Pattern to check name contains only alpha characters
    $regex = "^([a-zA-Z-]+)$"
    #Check name only contains alpha characters
    if (($GivenName -notmatch $regex) -or ($SurName -notmatch $regex))
        {
            logwrite "Names must only contain alpha characters a to z and A to Z. Unable to create $($GivenName) $($SurName)"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "Invalid User Name" -Force
            exit
        }
    
    $logonname = $GivenName + "." +$Surname

    $profilers_em = $profile.mail.ToString()
    $profilers_maildomain = $profilers_em -split "@"
    $object = $logonname+"@"+$profilers_maildomain[1]
    $samaccountname = $newuser.samaccountname
    $profile_oudn=($profile.DistinguishedName) -replace "^CN=[^,]+,"
      if (((Get-ADUser $profile -Properties *| select msexchrecipienttypedetails) -match "2147483648"))
        {
            try
                {
                   #Mark $newuser as online
                   Add-Member -InputObject $newuser -Name "online" -MemberType NoteProperty -Value "$true"
                }
            catch
                {
                    $ErrorMessage = $_.Exception.Message
                    LogWrite "Error setting online status for $($profile) because $ErrorMessage"
                }
        }
        else
            {
                try
                    {
                        #Mark $newuser as onprem
                        Add-Member -InputObject $newuser -Name "onprem" -MemberType NoteProperty -Value "$true"
                    }
                catch
                    {
                        $ErrorMessage = $_.Exception.Message
                        LogWrite "Error setting onprem status for $($profile) because $ErrorMessage"
                    }
            }        
   #endregion
   
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
        while (Get-ADObject -Properties mail, proxyAddresses , userprincipalname -Filter "mail -like '*$object*' -or proxyAddresses -like '*$object*' -or userprincipalname -like '*$object*'")
            {
                #The uniqueness check failed, add +1 to count and try again
                $Count++
                if ($Count -eq $object.Length)
                    {
                        logwrite "Username uniqueness failure. Please create AD account manually" ; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "logon name creation uniqueness fail"
                        break
                    }
                    if ($Count -gt 1)
                    {
                        $prefix = $object -split "@"
                        $username=$prefix[0]+$Count
                        $object= $username+"@"+$profilers_maildomain[1]
                    }
            }
        
                    
          
        #Check proposed name is unique in profilers OU
        if (get-aduser -Filter {name -eq $name} -searchscope Base -SearchBase $profile_oudn)
            {
                LogWrite "$($name) already exists in $($profile_oudn) so exiting script"
                exit
            }
        else
            {
                LogWrite "$($name) is unique in $($profile_oudn) so continuing with script"
               
            }
                  
            #get email domain and create U drive mapping based on profilers email domain suffix
            switch ($profilers_maildomain[1])
                {
                        "nottinghamcity.gov.uk" {$HomeDrive='U:';$UserRoot='\\nccfsw2k122\users4\';$HomeDirectory=$UserRoot+$samaccountname}
                        "collegest.org" {$HomeDrive='U:';$UserRoot='\\nccfsw2k122\users4\';$HomeDirectory=$UserRoot+$samaccountname}
                        "nottinghamcityhomes.org.uk" {$HomeDrive='U:';$UserRoot='\\NCCFSW2K122\NCHUsers\';$HomeDirectory=$UserRoot+$samaccountname} 
                        "robinhoodenergy.co.uk" {$HomeDrive='U:';$UserRoot='\\nccrhw2k121\users\';$HomeDirectory=$UserRoot+$samaccountname; $RHEBat='rhe.bat';$userprincipalname= $logonname+"@nottinghamcity.gov.uk";}
                        default {$HomeDrive='U:';$UserRoot='\\nccfsw2k122\users4\';$HomeDirectory=$UserRoot+$samaccountname}
                }
            
            #Got enough info to create a user
            logwrite "creating AD user object:"
            logwrite $Name

            #region new-aduser
            #try..catch new user creation
            try
                {
                    $obj_newuser = New-ADUser -Name $Name -SamAccountName $SamAccountName -UserPrincipalName $object -GivenName $GivenName `
                    -Surname $Surname -DisplayName $Name -Description $newuser.Description -Title $newuser.Title -Department $newuser.Department -Company $newuser.Company -Office $newuser.Office`
                    -OfficePhone $newuser.OfficePhone -Manager $Manager -StreetAddress $newuser.StreetAddress -City $newuser.City -PostalCode $newuser.PostalCode`
                    -Country $Country -Path $profile_oudn `
                    -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $True -changepasswordatlogon $true -HomeDrive $HomeDrive -HomeDirectory $HomeDirectory -PassThru -erroraction Continue
                }
            catch
                {
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    logwrite "Unable to create new user $($name) because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "new-aduser fail" -Force
                    break
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
                    logwrite "Unable to create home drive folder for $($newuser.Name) on $($HomeDirectory) because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "create home drive failure"
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
                        exit
                    }
                #endregion   
                   
                #region set ACL on home drive
                try
                    {
                        Set-Acl -path $Homedirectory -aclobject $Acl
                        logwrite "Permissions set for $name on $HomeDirectory `n"
                    }
                catch
                    {
                        $ErrorMessage = $_.Exception.Message
                        $FailedItem = $_.Exception.ItemName
                        logwrite "Unable to set premissions for homedrive of user named $($newuser.Name) because $ErrorMessage"; logwrite "Unable to set homedrive on $($newuser.Name) homedrive because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "Home drive permissions failure"
                    }
                #endregion

                }
         
            #endregion

            #Sets attribute for manual pin for followme Printing
            if (!($newuser.error))
                {
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
                        logwrite "Unable to set manual pin for $($name) because $ErrorMessage"; Add-Member -InputObject $newuser -Name "Error" -MemberType NoteProperty -Value "Manual PIN set fail" -Force
                        break
                        }
            #Assign RHE user logon script
            if ($profilers_maildomain[1] -eq "robinhoodenergy.co.uk")
                {
                    Set-aduser $SamAccountName -ScriptPath $RHEBat
                }  
            #Output password to log file
            if (!($newuser.Error))
            {
                logwrite "Password for user $($Name) is $Password`n"
            }
            if ($expiry)
                {
                    LogWrite "Account $($Name) is a temporary account so setting expiry date to $($Expiry)"
                    set-aduser $samaccountname -AccountExpirationDate $expiry
                }
            #If account is flagged as password must not be changed at first logon then deactivate requirement
            if ($changepw)
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
        "Crystal Reports"
        
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
                                logwrite "Unable to add user named $($newuser.Name) to AD security group $($item_innerloop.name) because $ErrorMessage"
                            }
                    }
               
             }
      
        #User is online so enable remote mailbox in Exchange online      
        if ($newuser.online)
            {
                LogWrite "Creating online mailbox for $name"
                Start-Sleep -Seconds 15

                enable-remotemailbox -identity $samaccountname -RemoteRoutingAddress $samaccountname@nottinghamcc.onmicorosoft.com -PrimarySmtpAddress $object
                #Add SIP Address for Skype For Business client SSO
                set-aduser $samaccountname -Add @{proxyAddresses = ("SIP:$object")}
            }
        else
            {
                LogWrite "Creating onprem mailbox for $name"
                start-Sleep -Seconds 15
                #enable mailbox based on new-aduser
                enable-mailbox -identity $samaccountname
                start-Sleep -Seconds 15
                #Add SIP Address for Skype For Business client SSO
                LogWrite "Adding SIP address for $($name)"
                set-aduser $samaccountname -Add @{proxyAddresses = ("SIP:$object")}
                #Make sure onprem Exchange assigned SMTP as per expected behaviour
                $smtptest = get-mailbox $samaccountname |select -expandproperty primarysmtpaddress
                if ($smtptest -ne $object)
                    {
                        try
                        {
                            #make primary email address and upn, mail and SIP matches (except RHE for now)
                            set-mailbox $samaccountname -PrimarySmtpAddress $object
                        }
                        catch
                        {		
                            $ErrorMessage = $_.Exception.Message
                            LogWrite "Error aligning Primary SMTP address with UPN status for $($name) because $ErrorMessage"
                            Add-Member -InputObject $newuser -Name "error" -MemberType NoteProperty -Value "UPN\SMTP mismatch"
                            break
				        }
                                        
                    }
                $usererror = $usererror + "Failed to complete user creation for $name because " +$newuser.error + " `n"
            }   
                        
    #endregion
    
        #region Get email address for output
        
        try
            {
                #wait for RUS to create the mailbox
                while (!(get-mailbox -Identity $samaccountname)){$true; Start-Sleep -Seconds 1}
                $rep_mail=get-mailbox $samaccountname|select -ExpandProperty windowsemailaddress
                #check mb has been created on non-GCSX database
                #Form MBX server from random number
                $mbxsrv="MBX-DB"+(get-random -Minimum 1 -Maximum 12)
                $check_mbx = Get-Mailbox $samaccountname|select -ExpandProperty database
                if ($check_mbx.startswith("G")) {LogWrite "Mailbox has been created on GCSX server so moving to $($mbxsrv)";New-MoveRequest $samaccountname -TargetDatabase $mbxsrv}

            }
        catch
            {
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    logwrite "Failed to get email address for $($newuser.Name) because $ErrorMessage"
            }
                            
                
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
                    LogWrite "Unable to add $($newuser.name) to $($grpdataownercheck.name) because it has a data owner"
                    }
            }

        }
    

#If any errors then create error log and attach to email
if ($usererror)
    {
        $usererror |Out-File "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\usererror.log"
    }

　
　
#send email if any users were created
If($neus -gt 0) 
    {
        #get username of logged on person and get email address to send report to
        $username = [Environment]::UserName
        
        if ($output_email = get-aduser $username -properties emailaddress|select -ExpandProperty emailaddress)
            {
                
                send-MailMessage -From 'itservicedesk@nottinghamcity.gov.uk' -SmtpServer outlook.nottinghamcity.gov.uk  -port "25" -To $output_email -Subject "New NCC Accounts"  -BodyAsHtml $output -Attachments "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log"
            }
        else
            {
                #logged on user isn't mail enabled so 
                $output_mail_user = Get-ADGroupMember -identity "IT email list"| Select-Object Name, samaccountname | Out-GridView -OutputMode Single -Title "Select User who will receive email report"
                $output_mail_address = $output_mail_user|select -ExpandProperty samaccountname |get-aduser -properties *
                Write-Host "Logged on user is not mail enabled so sending report to $($output_mail_address.mail)"
                send-MailMessage -From 'itservicedesk@nottinghamcity.gov.uk' -SmtpServer outlook.nottinghamcity.gov.uk  -port "25" -To $output_mail_address.mail -Subject "New NCC Accounts"  -BodyAsHtml $output -Attachments "M:\IT\IT Teams\I.T. Technical Services Support\Projects\Ofice 365 - Mailbox Migrations\Log\new_users.log"
            }
    }
    
$output=$null 
