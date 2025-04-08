function Connect-ToExchangeOnline {
    try {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        
        # Check if the EXO V2 module is installed
        if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Host "ExchangeOnlineManagement module is not installed. Installing..." -ForegroundColor Yellow
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
        }
        
        # Import the module
        Import-Module ExchangeOnlineManagement
        
        # Connect to Exchange Online
        Connect-ExchangeOnline
        
        Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error connecting to Exchange Online:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        return $false
    }
}

# Function to create Excel file with multiple sheets
function Create-ExcelFile {
    param(
        [string]$FilePath,
        [array]$FullAccessData,
        [array]$SendAsData,
        [array]$SendOnBehalfData
    )
    
    try {
        Write-Host "Creating Excel file..." -ForegroundColor Cyan
        
        # Check if ImportExcel module is installed
        if (!(Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "ImportExcel module is not installed. Installing..." -ForegroundColor Yellow
            Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser
        }
        
        # Import the module
        Import-Module ImportExcel
        
        # Create Excel file with multiple sheets
        $FullAccessData | Export-Excel -Path $FilePath -WorksheetName "FullAccess" -AutoSize -TableName "FullAccess"
        $SendAsData | Export-Excel -Path $FilePath -WorksheetName "SendAs" -AutoSize -TableName "SendAs"
        $SendOnBehalfData | Export-Excel -Path $FilePath -WorksheetName "SendOnBehalf" -AutoSize -TableName "SendOnBehalf"
        
        Write-Host "Excel file created successfully." -ForegroundColor Green
        Write-Host "File saved at: $FilePath" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Error creating Excel file:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
    }
}

# Function to get all user mailboxes
function Get-AllUserMailboxes {
    try {
        Write-Host "Retrieving all mailboxes (user and shared)..." -ForegroundColor Cyan
        $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox
        Write-Host "Found" $mailboxes.Count "mailboxes." -ForegroundColor Green
        return $mailboxes
    }
    catch {
        Write-Host "Error retrieving user mailboxes:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        return $null
    }
}

# Function to get permissions for all user mailboxes
function Get-AllMailboxPermissions {
    Write-Host "Getting permissions for ALL mailboxes (user and shared)..." -ForegroundColor Cyan
    
    # Get all user mailboxes
    $mailboxes = Get-AllUserMailboxes
    
    if ($null -eq $mailboxes) {
        Write-Host "No mailboxes found or error occurred." -ForegroundColor Red
        return
    }
    
    # Initialize arrays to store permission data
    $fullAccessPermissions = @()
    $sendAsPermissions = @()
    $sendOnBehalfPermissions = @()
    
    # Counter for progress display
    $counter = 0
    $total = $mailboxes.Count
    
    foreach ($mailbox in $mailboxes) {
        $counter++
        $percentComplete = ($counter / $total) * 100
        $statusMessage = "Processing mailbox $counter of $total"
        Write-Progress -Activity "Processing Mailbox Permissions" -Status $statusMessage -PercentComplete $percentComplete
        
        # Get FullAccess permissions
        try {
            $fullAccess = Get-MailboxPermission -Identity $mailbox.UserPrincipalName | 
                Where-Object {$_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY\*"}
            
            foreach ($permission in $fullAccess) {
                $permObject = New-Object PSObject
                $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
                $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
                $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $permission.User
                $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "FullAccess"
                $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value $permission.IsInherited
                $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value $permission.Deny
                $fullAccessPermissions += $permObject
            }
            
            Write-Host "  Processed FullAccess for" $mailbox.UserPrincipalName -ForegroundColor DarkGray
        }
        catch {
            Write-Host "  Error getting FullAccess permissions for" $mailbox.UserPrincipalName -ForegroundColor Red
            Write-Host $_ -ForegroundColor Red
        }
        
        # Get SendAs permissions
        try {
            $sendAs = Get-RecipientPermission -Identity $mailbox.UserPrincipalName | 
                Where-Object {$_.AccessRights -eq "SendAs" -and $_.IsInherited -eq $false -and $_.Trustee -notlike "NT AUTHORITY\*"}
            
            foreach ($permission in $sendAs) {
                $permObject = New-Object PSObject
                $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
                $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
                $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $permission.Trustee
                $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "SendAs"
                $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value $permission.IsInherited
                $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value $permission.Deny
                $sendAsPermissions += $permObject
            }
            
            Write-Host "  Processed SendAs for" $mailbox.UserPrincipalName -ForegroundColor DarkGray
        }
        catch {
            Write-Host "  Error getting SendAs permissions for" $mailbox.UserPrincipalName -ForegroundColor Red
            Write-Host $_ -ForegroundColor Red
        }
        
        # Get SendOnBehalf permissions
        try {
            $sendOnBehalf = $mailbox.GrantSendOnBehalfTo
            
            if ($sendOnBehalf -and $sendOnBehalf.Count -gt 0) {
                foreach ($user in $sendOnBehalf) {
                    $permObject = New-Object PSObject
                    $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
                    $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
                    $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $user
                    $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "SendOnBehalf"
                    $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value "N/A"
                    $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value "N/A"
                    $sendOnBehalfPermissions += $permObject
                }
            }
            
            Write-Host "  Processed SendOnBehalf for" $mailbox.UserPrincipalName -ForegroundColor DarkGray
        }
        catch {
            Write-Host "  Error getting SendOnBehalf permissions for" $mailbox.UserPrincipalName -ForegroundColor Red
            Write-Host $_ -ForegroundColor Red
        }
    }
    
    Write-Progress -Activity "Processing Mailbox Permissions" -Completed
    
    # Create timestamp for the filename
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $fileName = "M365_Mailbox_Permissions_All_" + $timestamp + ".xlsx"
    $filePath = "$env:USERPROFILE\Desktop\$fileName"
    
    # Create Excel file
    Create-ExcelFile -FilePath $filePath -FullAccessData $fullAccessPermissions -SendAsData $sendAsPermissions -SendOnBehalfData $sendOnBehalfPermissions
    
    # Display summary
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "Total mailboxes processed:" $total -ForegroundColor White
    Write-Host "FullAccess permissions found:" $fullAccessPermissions.Count -ForegroundColor White
    Write-Host "SendAs permissions found:" $sendAsPermissions.Count -ForegroundColor White
    Write-Host "SendOnBehalf permissions found:" $sendOnBehalfPermissions.Count -ForegroundColor White
}

# Function to get permissions for a specific mailbox
function Get-SpecificMailboxPermissions {
    param(
        [string]$UserPrincipalName
    )
    
    Write-Host "Getting permissions for mailbox:" $UserPrincipalName -ForegroundColor Cyan
    
    # Check if the mailbox exists
    try {
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
    }
    catch {
        Write-Host "Error: Mailbox not found." -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        return
    }
    
    # Initialize arrays to store permission data
    $fullAccessPermissions = @()
    $sendAsPermissions = @()
    $sendOnBehalfPermissions = @()
    
    # Get FullAccess permissions
    try {
        Write-Host "  Getting FullAccess permissions..." -ForegroundColor DarkGray
        $fullAccess = Get-MailboxPermission -Identity $UserPrincipalName | 
            Where-Object {$_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY\*"}
        
        foreach ($permission in $fullAccess) {
            $permObject = New-Object PSObject
            $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
            $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
            $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $permission.User
            $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "FullAccess"
            $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value $permission.IsInherited
            $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value $permission.Deny
            $fullAccessPermissions += $permObject
        }
        
        Write-Host "  Found" $fullAccessPermissions.Count "FullAccess permissions" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error getting FullAccess permissions:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
    }
    
    # Get SendAs permissions
    try {
        Write-Host "  Getting SendAs permissions..." -ForegroundColor DarkGray
        $sendAs = Get-RecipientPermission -Identity $UserPrincipalName | 
            Where-Object {$_.AccessRights -eq "SendAs" -and $_.IsInherited -eq $false -and $_.Trustee -notlike "NT AUTHORITY\*"}
        
        foreach ($permission in $sendAs) {
            $permObject = New-Object PSObject
            $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
            $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
            $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $permission.Trustee
            $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "SendAs"
            $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value $permission.IsInherited
            $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value $permission.Deny
            $sendAsPermissions += $permObject
        }
        
        Write-Host "  Found" $sendAsPermissions.Count "SendAs permissions" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error getting SendAs permissions:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
    }
    
    # Get SendOnBehalf permissions
    try {
        Write-Host "  Getting SendOnBehalf permissions..." -ForegroundColor DarkGray
        $sendOnBehalf = $mailbox.GrantSendOnBehalfTo
        
        if ($sendOnBehalf -and $sendOnBehalf.Count -gt 0) {
            foreach ($user in $sendOnBehalf) {
                $permObject = New-Object PSObject
                $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
                $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
                $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $user
                $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "SendOnBehalf"
                $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value "N/A"
                $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value "N/A"
                $sendOnBehalfPermissions += $permObject
            }
        }
        
        Write-Host "  Found" $sendOnBehalfPermissions.Count "SendOnBehalf permissions" -ForegroundColor Green
    }
    catch {
        Write-Host "  Error getting SendOnBehalf permissions:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
    }
    
    # Create timestamp for the filename
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $fileName = "M365_Mailbox_Permissions_" + $mailbox.DisplayName + "_" + $timestamp + ".xlsx"
    $filePath = "$env:USERPROFILE\Desktop\$fileName"
    
    # Create Excel file
    Create-ExcelFile -FilePath $filePath -FullAccessData $fullAccessPermissions -SendAsData $sendAsPermissions -SendOnBehalfData $sendOnBehalfPermissions
    
    # Display summary
    Write-Host ""
    Write-Host "Summary for mailbox" $mailbox.DisplayName ":" -ForegroundColor Cyan
    Write-Host "FullAccess permissions:" $fullAccessPermissions.Count -ForegroundColor White
    Write-Host "SendAs permissions:" $sendAsPermissions.Count -ForegroundColor White
    Write-Host "SendOnBehalf permissions:" $sendOnBehalfPermissions.Count -ForegroundColor White
}

# Function to get all mailboxes a specific user has access to
function Get-UserMailboxAccess {
    param(
        [string]$UserPrincipalName
    )
    
    Write-Host "Getting mailboxes that user $UserPrincipalName has access to..." -ForegroundColor Cyan
    
    # Check if the user exists
    try {
        # Try to get the user as a mailbox or recipient
        $user = Get-Recipient -Identity $UserPrincipalName -ErrorAction Stop
        Write-Host "  User $UserPrincipalName found as" $user.RecipientTypeDetails -ForegroundColor Green
    }
    catch {
        Write-Host "Error: User not found." -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        return
    }
    
    # Initialize arrays to store permission data
    $fullAccessPermissions = @()
    $sendAsPermissions = @()
    $sendOnBehalfPermissions = @()
    
    # Get all user mailboxes
    $mailboxes = Get-AllUserMailboxes
    
    if ($null -eq $mailboxes) {
        Write-Host "No mailboxes found or error occurred." -ForegroundColor Red
        return
    }
    
    Write-Host "Searching through" $mailboxes.Count "mailboxes for permissions..." -ForegroundColor Cyan
    
    # Counter for progress display
    $counter = 0
    $total = $mailboxes.Count
    
    # Check for FullAccess permissions across all mailboxes
    Write-Host "  Checking for FullAccess permissions..." -ForegroundColor DarkGray
    foreach ($mailbox in $mailboxes) {
        $counter++
        $percentComplete = ($counter / $total) * 100
        $statusMessage = "Processing mailbox $counter of $total"
        Write-Progress -Activity "Checking FullAccess Permissions" -Status $statusMessage -PercentComplete $percentComplete
        
        try {
            $fullAccess = Get-MailboxPermission -Identity $mailbox.UserPrincipalName | 
                Where-Object {$_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false}
            
            # Check if this user has access
            $userAccess = $fullAccess | Where-Object {
                $_.User -like "*$UserPrincipalName*" -or 
                $_.User -like "*$($user.Name)*" -or 
                $_.User -like "*$($user.DisplayName)*"
            }
            
            if ($userAccess) {
                foreach ($permission in $userAccess) {
                    $permObject = New-Object PSObject
                    $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
                    $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
                    $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $permission.User
                    $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "FullAccess"
                    $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value $permission.IsInherited
                    $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value $permission.Deny
                    $fullAccessPermissions += $permObject
                    
                    Write-Host "    Found FullAccess to:" $mailbox.DisplayName -ForegroundColor Green
                }
            }
        }
        catch {
            # Silently continue on errors
        }
    }
    
    Write-Progress -Activity "Checking FullAccess Permissions" -Completed
    
    # Reset counter for SendAs check
    $counter = 0
    
    # Check for SendAs permissions across all mailboxes
    Write-Host "  Checking for SendAs permissions..." -ForegroundColor DarkGray
    foreach ($mailbox in $mailboxes) {
        $counter++
        $percentComplete = ($counter / $total) * 100
        $statusMessage = "Processing mailbox $counter of $total"
        Write-Progress -Activity "Checking SendAs Permissions" -Status $statusMessage -PercentComplete $percentComplete
        
        try {
            $sendAs = Get-RecipientPermission -Identity $mailbox.UserPrincipalName | 
                Where-Object {$_.AccessRights -eq "SendAs" -and $_.IsInherited -eq $false}
            
            # Check if this user has access
            $userAccess = $sendAs | Where-Object {
                $_.Trustee -like "*$UserPrincipalName*" -or 
                $_.Trustee -like "*$($user.Name)*" -or 
                $_.Trustee -like "*$($user.DisplayName)*"
            }
            
            if ($userAccess) {
                foreach ($permission in $userAccess) {
                    $permObject = New-Object PSObject
                    $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
                    $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
                    $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $permission.Trustee
                    $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "SendAs"
                    $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value $permission.IsInherited
                    $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value $permission.Deny
                    $sendAsPermissions += $permObject
                    
                    Write-Host "    Found SendAs to:" $mailbox.DisplayName -ForegroundColor Green
                }
            }
        }
        catch {
            # Silently continue on errors
        }
    }
    
    Write-Progress -Activity "Checking SendAs Permissions" -Completed
    
    # Reset counter for SendOnBehalf check
    $counter = 0
    
    # Check for SendOnBehalf permissions across all mailboxes
    Write-Host "  Checking for SendOnBehalf permissions..." -ForegroundColor DarkGray
    foreach ($mailbox in $mailboxes) {
        $counter++
        $percentComplete = ($counter / $total) * 100
        $statusMessage = "Processing mailbox $counter of $total"
        Write-Progress -Activity "Checking SendOnBehalf Permissions" -Status $statusMessage -PercentComplete $percentComplete
        
        try {
            $sendOnBehalf = $mailbox.GrantSendOnBehalfTo
            
            # Check if this user is in the list
            if ($sendOnBehalf -and $sendOnBehalf.Count -gt 0) {
                foreach ($delegateUser in $sendOnBehalf) {
                    # Try to match by UPN, alias, or display name
                    if ($delegateUser -like "*$UserPrincipalName*" -or 
                        $delegateUser -like "*$($user.Name)*" -or 
                        $delegateUser -like "*$($user.DisplayName)*") {
                        
                        $permObject = New-Object PSObject
                        $permObject | Add-Member -MemberType NoteProperty -Name "Mailbox" -Value $mailbox.UserPrincipalName
                        $permObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mailbox.DisplayName
                        $permObject | Add-Member -MemberType NoteProperty -Name "UserWithAccess" -Value $delegateUser
                        $permObject | Add-Member -MemberType NoteProperty -Name "AccessType" -Value "SendOnBehalf"
                        $permObject | Add-Member -MemberType NoteProperty -Name "Inherited" -Value "N/A"
                        $permObject | Add-Member -MemberType NoteProperty -Name "DeniedRights" -Value "N/A"
                        $sendOnBehalfPermissions += $permObject
                        
                        Write-Host "    Found SendOnBehalf to:" $mailbox.DisplayName -ForegroundColor Green
                    }
                }
            }
        }
        catch {
            # Silently continue on errors
        }
    }
    
    Write-Progress -Activity "Checking SendOnBehalf Permissions" -Completed
    
    # Create timestamp for the filename
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $userPart = $UserPrincipalName.Split('@')[0]
    $fileName = "M365_User_Access_" + $userPart + "_" + $timestamp + ".xlsx"
    $filePath = "$env:USERPROFILE\Desktop\$fileName"
    
    # Create Excel file
    Create-ExcelFile -FilePath $filePath -FullAccessData $fullAccessPermissions -SendAsData $sendAsPermissions -SendOnBehalfData $sendOnBehalfPermissions
    
    # Display summary
    Write-Host ""
    Write-Host "Summary for user" $UserPrincipalName ":" -ForegroundColor Cyan
    Write-Host "Has FullAccess to" $fullAccessPermissions.Count "mailboxes" -ForegroundColor White
    Write-Host "Has SendAs to" $sendAsPermissions.Count "mailboxes" -ForegroundColor White
    Write-Host "Has SendOnBehalf to" $sendOnBehalfPermissions.Count "mailboxes" -ForegroundColor White
}

# Main menu function
function Show-Menu {
    Clear-Host
    Write-Host "===== M365 Mailbox Permissions Audit =====" -ForegroundColor Cyan
    Write-Host "1: Get permissions for ALL mailboxes (user and shared)" -ForegroundColor White
    Write-Host "2: Get permissions for a SPECIFIC mailbox" -ForegroundColor White
    Write-Host "3: Get mailboxes a specific user has access to" -ForegroundColor White
    Write-Host "4: Exit program" -ForegroundColor White
    Write-Host "=========================================" -ForegroundColor Cyan
    
    $choice = Read-Host "Please enter your choice (1-4)"
    
    switch ($choice) {
        "1" {
            Get-AllMailboxPermissions
            Write-Host "Press Enter to continue..."
            Read-Host | Out-Null
            Show-Menu
        }
        "2" {
            $upn = Read-Host "Enter the user principal name (UPN/email) of the mailbox"
            Get-SpecificMailboxPermissions -UserPrincipalName $upn
            Write-Host "Press Enter to continue..."
            Read-Host | Out-Null
            Show-Menu
        }
        "3" {
            $upn = Read-Host "Enter the user principal name (UPN/email) of the user"
            Get-UserMailboxAccess -UserPrincipalName $upn
            Write-Host "Press Enter to continue..."
            Read-Host | Out-Null
            Show-Menu
        }
        "4" {
            Write-Host "Exiting program..." -ForegroundColor Yellow
            exit
        }
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Write-Host "Press Enter to continue..."
            Read-Host | Out-Null
            Show-Menu
        }
    }
}

# Main script execution
# Check if connected to Exchange Online
$connected = Connect-ToExchangeOnline

if ($connected) {
    # Display the menu
    Show-Menu
}
else {
    Write-Host "Unable to connect to Exchange Online. Please make sure you have the necessary permissions and try again." -ForegroundColor Red
    Write-Host "Press Enter to exit..."
    Read-Host | Out-Null
}