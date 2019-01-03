if (Get-Module -ListAvailable -Name MSOnline) {
} else {
    Write-Host "Microsoft Online Powershell Module is Missing, Please install before re-opening powershell"
    Exit
}

Add-Type -AssemblyName System.Windows.Forms

function refreshForm {
    $Form.Close();
    $Form.Dispose();
    Makeform;
}

function MakeForm {
    $Global:Form = New-Object Windows.Forms.Form;
    $Form.Text = "Office 365 - Licensing & Permissions - Reporting GUI";
    $Global:WorkingPath = Get-Location | Select-Object Path
    $Form.Width = 800;
    $Form.Height = 600;
    $Form.BackColor = "#dbdbdb"
    $Font = New-Object System.Drawing.Font("Calibri",13,[System.Drawing.FontStyle]::Regular);
    $Form.Font = $Font;
    $Form.Topmost = $True;
    if ($PermissionsView -eq $True){
        $DelegateLabel = New-Object System.Windows.Forms.Label;
        $DelegateLabel.Location = New-Object System.Drawing.Point(430,55);
        $DelegateLabel.AutoSize = $True;
        $DelegateLabel.Text = 'Access to Mailboxes:';
        $DelegateLabel.BackColor = "Transparent";
        $Form.Controls.Add($DelegateLabel);

        $DelegateLabel = New-Object System.Windows.Forms.Label;
        $DelegateLabel.Location = New-Object System.Drawing.Point(430,328);
        $DelegateLabel.AutoSize = $True;
        $DelegateLabel.Text = 'Access to Calendars:';
        $DelegateLabel.BackColor = "Transparent";
        $Form.Controls.Add($DelegateLabel);

        $DelegateLabel = New-Object System.Windows.Forms.Label;
        $DelegateLabel.Location = New-Object System.Drawing.Point(430,191);
        $DelegateLabel.AutoSize = $True;
        $DelegateLabel.Text = 'Member of Groups:';
        $DelegateLabel.BackColor = "Transparent";
        $Form.Controls.Add($DelegateLabel);

        $MailboxLabel = New-Object System.Windows.Forms.Label;
        $MailboxLabel.Location = New-Object System.Drawing.Point(30,50);
        $MailboxLabel.AutoSize = $True;
        $MailboxLabel.Text = 'Select a user to query their permissions';
        $MailboxLabel.BackColor = "Transparent";
        $Form.Controls.Add($MailboxLabel);
        $Font = New-Object System.Drawing.Font("Calibri",16,[System.Drawing.FontStyle]::Bold)
        $MailboxLabel.Font = $Font

        $MailboxList = New-Object System.Windows.Forms.ListBox;
        $MailboxList.Location = New-Object System.Drawing.Point(15,85);
        $MailboxList.Size = New-Object System.Drawing.Size(400,400);

        $ExportPermissionsButton = New-Object System.Windows.Forms.Button;
        $ExportPermissionsButton.Location = New-Object System.Drawing.Point(15,520);
        $ExportPermissionsButton.Size = New-Object System.Drawing.Size(100,30);
        $ExportPermissionsButton.Text = 'Export-All';
        $ExportPermissionsButton.Add_click({ExportAllPermissions});
        $Form.Controls.Add($ExportPermissionsButton);
        
        $MailboxCheckbox = New-Object System.Windows.Forms.Checkbox 
        $MailboxCheckbox.Location = New-Object System.Drawing.Point(125,522) 
        $MailboxCheckbox.AutoSize = $True;
        $Global:MailboxCheck = 0
        $MailboxCheckbox.Add_click({
            if ($MailboxCheck -eq 0 ) {
                $Global:ExportMailboxes = $True;
                $MailboxCheck = $MailboxCheck + 1
            } else {
                $Global:ExportMailboxes = $False;
                $MailboxCheck = $MailboxCheck - 1
            }
        })
        $MailboxCheckbox.Text = "Mailboxes"
        $MailboxCheckbox.BackColor = "Transparent";
        $Form.Controls.Add($MailboxCheckbox)

        $GroupsCheckbox = New-Object System.Windows.Forms.Checkbox 
        $GroupsCheckbox.Location = New-Object System.Drawing.Point(227,522) 
        $GroupsCheckbox.AutoSize = $True;
        $Global:GroupsCheck = 0
        $GroupsCheckbox.Add_click({
            if ($GroupsCheck -eq 0 ) {
                $Global:ExportGroups = $True;
                $GroupsCheck = $GroupsCheck + 1
            } else {
                $Global:ExportGroups = $False;
                $GroupsCheck = $GroupsCheck - 1
            }
        })
        $GroupsCheckbox.Text = "Groups"
        $GroupsCheckbox.BackColor = "Transparent";
        $Form.Controls.Add($GroupsCheckbox)

        $CalendarCheckbox = New-Object System.Windows.Forms.Checkbox 
        $CalendarCheckbox.Location = New-Object System.Drawing.Point(309,522) 
        $CalendarCheckbox.AutoSize = $True;
        $Global:CalendarsCheck = 0
        $CalendarCheckbox.Add_click({
            if ($CalendarsCheck -eq 0 ) {
                $Global:ExportCalendars = $True;
                $CalendarsCheck = $CalendarsCheck + 1
            } else {
                $Global:ExportCalendars = $False;
                $CalendarsCheck = $CalendarsCheck - 1
            }
        })
        $CalendarCheckbox.Text = "Calendars"
        $CalendarCheckbox.BackColor = "Transparent";
        $Form.Controls.Add($CalendarCheckbox)

        $Global:MailboxUsers = Get-Mailbox
        $UserUsers = $MailboxUsers | Where-Object RecipientTypeDetails -eq UserMailbox | Select-Object UserPrincipalName 
        foreach ($UserUser in $UserUsers | Sort-Object -Property UserPrincipalName){
            $MailboxList.Items.Add($UserUser.UserprincipalName)
        }
        $MailboxList.Add_click({
            $Global:QueryUser = $MailboxList.Text
            if ($MailboxPermissions){
                $Form.Controls.Remove($MailboxPermissions)
            }
            if ($CalendarPermissions){
                $Form.Controls.Remove($CalendarPermissions)
            }
            if ($GroupPermissions){
                $Form.Controls.Remove($GroupPermissions)
            }
            QueryButton;
        })
        $Form.Controls.Add($MailboxList);
    }
    if ($LicenseView -eq $true) {
        LicensingStatus;
        $LicensedUsers = Get-MsolUser | Where-Object isLicensed -eq $true | Select-Object UserPrincipalName, Licenses;
        $UsersLabel = New-Object System.Windows.Forms.Label;
        $UsersLabel.Location = New-Object System.Drawing.Point(20,50);
        $UsersLabel.AutoSize = $True;
        $UsersLabel.Text = "Select a user to query their licenses";
        $UsersLabel.BackColor = "Transparent";
        $Form.Controls.Add($UsersLabel)
        $Font = New-Object System.Drawing.Font("Calibri",16,[System.Drawing.FontStyle]::Bold)
        $UsersLabel.Font = $Font
        $UserList = New-Object System.Windows.Forms.ListBox;
        $UserList.Location = New-Object System.Drawing.Point(20,90);
        $UserList.Size = New-Object System.Drawing.Size(350,300);
        foreach ($LicensedUser in $LicensedUsers | Sort-Object -Property UserPrincipalName){
            $UserList.Items.Add($LicensedUser.UserprincipalName)
        }
        $UserList.Add_click({
            QueryLicense $UserList.Text;
        })
        $Form.Controls.Add($UserList);

        $AssignedLabel = New-Object System.Windows.Forms.Label;
        $AssignedLabel.Location = New-Object System.Drawing.Point(400,60);
        $AssignedLabel.AutoSize = $True;
        $AssignedLabel.Text = "Assigned Licenses:";
        $AssignedLabel.BackColor = "Transparent";
        $Form.Controls.Add($AssignedLabel)

        $ExportCsvButton = New-Object System.Windows.Forms.Button;
        $ExportCsvButton.Location = New-Object System.Drawing.Point(415,390);
        $ExportCsvButton.Size = New-Object System.Drawing.Size(115,30);
        $ExportCsvButton.Text = 'Export-All';
        $ExportCsvButton.Add_click({ExportAllLicensedUsers});
        $Form.Controls.Add($ExportCsvButton);
    } 
    if ($LoggedIn -eq $true) {
        Navigation
    } else {
        $SignInFont = New-Object System.Drawing.Font("Calibri",15,[System.Drawing.FontStyle]::Bold);
        $SigninLabel = New-Object System.Windows.Forms.Label;
        $SigninLabel.Location = New-Object System.Drawing.Point(255,75);
        $SigninLabel.Size = New-Object System.Drawing.Size(300,40);
        $SigninLabel.Text = 'Office365 Admin Logon';
        $SigninLabel.BackColor = "Transparent";
        $SigninLabel.Font = $SignInFont
        $Form.Controls.Add($SigninLabel);
        $UserBox = New-Object System.Windows.Forms.TextBox;
        $UserBox.Location = New-Object System.Drawing.Point(250,120);
        $UserBox.Size = New-Object System.Drawing.Size(260,20);
        $UserBox.Text = '<username>'
        $Form.Controls.Add($UserBox);
        $Form.Add_Shown({$UserBox.Select()})
        $PassBox = New-Object System.Windows.Forms.MaskedTextBox;
        $PassBox.Location = New-Object System.Drawing.Point(250,150);
        $PassBox.Size = New-Object System.Drawing.Size(260,20);
        $PassBox.Text = '<password>'
        $PassBox.PasswordChar = '*'
        $Form.Controls.Add($PassBox);
        $SigninButton = New-Object System.Windows.Forms.Button;
        $SigninButton.Location = New-Object System.Drawing.Point(355,190);
        $SigninButton.Size = New-Object System.Drawing.Size(75,25);
        $SigninButton.Text = 'Sign-In';
        $SigninButton.Add_click({
            $Global:Username = $UserBox.Text;
            $Global:Password = $PassBox.Text;
            PleaseWait
            Login
        });
        $Form.AcceptButton = $SigninButton;
        $Form.Controls.Add($SigninButton);
        $CancelButton = New-Object System.Windows.Forms.Button;
        $CancelButton.Location = New-Object System.Drawing.Point(435,190);
        $CancelButton.Size = New-Object System.Drawing.Size(75,25);
        $CancelButton.Text = 'Cancel';
        $Form.CancelButton = $CancelButton;
        $Form.Controls.Add($CancelButton);
    }
    $Form.ShowDialog()
}

function PermissionsIndex(){
    $Global:PermissionsView = $True
    $Global:LicenseView = $False;
    $Global:LoggedIn = $True;
    refreshForm;
}
function LicenseIndex(){
    $Global:LicenseView = $True;
    $Global:LoggedIn = $True;
    $Global:PermissionsView = $False;
    refreshForm;
}

function Login (){
    $SecureStringPwd = $Password | ConvertTo-SecureString -AsPlainText -Force ;
    $Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecureStringPwd;
    Connect-MsolService -Credential $Creds;
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
        -Credential $Creds -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    $Global:LoggedIn = $True;
    $Global:LicenseView = $True
    refreshForm;
}

function PleaseWait {
    $Global:PleaseWait = New-Object System.Windows.Forms.Label;
    $PleaseWait.Location = New-Object System.Drawing.Point(260,180);
    $PleaseWait.Size = New-Object System.Drawing.Size(280,20);
    $PleaseWait.BackColor = "#14a919"
    $PleaseWait.Text = '                      Please Wait...';
    $Form.Controls.Add($PleaseWait);
    $PleaseWait.BringToFront();
}

function Navigation (){
    $NavFont = New-Object System.Drawing.Font("Calibri",11,[System.Drawing.FontStyle]::Regular);

    $LicenseButton = New-Object System.Windows.Forms.Button;
    $LicenseButton.Location = New-Object System.Drawing.Point(10,5);
    $LicenseButton.Size = New-Object System.Drawing.Size(90,30);
    $LicenseButton.Text = 'Licensing';
    $LicenseButton.Add_click({LicenseIndex});
    $Form.Controls.Add($LicenseButton);
    $Form.AcceptButton = $LicenseButton
    $LicenseButton.Font = $NavFont;

    $PermissionsButton = New-Object System.Windows.Forms.Button;
    $PermissionsButton.Location = New-Object System.Drawing.Point(105,5);
    $PermissionsButton.Size = New-Object System.Drawing.Size(90,30);
    $PermissionsButton.Text = 'Permissions';
    $PermissionsButton.Add_click({PermissionsIndex});
    $Form.Controls.Add($PermissionsButton);
    $PermissionsButton.Font = $NavFont;

    $ExitButton = New-Object System.Windows.Forms.Button;
    $ExitButton.Location = New-Object System.Drawing.Point(700,5);
    $ExitButton.Size = New-Object System.Drawing.Size(80,30);
    $ExitButton.Text = 'Exit';
    $ExitButton.Add_click({
        $Global:LoggedIn = $False;
    });
    $Form.Controls.Add($ExitButton);
    $Form.CancelButton = $ExitButton
    $ExitButton.Font = $NavFont

    TenantName
}
function TenantName (){
    $Global:CompanyName = Get-MsolPartnerInFormation;
    $CompanyName = $CompanyName.PartnerCompanyName;
    $NameFont = New-Object System.Drawing.Font("Calibri",16,[System.Drawing.FontStyle]::Bold);
    $TenantName = New-Object System.Windows.Forms.Label;
    $TenantName.Location = New-Object System.Drawing.Point(200,10);
    $TenantName.AutoSize = $True;
    $TenantName.Text = $CompanyName;
    $TenantName.BackColor = "Transparent";
    $TenantName.Font = $NameFont
    $Form.Controls.Add($TenantName)
}

function QueryPermissions {
    $Global:MailboxPermissions = New-Object System.Windows.Forms.ListBox;
    $MailboxPermissions.Location = New-Object System.Drawing.Point(430,82);
    $MailboxPermissions.Size = New-Object System.Drawing.Size(345,110);

    $UserPermissions = $MailboxUsers | Get-MailboxPermission -User $QueryUser | Select-Object Identity
    foreach ($UserPermission in $UserPermissions) {
        $GrantedMailbox = $UserPermission.Identity
        $GrantedMailbox = $GrantedMailbox.split("/")
        if ($GrantedMailbox) {
            $MailboxPermissions.Items.Add($GrantedMailbox[-1])
        }
    }
    $Form.Controls.Add($MailboxPermissions)
}

function QueryGroups {
    $Global:GroupPermissions = New-Object System.Windows.Forms.ListBox;
    $GroupPermissions.Location = New-Object System.Drawing.Point(430,218);
    $GroupPermissions.Size = New-Object System.Drawing.Size(345,110);

    $QueryUser = $QueryUser.split("@")
    $Groups = Get-DistributionGroup
    foreach ($Group in $Groups) {
        $MemberOf = Get-DistributionGroupMember -Identity $Group.PrimarySmtpAddress | Where-Object Name -match $QueryUser[0]
        if ($MemberOf) {
            $GroupPermissions.Items.Add($Group.PrimarySmtpAddress)
        }
    }
    $Form.Controls.Add($GroupPermissions)
}

function QueryCalendars {
    $Global:CalendarPermissions = New-Object System.Windows.Forms.ListBox;
    $CalendarPermissions.Location = New-Object System.Drawing.Point(430,355);
    $CalendarPermissions.Size = New-Object System.Drawing.Size(345,110);

    $QueryUser = $QueryUser.split("@")
    foreach ($MailboxUser in $MailboxUsers){
        $DelegatedCalendar = Get-MailboxFolderPermission -Identity "$($MailboxUser.UserPrincipalName):\Calendar"
        $DelegatedCalendarUser = $DelegatedCalendar | Where-Object User -match $QueryUser[0] | Select-Object Accessrights
        if ($DelegatedCalendarUser) {
            $Row = $MailboxUser.UserPrincipalName + "  " + $DelegatedCalendarUser.Accessrights
            $CalendarPermissions.Items.Add($Row)
        }
    }
    $Form.Controls.Add($CalendarPermissions)
}

function QueryButton {
    $QueryFont = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)

    $QueryButton = New-Object System.Windows.Forms.Button;
    $QueryButton.Location = New-Object System.Drawing.Point(15,470);
    $QueryButton.Size = New-Object System.Drawing.Size(130,30);
    $QueryButton.Text = 'Query-Mailboxes';
    $QueryButton.Add_click({
        PleaseWait
        QueryPermissions
        $Form.Controls.Remove($PleaseWait)
    });
    $Form.Controls.Add($QueryButton);
    $QueryButton.Font = $QueryFont;

    $QueryButton = New-Object System.Windows.Forms.Button;
    $QueryButton.Location = New-Object System.Drawing.Point(285,470);
    $QueryButton.Size = New-Object System.Drawing.Size(130,30);
    $QueryButton.Text = 'Query-Calendars';
    $QueryButton.Add_click({
        Pleasewait
        QueryCalendars
        $Form.Controls.Remove($PleaseWait)
    });
    $Form.Controls.Add($QueryButton);
    $QueryButton.Font = $QueryFont;

    $QueryButton = New-Object System.Windows.Forms.Button;
    $QueryButton.Location = New-Object System.Drawing.Point(150,470);
    $QueryButton.Size = New-Object System.Drawing.Size(130,30);
    $QueryButton.Text = 'Query-Groups';
    $QueryButton.Add_click({
        Pleasewait
        QueryGroups
        $Form.Controls.Remove($PleaseWait)
    });
    $Form.Controls.Add($QueryButton);
    $QueryButton.Font = $QueryFont;

    $SlowFont = New-Object System.Drawing.Font("Calibri",10,[System.Drawing.FontStyle]::Regular)
    $SlowLabel = New-Object System.Windows.Forms.Label;
    $SlowLabel.Location = New-Object System.Drawing.Point(330,502);
    $SlowLabel.AutoSize = $True;
    $SlowLabel.Text = '(slow)';
    $SlowLabel.BackColor = "Transparent";
    $SlowLabel.Font = $SlowFont;
    $Form.Controls.Add($SlowLabel);

}
function QueryLicense ([string]$arg1) {
    if ($ProductList){
        $Form.Controls.Remove($ProductList)
    }
    $Global:ProductList = New-Object System.Windows.Forms.ListBox;
    $ProductList.Location = New-Object System.Drawing.Point(400,90);
    $ProductList.Size = New-Object System.Drawing.Size(250,100);
    $ProductList.Add_click({
        $Global:SelectSku = $ProductList.Text
    })
    $LicensedUser = Get-MsolUser | Where-Object UserPrincipalName -eq $arg1
    $Products = $LicensedUser.Licenses
    foreach ($Product in $Products) {
        $ProductName = $Product.AccountSkuId.split(':')
        $ProductList.Items.Add($ProductName[1])
    }
    $Form.Controls.Add($ProductList)
}

function LicensingStatus () {
    $Skus = Get-MsolAccountSku | Select-Object AccountSkuId, ActiveUnits, ConsumedUnits
    $SkuCoords = 390
    foreach ($Sku in $Skus) {
        $TypeLabel = New-Object System.Windows.Forms.Label;
        $TypeLabel.Location = New-Object System.Drawing.Point(20,390);
        $TypeLabel.AutoSize = $True;
        $TypeLabel.Text = "License Types";
        $TypeLabel.BackColor = "Transparent";
        $Form.Controls.Add($TypeLabel)
        $Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Bold)
        $TypeLabel.Font = $Font

        $Type = $Sku.AccountSkuId.split(":")
        $SkuCoords = $SkuCoords + 23
        $SkuType = New-Object System.Windows.Forms.Label;
        $SkuType.Location = New-Object System.Drawing.Point(20,$SkuCoords);
        $SkuType.AutoSize = $True;
        $SkuType.Text = $Type[1];
        $SkuType.BackColor = "Transparent";
        $Form.Controls.Add($SkuType)

        $ActiveLabel = New-Object System.Windows.Forms.Label;
        $ActiveLabel.Location = New-Object System.Drawing.Point(260,390);
        $ActiveLabel.AutoSize = $True;
        $ActiveLabel.Text = "In-Use";
        $ActiveLabel.BackColor = "Transparent";
        $Form.Controls.Add($ActiveLabel)
        $Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Bold)
        $ActiveLabel.Font = $Font

        $SkuAssigned = New-Object System.Windows.Forms.Label;
        $SkuAssigned.Location = New-Object System.Drawing.Point(260,$SkuCoords);
        $SkuAssigned.AutoSize = $True;
        $SkuAssigned.Text = $Sku.ConsumedUnits;
        $SkuAssigned.BackColor = "Transparent";
        $Form.Controls.Add($SkuAssigned)

        $TotalLabel = New-Object System.Windows.Forms.Label;
        $TotalLabel.Location = New-Object System.Drawing.Point(330,390);
        $TotalLabel.AutoSize = $True;
        $TotalLabel.Text = "Total";
        $TotalLabel.BackColor = "Transparent";
        $Form.Controls.Add($TotalLabel)
        $Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Bold)
        $TotalLabel.Font = $Font

        $SkuTotal = New-Object System.Windows.Forms.Label;
        $SkuTotal.Location = New-Object System.Drawing.Point(330,$SkuCoords);
        $SkuTotal.AutoSize = $True;
        $SkuTotal.Text = $Sku.ActiveUnits;
        $SkuTotal.BackColor = "Transparent";
        $Form.Controls.Add($SkuTotal)
    }
}

function ExportAllLicensedUsers {
    PleaseWait
    $Filename = $WorkingPath.Path + "\" + $CompanyName.PartnerCompanyName + "-ALL-ASSIGNED-LICENSES.csv"
    $LicensedAccounts = Get-MSOLUser -All | Where-Object {$_.isLicensed -eq $True}
    foreach ($LicensedAccount in $LicensedAccounts) {
        $DisplayName = $LicensedAccount.DisplayName
        $UPN = $LicensedAccount.UserPrincipalName
        $Licenses = ""
        $Skus = $LicensedAccount.Licenses
        foreach ($Sku in $Skus) {
            $License = $Sku.AccountSkuId.split(':')
            if ($Licenses -ne "" ) {
                $Licenses = $Licenses + "; " + $License[1]
            } else {
                $Licenses = $Licenses + $License[1]
            }
        }
        [PSCustomObject] @{
            DisplayName = $DisplayName
            Upn = $UPN
            Licenses = $Licenses
            } | Export-CSV -append $Filename -NoTypeInformation -Encoding UTF8
        }

    $Note = "CSV File Created in: " + $WorkingPath.Path + "\"
    $CreatedNote = New-Object System.Windows.Forms.Label;
    $CreatedNote.Location = New-Object System.Drawing.Point(415,425);
    $CreatedNote.AutoSize = $True;
    $CreatedNote.Text = $Note;
    $CreatedNote.BackColor = "Transparent";
    $CreatedNote.ForeColor='Green'
    $Form.Controls.Add($CreatedNote);
    $CreatedFont = New-Object System.Drawing.Font("Calibri",11,[System.Drawing.FontStyle]::Bold);
    $CreatedNote.Font = $CreatedFont;
    $Form.Controls.Remove($PleaseWait)
}

function ExportAllPermissions {
    PleaseWait;
    if ($ExportMailboxes -eq $True) {
        $MailboxFile = $WorkingPath.Path + "\" + $CompanyName.PartnerCompanyName + "-DELEGATE-MAILBOX-PERMISSIONS.csv"
        $DelegateMailboxes = Get-Mailbox | Get-MailboxPermission `
        | Where-Object {($_.IsInherited -eq $False) -and ($_.User -notmatch "SELF") -and ($_.User -notmatch "Discovery")} `
        | Select-Object Identity, User, AccessRights
        foreach ($DelegateMailbox in $DelegateMailboxes) {
            $Identity = $DelegateMailbox.Identity
            $Identity = $Identity.split("/")
            $User = $DelegateMailbox.User
            $AccessRights = $DelegateMailbox.Accessrights
            [PSCustomObject] @{
                Mailbox = $Identity[-1]
                DelegatedUser = $User
                AccessRights = $AccessRights
                } | Export-CSV -append $MailboxFile -NoTypeInformation -Encoding UTF8
            }

    }
    if ($ExportGroups -eq $True) {
        $GroupsFile = $WorkingPath.Path + "\" + $CompanyName.PartnerCompanyName + "-GROUP-MEMBERSHIPS.csv"
        $Groups = Get-DistributionGroup
        $Groups | ForEach-Object {
        $Group = $_.PrimarySmtpAddress
        $Members = ''
        Get-DistributionGroupMember $Group | ForEach-Object {
                if($Members) {
                      $Members=$Members + ";" + $_.Name
                   } else {
                      $Members=$_.Name
                   }
          }
        [PSCustomObject] @{
            GroupName = $Group
            Members = $Members
             }
        } | Export-CSV $GroupsFile -NoTypeInformation -Encoding UTF8

    }
    if ($ExportCalendars -eq $True) { 
        $CalendarFile = $WorkingPath.Path + "\" + $CompanyName.PartnerCompanyName + "-ALL-DELEGATED-CALENDARS.csv"
        foreach ($MailboxUser in $MailboxUsers){
            $DelegatedCalendars = Get-MailboxFolderPermission -Identity "$($MailboxUser.UserPrincipalName):\Calendar" `
            | Where-Object {($_.User -notmatch "SELF") -and ($_.User -notmatch "Discovery") -and ($_.Accessrights -notmatch "None")}
            foreach ($DelegatedCalendar in $DelegatedCalendars){
                if ($MailboxUser.UserPrincipalName -notmatch "Discovery") {
                    $Identity = $MailboxUser.UserPrincipalName
                    $User = $DelegatedCalendar.User
                    $AccessRights = $DelegatedCalendar.Accessrights
                    [PSCustomObject] @{
                        Identity = $Identity
                        User = $User
                        AccessRights = $AccessRights
                        } | Export-CSV -append $CalendarFile -NoTypeInformation -Encoding UTF8
                    }
                }
        }       
    }
    $Form.Controls.Remove($PleaseWait)
    $Note = "CSV File Created in: " + $WorkingPath.Path + "\"
    $CreatedNote = New-Object System.Windows.Forms.Label;
    $CreatedNote.Location = New-Object System.Drawing.Point(410,530);
    $CreatedNote.AutoSize = $True;
    $CreatedNote.Text = $Note;
    $CreatedNote.BackColor = "Transparent";
    $CreatedNote.ForeColor='Green'
    $Form.Controls.Add($CreatedNote);
    $CreatedFont = New-Object System.Drawing.Font("Calibri",11,[System.Drawing.FontStyle]::Bold);
    $CreatedNote.Font = $CreatedFont;
}

Makeform