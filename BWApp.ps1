
################################################
## START NEW WINDOW AS ADMINISTRATOR ELEVATED ##

if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   
  Exit
 }
}
$tempfile = "C:\Temp\path.bwapp"
$Path = Get-Content -path "$tempfile"
cd $Path

##########################
# CHECK CONSOLE SETTINGS #
##########################
$checkconsole = Get-Content -Path ".\bin\config.bwapp"
    if ($checkconsole -contains "console = enabled"){
    write-host -ForegroundColor Green 'Showing Console Window'
    $ShowConsole = 1
    }
    else {
    $ShowConsole = 0
    ## HIDE THE CONSOLE WINDOW BEFORE GUI LAUNCHES ##
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '    
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, $ShowConsole)
    }


Start-Sleep -Seconds 3

#################################################################

$version = '3'
$build = '2003.28'
$github = 'https://github.com/valllem'

#################################################################


# Load external assemblies
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$MS_Main = new-object System.Windows.Forms.MenuStrip



# BWApp Menu #
##############
$BWAppToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    ##############
    $LoginToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $PasswordToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $365PortalToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $AzurePortalToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $LogFileToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $LogFolderToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $QuitToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem

# Settings Menu #
#################
$SettingsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    #################
    $UpdateToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $ModulesToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    #################



# Help Menu #
#############
$HelpToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
#############
    $AboutToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem



# Reports Menu #
#############
$ReportsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
#############
    $OverviewToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $MailPermsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $AllMailPermsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $CheckForwardsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $DistMembersToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $MailLogs48ToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $MailLogs24ToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $LicensedUsersToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $ListEmailsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $CalendarPermsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem


# Tenant Menu #
#############
$TenantToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
#############
    $TenancyPrepWizToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $OrgCustomizationToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $AuditLogsToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $POPIMAPToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $IntuneImportToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $OutsideSenderToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $SpamPolicyToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
    $MalwarePolicyToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem









###########################################
                ###########
                # MS_Main #
                ###########
###########################################


###########
# MS_Main #                  <------- Add Top Menu Items Here
###########
$MS_Main.Items.AddRange(@(
$BWAppToolStripMenuItem,
$SettingsToolStripMenuItem,
$HelpToolStripMenuItem,
$ReportsToolStripMenuItem,
$TenantToolStripMenuItem
))
$MS_Main.Location = new-object System.Drawing.Point(0, 0)

$MS_Main.Name = "MS_Main"
$MS_Main.Size = new-object System.Drawing.Size(354, 24)
$MS_Main.TabIndex = 0
$MS_Main.Text = "menuStrip1"




##########################
# BWAppToolStripMenuItem #
##########################
$BWAppToolStripMenuItem.DropDownItems.AddRange(@(
$LoginToolStripMenuItem,
$PasswordToolStripMenuItem,
$365PortalToolStripMenuItem,
$AzurePortalToolStripMenuItem,
$LogFileToolStripMenuItem,
$LogFolderToolStripMenuItem,
$QuitToolStripMenuItem))
$BWAppToolStripMenuItem.Name = "BWAppToolStripMenuItem"
$BWAppToolStripMenuItem.Size = new-object System.Drawing.Size(35, 20)
$BWAppToolStripMenuItem.Text = "&BWApp"



        ##########################
        # LoginToolStripMenuItem #
        ##########################
        $LoginToolStripMenuItem.Name = "LoginToolStripMenuItem"
        $LoginToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $LoginToolStripMenuItem.Text = "&Login"
        function OnClick_LoginToolStripMenuItem($Sender,$e){
           #Login
            Clear-host
            write-host -ForegroundColor Yellow 'Logging In...'
            $UserCredential = Get-Credential
            Clear-host
            write-host -ForegroundColor Yellow 'Logging In...'
            write-host -ForegroundColor Yellow 'Connecting AzureAD'
            Connect-AzureAD -Credential $UserCredential
            
            
            try {
                write-host -ForegroundColor Yellow 'Connecting MSolService'
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
                Import-PSSession $Session -DisableNameChecking
                Start-Sleep -seconds 1
                Connect-MsolService -Credential $UserCredential
                Clear-host
                write-host -ForegroundColor Green 'Logged In'
                
            }
            catch {
                write-host -ForegroundColor Yellow 'MSolService Failed, this is normal for some Tenancies. Attempting workaround...'
                Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") `
                -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"} `
                |select -First 1)
                $EXOSession = New-ExoPSSession
                Import-PSSession $EXOSession

            }
        }
        $LoginToolStripMenuItem.Add_Click( { OnClick_LoginToolStripMenuItem $LoginToolStripMenuItem $EventArgs} )
        

        ##########################
        # PasswordToolStripMenuItem #
        ##########################
        $PasswordToolStripMenuItem.Name = "PasswordToolStripMenuItem"
        $PasswordToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $PasswordToolStripMenuItem.Text = "&Change Password"
        function OnClick_PasswordToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Change Password Script", "BWApp v$version")
        }
        $PasswordToolStripMenuItem.Add_Click( { OnClick_PasswordToolStripMenuItem $PasswordToolStripMenuItem $EventArgs} )
 
        
    
        ##########################
        # 365PortalToolStripMenuItem #
        ##########################
        $365PortalToolStripMenuItem.Name = "365PortalToolStripMenuItem"
        $365PortalToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $365PortalToolStripMenuItem.Text = "&365 Admin"
        function OnClick_365PortalToolStripMenuItem($Sender,$e){
            Start-Process "https://admin.microsoft.com"
        }
        $365PortalToolStripMenuItem.Add_Click( { OnClick_365PortalToolStripMenuItem $365PortalToolStripMenuItem $EventArgs} )
 
        ##########################
        # AzurePortalToolStripMenuItem #
        ##########################
        $AzurePortalToolStripMenuItem.Name = "AzurePortalToolStripMenuItem"
        $AzurePortalToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $AzurePortalToolStripMenuItem.Text = "&Azure Portal"
        function OnClick_AzurePortalToolStripMenuItem($Sender,$e){
            Start-Process "https://portal.azure.com"
        }
        $AzurePortalToolStripMenuItem.Add_Click( { OnClick_AzurePortalToolStripMenuItem $AzurePortalToolStripMenuItem $EventArgs} )

        ##########################
        # LogFileToolStripMenuItem #
        ##########################
        $LogFileToolStripMenuItem.Name = "LogFileToolStripMenuItem"
        $LogFileToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $LogFileToolStripMenuItem.Text = "Open Lof F&ile"
        function OnClick_LogFileToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Opens Log File", "BWApp v$version")
        }
        $LogFileToolStripMenuItem.Add_Click( { OnClick_LogFileToolStripMenuItem $LogFileToolStripMenuItem $EventArgs} )
 
        ##########################
        # LogFolderToolStripMenuItem #
        ##########################
        $LogFolderToolStripMenuItem.Name = "LogFolderToolStripMenuItem"
        $LogFolderToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $LogFolderToolStripMenuItem.Text = "Open Log &Folder"
        function OnClick_LogFolderToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Opens Log Folder", "BWApp v$version")
        }
        $LogFolderToolStripMenuItem.Add_Click( { OnClick_LogFolderToolStripMenuItem $LogFolderToolStripMenuItem $EventArgs} )

        #########################
        # QuitToolStripMenuItem #
        #########################
        $QuitToolStripMenuItem.Name = "QuitToolStripMenuItem"
        $QuitToolStripMenuItem.Size = new-object System.Drawing.Size(152, 24)
        $QuitToolStripMenuItem.Text = "&Sign Out"
        function OnClick_QuitToolStripMenuItem($Sender,$e){
            Get-PSSession | Remove-PSSession
            $MenuForm.close()
        }
        $QuitToolStripMenuItem.Add_Click( { OnClick_QuitToolStripMenuItem $QuitToolStripMenuItem $EventArgs} )
        

##########################
# SettingsToolStripMenuItem #
##########################
$SettingsToolStripMenuItem.DropDownItems.AddRange(@(
$UpdateToolStripMenuItem,
$ModulesToolStripMenuItem))
$SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
$SettingsToolStripMenuItem.Size = new-object System.Drawing.Size(51, 20)
$SettingsToolStripMenuItem.Text = "&Settings"

        #########################
        # UpdateToolStripMenuItem #
        #########################
        $UpdateToolStripMenuItem.Name = "UpdateToolStripMenuItem"
        $UpdateToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $UpdateToolStripMenuItem.Text = "&Update App"
        function OnClick_UpdateToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Update App", "BWApp v$version")
        }
        $UpdateToolStripMenuItem.Add_Click( { OnClick_UpdateToolStripMenuItem $UpdateToolStripMenuItem $EventArgs} )
        
        #########################
        # ModulesToolStripMenuItem #
        #########################
        $ModulesToolStripMenuItem.Name = "ModulesToolStripMenuItem"
        $ModulesToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $ModulesToolStripMenuItem.Text = "&Update Modules"
        function OnClick_ModulesToolStripMenuItem($Sender,$e){
           [void][System.Windows.Forms.MessageBox]::Show("Update Modules", "BWApp v$version")
        }
        $ModulesToolStripMenuItem.Add_Click( { OnClick_ModulesToolStripMenuItem $ModulesToolStripMenuItem $EventArgs} )


##########################
# HelpToolStripMenuItem #
##########################
$HelpToolStripMenuItem.DropDownItems.AddRange(@(
$AboutToolStripMenuItem))
$HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
$HelpToolStripMenuItem.Size = new-object System.Drawing.Size(67, 20)
$HelpToolStripMenuItem.Text = "&Help"
        
        #########################
        # AboutToolStripMenuItem #
        #########################
        $AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        $AboutToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
        $AboutToolStripMenuItem.Text = "&About"
        function OnClick_AboutToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("
            BWApp
            Version $version Build $build
            Email:   dale@dcnet.nz
            Mobile:  +61426005007
            Github:  $github
                    ", "About BWAppv3")
        }

        $AboutToolStripMenuItem.Add_Click( { OnClick_AboutToolStripMenuItem $AboutToolStripMenuItem $EventArgs} )





##########################
# ReportsToolStripMenuItem #
##########################
$ReportsToolStripMenuItem.DropDownItems.AddRange(@(
$OverviewToolStripMenuItem,
$MailPermsToolStripMenuItem,
$AllMailPermsToolStripMenuItem,
$CheckForwardsToolStripMenuItem,
$DistMembersToolStripMenuItem,
$MailLogs48ToolStripMenuItem,
$MailLogs24ToolStripMenuItem,
$LicensedUsersToolStripMenuItem,
$ListEmailsToolStripMenuItem,
$CalendarPermsToolStripMenuItem))
$ReportsToolStripMenuItem.Name = "ReportsToolStripMenuItem"
$ReportsToolStripMenuItem.Size = new-object System.Drawing.Size(168, 20)
$ReportsToolStripMenuItem.Text = "&Reports"
        
        #############################
        # OverviewToolStripMenuItem #
        #############################
        $OverviewToolStripMenuItem.Name = "OverviewToolStripMenuItem"
        $OverviewToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $OverviewToolStripMenuItem.Text = "&Overview Report"
        function OnClick_OverviewToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Overview Report", "BWApp v$version")

            .\scripts\ReportsOverview.ps1
        }

        $OverviewToolStripMenuItem.Add_Click( { OnClick_OverviewToolStripMenuItem $OverviewToolStripMenuItem $EventArgs} )

        
        #############################
        # MailPermsToolStripMenuItem #
        #############################
        $MailPermsToolStripMenuItem.Name = "MailPermsToolStripMenuItem"
        $MailPermsToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $MailPermsToolStripMenuItem.Text = "&Mailbox Perms (One)"
        function OnClick_MailPermsToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Mail Perms 1", "BWApp v$version")
        }

        $MailPermsToolStripMenuItem.Add_Click( { OnClick_MailPermsToolStripMenuItem $MailPermsToolStripMenuItem $EventArgs} )

        
        #############################
        # AllMailPermsToolStripMenuItem #
        #############################
        $AllMailPermsToolStripMenuItem.Name = "AllMailPermsToolStripMenuItem"
        $AllMailPermsToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $AllMailPermsToolStripMenuItem.Text = "&Mailbox Perms (All)"
        function OnClick_AllMailPermsToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Mail Perms All", "BWApp v$version")
        }

        $AllMailPermsToolStripMenuItem.Add_Click( { OnClick_AllMailPermsToolStripMenuItem $AllMailPermsToolStripMenuItem $EventArgs} )

        
        #############################
        # CheckForwardsToolStripMenuItem #
        #############################
        $CheckForwardsToolStripMenuItem.Name = "CheckForwardsToolStripMenuItem"
        $CheckForwardsToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $CheckForwardsToolStripMenuItem.Text = "&Check Forwards"
        function OnClick_CheckForwardsToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Check Forwards", "BWApp v$version")
        }

        $CheckForwardsToolStripMenuItem.Add_Click( { OnClick_CheckForwardsToolStripMenuItem $CheckForwardsToolStripMenuItem $EventArgs} )

        
        #############################
        # DistMembersToolStripMenuItem #
        #############################
        $DistMembersToolStripMenuItem.Name = "DistMembersToolStripMenuItem"
        $DistMembersToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $DistMembersToolStripMenuItem.Text = "&Dist Grp Members"
        function OnClick_DistMembersToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Distribution Group Members", "BWApp v$version")
        }

        $DistMembersToolStripMenuItem.Add_Click( { OnClick_DistMembersToolStripMenuItem $DistMembersToolStripMenuItem $EventArgs} )

        
        #############################
        # MailLogs48ToolStripMenuItem #
        #############################
        $MailLogs48ToolStripMenuItem.Name = "MailLogs48ToolStripMenuItem"
        $MailLogs48ToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $MailLogs48ToolStripMenuItem.Text = "&Mail Logs (48 Hrs)"
        function OnClick_MailLogs48ToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Mail Logs for last 48 Hrs", "BWApp v$version")
        }

        $MailLogs48ToolStripMenuItem.Add_Click( { OnClick_MailLogs48ToolStripMenuItem $MailLogs48ToolStripMenuItem $EventArgs} )

        
        #############################
        # MailLogs24ToolStripMenuItem #
        #############################
        $MailLogs24ToolStripMenuItem.Name = "MailLogs24ToolStripMenuItem"
        $MailLogs24ToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $MailLogs24ToolStripMenuItem.Text = "&Mail Logs (24 Hrs)"
        function OnClick_MailLogs24ToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Mail Logs for last 24 Hrs", "BWApp v$version")
        }

        $MailLogs24ToolStripMenuItem.Add_Click( { OnClick_MailLogs24ToolStripMenuItem $MailLogs24ToolStripMenuItem $EventArgs} )

        
        #############################
        # LicensedUsersToolStripMenuItem #
        #############################
        $LicensedUsersToolStripMenuItem.Name = "LicensedUsersToolStripMenuItem"
        $LicensedUsersToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $LicensedUsersToolStripMenuItem.Text = "&Licensed Users"
        function OnClick_LicensedUsersToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Licensed Users", "BWApp v$version")
        }

        $LicensedUsersToolStripMenuItem.Add_Click( { OnClick_LicensedUsersToolStripMenuItem $LicensedUsersToolStripMenuItem $EventArgs} )

        
        #############################
        # ListEmailsToolStripMenuItem #
        #############################
        $ListEmailsToolStripMenuItem.Name = "ListEmailsToolStripMenuItem"
        $ListEmailsToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $ListEmailsToolStripMenuItem.Text = "&List Email Addresses"
        function OnClick_ListEmailsToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("List Email Addresses", "BWApp v$version")
        }

        $ListEmailsToolStripMenuItem.Add_Click( { OnClick_ListEmailsToolStripMenuItem $ListEmailsToolStripMenuItem $EventArgs} )

        
        #############################
        # CalendarPermsToolStripMenuItem #
        #############################
        $CalendarPermsToolStripMenuItem.Name = "CalendarPermsToolStripMenuItem"
        $CalendarPermsToolStripMenuItem.Size = new-object System.Drawing.Size(168, 22)
        $CalendarPermsToolStripMenuItem.Text = "&Calendar Perms"
        function OnClick_CalendarPermsToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Calendar Permissions", "BWApp v$version")
        }

        $CalendarPermsToolStripMenuItem.Add_Click( { OnClick_CalendarPermsToolStripMenuItem $CalendarPermsToolStripMenuItem $EventArgs} )

#################


##########################
# TenantToolStripMenuItem #
##########################
$TenantToolStripMenuItem.DropDownItems.AddRange(@(
$TenancyPrepWizToolStripMenuItem,
$OrgCustomizationToolStripMenuItem,
$AuditLogsToolStripMenuItem,
$POPIMAPToolStripMenuItem,
$IntuneImportToolStripMenuItem,
$OutsideSenderToolStripMenuItem,
$SpamPolicyToolStripMenuItem,
$MalwarePolicyToolStripMenuItem))
$TenantToolStripMenuItem.Name = "TenantToolStripMenuItem"
$TenantToolStripMenuItem.Size = new-object System.Drawing.Size(184, 20)
$TenantToolStripMenuItem.Text = "&Tenancy"
        
        #############################
        # TenancyPrepWizToolStripMenuItem #
        #############################
        $TenancyPrepWizToolStripMenuItem.Name = "TenancyPrepWizToolStripMenuItem"
        $TenancyPrepWizToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $TenancyPrepWizToolStripMenuItem.Text = "&Prep Tenancy"
        function OnClick_TenancyPrepWizToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Prep Tenancy", "BWApp v$version")
        }

        $TenancyPrepWizToolStripMenuItem.Add_Click( { OnClick_TenancyPrepWizToolStripMenuItem $TenancyPrepWizToolStripMenuItem $EventArgs} )

        
        #############################
        # OrgCustomizationToolStripMenuItem #
        #############################
        $OrgCustomizationToolStripMenuItem.Name = "OrgCustomizationToolStripMenuItem"
        $OrgCustomizationToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $OrgCustomizationToolStripMenuItem.Text = "&Enable Org Customization"
        function OnClick_OrgCustomizationToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Enable Org Customization", "BWApp v$version")
        }

        $OrgCustomizationToolStripMenuItem.Add_Click( { OnClick_OrgCustomizationToolStripMenuItem $OrgCustomizationToolStripMenuItem $EventArgs} )

        
        #############################
        # AuditLogsToolStripMenuItem #
        #############################
        $AuditLogsToolStripMenuItem.Name = "AuditLogsToolStripMenuItem"
        $AuditLogsToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $AuditLogsToolStripMenuItem.Text = "&Enable Audit Logs"
        function OnClick_AuditLogsToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Enable Audit Logging", "BWApp v$version")
        }

        $AuditLogsToolStripMenuItem.Add_Click( { OnClick_AuditLogsToolStripMenuItem $AuditLogsToolStripMenuItem $EventArgs} )

        
        #############################
        # POPIMAPToolStripMenuItem #
        #############################
        $POPIMAPToolStripMenuItem.Name = "POPIMAPToolStripMenuItem"
        $POPIMAPToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $POPIMAPToolStripMenuItem.Text = "&Disable POP/IMAP"
        function OnClick_POPIMAPToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Disable POP / IMAP", "BWApp v$version")
        }

        $POPIMAPToolStripMenuItem.Add_Click( { OnClick_POPIMAPToolStripMenuItem $POPIMAPToolStripMenuItem $EventArgs} )

        
        #############################
        # IntuneImportToolStripMenuItem #
        #############################
        $IntuneImportToolStripMenuItem.Name = "IntuneImportToolStripMenuItem"
        $IntuneImportToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $IntuneImportToolStripMenuItem.Text = "&Import Intune Policies"
        function OnClick_IntuneImportToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Import Intune Policies", "BWApp v$version")
        }

        $IntuneImportToolStripMenuItem.Add_Click( { OnClick_IntuneImportToolStripMenuItem $IntuneImportToolStripMenuItem $EventArgs} )

        
        #############################
        # OutsideSenderToolStripMenuItem #
        #############################
        $OutsideSenderToolStripMenuItem.Name = "OutsideSenderToolStripMenuItem"
        $OutsideSenderToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $OutsideSenderToolStripMenuItem.Text = "&Outside Sender"
        function OnClick_OutsideSenderToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Enable Outside Senders Rule", "BWApp v$version")
        }

        $OutsideSenderToolStripMenuItem.Add_Click( { OnClick_OutsideSenderToolStripMenuItem $OutsideSenderToolStripMenuItem $EventArgs} )

        
        #############################
        # SpamPolicyToolStripMenuItem #
        #############################
        $SpamPolicyToolStripMenuItem.Name = "SpamPolicyToolStripMenuItem"
        $SpamPolicyToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $SpamPolicyToolStripMenuItem.Text = "&Create Spam Policy"
        function OnClick_SpamPolicyToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Enable Spam Policy", "BWApp v$version")
        }

        $SpamPolicyToolStripMenuItem.Add_Click( { OnClick_SpamPolicyToolStripMenuItem $SpamPolicyToolStripMenuItem $EventArgs} )

        
        #############################
        # MalwarePolicyToolStripMenuItem #
        #############################
        $MalwarePolicyToolStripMenuItem.Name = "MalwarePolicyToolStripMenuItem"
        $MalwarePolicyToolStripMenuItem.Size = new-object System.Drawing.Size(184, 22)
        $MalwarePolicyToolStripMenuItem.Text = "&Create Malware Policy"
        function OnClick_MalwarePolicyToolStripMenuItem($Sender,$e){
            [void][System.Windows.Forms.MessageBox]::Show("Enable Malware Policy", "BWApp v$version")
        }

        $MalwarePolicyToolStripMenuItem.Add_Click( { OnClick_MalwarePolicyToolStripMenuItem $MalwarePolicyToolStripMenuItem $EventArgs} )







$monitor = [System.Windows.Forms.Screen]::PrimaryScreen
$Width = $monitor.WorkingArea.Width
$Height = $monitor.WorkingArea.Height

###################################################################
##MENU FORM#
$MenuForm = new-object System.Windows.Forms.form
#
$MenuForm.ClientSize = new-object System.Drawing.Size($Width, 25)
$MenuForm.Controls.Add($MS_Main)
$MenuForm.MainMenuStrip = $MS_Main
$MenuForm.Name = "MenuForm"
$MenuForm.StartPosition = "0"
$MenuForm.Text = "BWApp v$version Build $build"
function OnFormClosing_MenuForm($Sender,$e){ 
    # $this represent sender (object)
    # $_ represent  e (eventarg)

    # Allow closing
    ($_).Cancel= $False
}
$MenuForm.Add_FormClosing( { OnFormClosing_MenuForm $MenuForm $EventArgs} )
$MenuForm.Add_Shown({$MenuForm.Activate()})

########

$Path = Get-Content -path "C:\Temp\path.bwapp"
Set-Location -path "$Path"


write-host BWApp
Write-host Version: $version
Write-host Build: $build
write-host $github
write-host -ForegroundColor Cyan 'READY...'
Write-Host "$path"
########


$MenuForm.ShowDialog()
#Free ressources
$MenuForm.Dispose()
