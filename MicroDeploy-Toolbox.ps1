<# 
 MicroDeploy Toolbox – GUI App Selector
 v0.9 – Profiles, winget/choco backends, tooltips, runtime-aware update/uninstall, cleanup,
        Chocolatey auto-install/auto-upgrade, safer winget detection, per-app progress + summary popup

 - Windows 10 / 11
 - Profiles: Gamer, Everyday user, Power user, SysAdmin
 - Profile controls which apps are visible (all start UNCHECKED)
 - Buttons:
     * Install   -> winget / choco install
     * Update    -> winget upgrade / choco upgrade
         - If app has known processes/services:
             * Prompt to stop them for update
             * After successful update, try to relaunch what was stopped
     * Uninstall -> winget uninstall / choco uninstall
         - Stop known services/processes
         - Run uninstall
         - Run targeted cleanup: folders + registry keys from CleanupTokens
 - Extras:
     * Hover tooltips per app
     * "Working..." overlay for long actions and startup init
     * Installed-status detection (via winget list)
     * Backend selector (Winget / Chocolatey)
     * Select All / Clear visible apps
     * Save / Load profiles as JSON
     * Chocolatey auto-install / auto-upgrade core package
     * Per-app progress text + completion summary popup
#>

# Relaunch as STA if needed (required for WPF)
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File $PSCommandPath @args
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

# Are we running elevated?
$script:IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

# Track one-time checks
$script:ChocoChecked   = $false
$global:WingetPromptShown = $false

# ---------------------------
# XAML UI
# ---------------------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MicroDeploy Toolbox" Height="620" Width="980"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="12">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Top bar -->
    <DockPanel Grid.Row="0" Margin="0,0,0,8">
      <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
        <TextBlock Text="Profile:" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <ComboBox x:Name="cmbProfile" Width="180" SelectedIndex="0">
          <ComboBoxItem Content="Gamer"/>
          <ComboBoxItem Content="Everyday user"/>
          <ComboBoxItem Content="Power user"/>
          <ComboBoxItem Content="SysAdmin"/>
        </ComboBox>

        <TextBlock Text="Backend:" VerticalAlignment="Center" Margin="12,0,4,0"/>
        <ComboBox x:Name="cmbBackend" Width="140" SelectedIndex="0">
          <ComboBoxItem Content="Winget"/>
          <ComboBoxItem Content="Chocolatey"/>
        </ComboBox>
      </StackPanel>

      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
        <TextBlock Text="Profile = visible apps. Backend = installer."
                   VerticalAlignment="Center" Foreground="#666" Margin="0,0,12,0"/>
        <Button x:Name="btnSaveProfile" Content="Save Profile" Width="100" Margin="0,0,4,0"/>
        <Button x:Name="btnLoadProfile" Content="Load Profile" Width="100"/>
      </StackPanel>
    </DockPanel>

    <!-- Main content -->
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="2*"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- Tool list -->
      <GroupBox Header="Available Tools" Grid.Column="0" Margin="0,0,8,0">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <UniformGrid Columns="2" Margin="6">

            <!-- General / Browsers -->
            <CheckBox x:Name="chk_NotepadPP" Content="Notepad++" Margin="2"/>
            <CheckBox x:Name="chk_VSCode" Content="Visual Studio Code" Margin="2"/>
            <CheckBox x:Name="chk_Chrome" Content="Google Chrome" Margin="2"/>
            <CheckBox x:Name="chk_Firefox" Content="Mozilla Firefox" Margin="2"/>
            <CheckBox x:Name="chk_Edge" Content="Microsoft Edge" Margin="2"/>
            <CheckBox x:Name="chk_7zip" Content="7-Zip" Margin="2"/>
            <CheckBox x:Name="chk_VLC" Content="VLC media player" Margin="2"/>
            <CheckBox x:Name="chk_ShareX" Content="ShareX" Margin="2"/>
            <CheckBox x:Name="chk_Everything" Content="Everything (voidtools)" Margin="2"/>

            <CheckBox x:Name="chk_LibreOffice" Content="LibreOffice" Margin="2"/>
            <CheckBox x:Name="chk_AcrobatReader" Content="Adobe Acrobat Reader" Margin="2"/>
            <CheckBox x:Name="chk_Spotify" Content="Spotify" Margin="2"/>
            <CheckBox x:Name="chk_Zoom" Content="Zoom" Margin="2"/>

            <!-- Dev / Power -->
            <CheckBox x:Name="chk_Git" Content="Git" Margin="2"/>
            <CheckBox x:Name="chk_GitHubDesktop" Content="GitHub Desktop" Margin="2"/>
            <CheckBox x:Name="chk_PowerToys" Content="Microsoft PowerToys" Margin="2"/>
            <CheckBox x:Name="chk_Terminal" Content="Windows Terminal" Margin="2"/>
            <CheckBox x:Name="chk_Putty" Content="PuTTY" Margin="2"/>
            <CheckBox x:Name="chk_WinSCP" Content="WinSCP" Margin="2"/>
            <CheckBox x:Name="chk_DockerDesktop" Content="Docker Desktop" Margin="2"/>
            <CheckBox x:Name="chk_Postman" Content="Postman" Margin="2"/>
            <CheckBox x:Name="chk_AutoHotkey" Content="AutoHotkey" Margin="2"/>
            <CheckBox x:Name="chk_Python" Content="Python 3" Margin="2"/>
            <CheckBox x:Name="chk_NodeLTS" Content="Node.js LTS" Margin="2"/>

            <!-- Gaming / Social -->
            <CheckBox x:Name="chk_Steam" Content="Steam" Margin="2"/>
            <CheckBox x:Name="chk_Epic" Content="Epic Games Launcher" Margin="2"/>
            <CheckBox x:Name="chk_Discord" Content="Discord" Margin="2"/>
            <CheckBox x:Name="chk_Battlenet" Content="Battle.net" Margin="2"/>
            <CheckBox x:Name="chk_Origin" Content="EA App / Origin" Margin="2"/>
            <CheckBox x:Name="chk_GOG" Content="GOG Galaxy" Margin="2"/>
            <CheckBox x:Name="chk_OBS" Content="OBS Studio" Margin="2"/>
            <CheckBox x:Name="chk_FanControl" Content="FanControl" Margin="2"/>
            <CheckBox x:Name="chk_HWiNFO" Content="HWiNFO64" Margin="2"/>
            <CheckBox x:Name="chk_MSIAB" Content="MSI Afterburner" Margin="2"/>

            <!-- Sysadmin / Tools -->
            <CheckBox x:Name="chk_SysInternals" Content="Sysinternals Suite" Margin="2"/>
            <CheckBox x:Name="chk_Wireshark" Content="Wireshark" Margin="2"/>
            <CheckBox x:Name="chk_RDCMan" Content="Remote Desktop Client" Margin="2"/>
            <CheckBox x:Name="chk_VSBuildTools" Content="VS Build Tools" Margin="2"/>
            <CheckBox x:Name="chk_Notepad3" Content="Notepad3" Margin="2"/>
            <CheckBox x:Name="chk_PowerShell7" Content="PowerShell 7" Margin="2"/>
            <CheckBox x:Name="chk_AzureCLI" Content="Azure CLI" Margin="2"/>
            <CheckBox x:Name="chk_AWSCLI" Content="AWS CLI" Margin="2"/>
            <CheckBox x:Name="chk_Terraform" Content="Terraform" Margin="2"/>

            <!-- Extra Utilities / Disk / Security / Comms -->
            <CheckBox x:Name="chk_Malwarebytes" Content="Malwarebytes" Margin="2"/>
            <CheckBox x:Name="chk_Rufus" Content="Rufus (USB creator)" Margin="2"/>
            <CheckBox x:Name="chk_CrystalDiskInfo" Content="CrystalDiskInfo" Margin="2"/>
            <CheckBox x:Name="chk_CrystalDiskMark" Content="CrystalDiskMark" Margin="2"/>
            <CheckBox x:Name="chk_TreeSizeFree" Content="TreeSize Free" Margin="2"/>
            <CheckBox x:Name="chk_CPUZ" Content="CPU-Z" Margin="2"/>
            <CheckBox x:Name="chk_VirtualBox" Content="Oracle VirtualBox" Margin="2"/>
            <CheckBox x:Name="chk_Teams" Content="Microsoft Teams" Margin="2"/>
            <CheckBox x:Name="chk_Slack" Content="Slack" Margin="2"/>
            <CheckBox x:Name="chk_Telegram" Content="Telegram Desktop" Margin="2"/>

          </UniformGrid>
        </ScrollViewer>
      </GroupBox>

      <!-- Info / Log -->
      <GroupBox Header="Output / Log" Grid.Column="1">
        <Grid Margin="4">
          <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <TextBox x:Name="txtLog"
                   Grid.Row="0"
                   Margin="0,0,0,4"
                   IsReadOnly="True"
                   TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto"
                   AcceptsReturn="True" />

          <TextBlock Grid.Row="1"
                     Foreground="#666"
                     FontSize="11"
                     Text="Install / update / uninstall actions, backend info, and profile loads will appear here."/>
        </Grid>
      </GroupBox>
    </Grid>

    <!-- Buttons -->
    <StackPanel Grid.Row="2"
                Orientation="Horizontal"
                HorizontalAlignment="Right"
                Margin="0,8,0,0">
      <Button x:Name="btnSelectAll" Content="Select All" Width="90" Margin="0,0,6,0"/>
      <Button x:Name="btnClearAll" Content="Clear" Width="80" Margin="0,0,18,0"/>
      <Button x:Name="btnInstall" Content="Install" Width="90" Margin="0,0,6,0"/>
      <Button x:Name="btnUpdate" Content="Update" Width="90" Margin="0,0,6,0"/>
      <Button x:Name="btnUninstall" Content="Uninstall" Width="90" Margin="0,0,6,0"/>
      <Button x:Name="btnClose" Content="Close" Width="80"/>
    </StackPanel>

    <!-- Busy overlay -->
    <Border x:Name="busyOverlay"
            Grid.RowSpan="3"
            Background="#80000000"
            Visibility="Collapsed">
      <Border HorizontalAlignment="Center"
              VerticalAlignment="Center"
              Width="260"
              Background="#FF1E1E1E"
              Opacity="0.95"
              Padding="20">
        <StackPanel>
          <TextBlock x:Name="txtBusy"
                     Text="Working..."
                     Foreground="White"
                     FontSize="14"
                     FontWeight="Bold"
                     TextAlignment="Center"
                     Margin="0,0,0,8"/>
          <ProgressBar IsIndeterminate="True" Height="18" Margin="0,0,0,4"/>
          <TextBlock Text="Please wait while operations complete."
                     Foreground="#CCCCCC"
                     FontSize="11"
                     TextAlignment="Center"/>
        </StackPanel>
      </Border>
    </Border>

  </Grid>
</Window>
"@

# ---------------------------
# Load XAML
# ---------------------------
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Grab controls
$cmbProfile     = $Window.FindName('cmbProfile')
$cmbBackend     = $Window.FindName('cmbBackend')
$txtLog         = $Window.FindName('txtLog')

$btnSaveProfile = $Window.FindName('btnSaveProfile')
$btnLoadProfile = $Window.FindName('btnLoadProfile')

$btnSelectAll   = $Window.FindName('btnSelectAll')
$btnClearAll    = $Window.FindName('btnClearAll')
$btnInstall     = $Window.FindName('btnInstall')
$btnUpdate      = $Window.FindName('btnUpdate')
$btnUninstall   = $Window.FindName('btnUninstall')
$btnClose       = $Window.FindName('btnClose')

$busyOverlay    = $Window.FindName('busyOverlay')
$txtBusy        = $Window.FindName('txtBusy')

# ---------------------------
# Helpers: logging & backend/profile getters
# ---------------------------
function Write-Log {
    param([string]$Message)
    if (-not $txtLog) { return }
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $txtLog.AppendText("[$timestamp] $Message`r`n")
    $txtLog.ScrollToEnd()
}

function Get-Backend {
    if (-not $cmbBackend) { return 'Winget' }
    $item = $cmbBackend.SelectedItem
    if ($item -and $item.Content) { $item.Content.ToString() } else { 'Winget' }
}

function Get-ProfileName {
    $item = $cmbProfile.SelectedItem
    if ($item -and $item.Content) { $item.Content.ToString() } else { '' }
}

# ---------------------------
# Helper: busy overlay
# ---------------------------
function Show-Busy {
    param(
        [string]$Message = "Working..."
    )
    if ($txtBusy) {
        $txtBusy.Text = $Message
    }
    if ($busyOverlay) {
        $busyOverlay.Visibility = 'Visible'
    }
    $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Hide-Busy {
    if ($busyOverlay) {
        $busyOverlay.Visibility = 'Collapsed'
    }
    $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

# ---------------------------
# App catalog (Profiles, IDs, descriptions)
# ---------------------------
$script:Apps = @()

function Add-App {
    param(
        [string]$Name,
        [string]$Key,
        [string]$WingetId,
        [string]$CheckboxName,
        [string[]]$Profiles,
        [string]$Description,
        [string]$ChocoId = $null
    )
    $script:Apps += New-Object PSObject -Property @{
        Name         = $Name
        Key          = $Key
        WingetId     = $WingetId
        ChocoId      = $ChocoId
        Checkbox     = $Window.FindName($CheckboxName)
        Profiles     = $Profiles
        Description  = $Description
        Installed    = $false
        Processes    = @()
        Services     = @()
        CleanupTokens= @()
    }
}

# General / browsers / utilities
Add-App 'Notepad++'            'NotepadPP'     'Notepad++.Notepad++'                 'chk_NotepadPP'     @('Everyday user','Power user','SysAdmin') 'Tabbed text editor for quick config edits and coding.'        'notepadplusplus'
Add-App 'Visual Studio Code'   'VSCode'        'Microsoft.VisualStudioCode'          'chk_VSCode'        @('Power user','SysAdmin')                  'Extensible code editor for almost any language.'              'vscode'
Add-App 'Google Chrome'        'Chrome'        'Google.Chrome'                       'chk_Chrome'        @('Gamer','Everyday user','Power user','SysAdmin') 'Mainstream Chromium browser with wide support.'       'googlechrome'
Add-App 'Mozilla Firefox'      'Firefox'       'Mozilla.Firefox'                     'chk_Firefox'       @('Everyday user','Power user','SysAdmin')  'Privacy-focused browser.'                                    'firefox'
Add-App 'Microsoft Edge'       'Edge'          'Microsoft.Edge'                      'chk_Edge'          @('Everyday user','Power user','SysAdmin')  'Bundled Chromium-based browser.'
Add-App '7-Zip'                '7zip'          '7zip.7zip'                           'chk_7zip'          @('Gamer','Everyday user','Power user','SysAdmin') 'Archive manager for ZIP/7z/RAR and more.'           '7zip'
Add-App 'VLC media player'     'VLC'           'VideoLAN.VLC'                        'chk_VLC'           @('Gamer','Everyday user','Power user')     'Plays practically any media format.'                         'vlc'
Add-App 'ShareX'               'ShareX'        'ShareX.ShareX'                       'chk_ShareX'        @('Everyday user','Power user')             'Advanced screenshot and screen recording tool.'              'sharex'
Add-App 'Everything (voidtools)' 'Everything'  'voidtools.Everything'                'chk_Everything'    @('Power user','SysAdmin')                  'Instant file search for NTFS volumes.'                       'everything'
Add-App 'LibreOffice'          'LibreOffice'   'TheDocumentFoundation.LibreOffice'   'chk_LibreOffice'   @('Everyday user','Power user')             'Open-source office suite.'                                  'libreoffice-fresh'
Add-App 'Adobe Acrobat Reader' 'AcrobatReader' 'Adobe.Acrobat.Reader.64-bit'        'chk_AcrobatReader' @('Everyday user','Power user','SysAdmin')  'Standard PDF reader.'                                       'adobereader'
Add-App 'Spotify'              'Spotify'       'Spotify.Spotify'                     'chk_Spotify'       @('Gamer','Everyday user')                  'Music streaming client.'                                    'spotify'
Add-App 'Zoom'                 'Zoom'          'Zoom.Zoom'                           'chk_Zoom'          @('Everyday user','Power user')             'Video conferencing client.'                                 'zoom'

# Dev / Power tools
Add-App 'Git'                  'Git'           'Git.Git'                             'chk_Git'           @('Power user','SysAdmin')                  'Distributed version control.'                               'git'
Add-App 'GitHub Desktop'       'GitHubDesktop' 'GitHub.GitHubDesktop'                'chk_GitHubDesktop' @('Power user')                             'Git GUI for GitHub.'                                       'github-desktop'
Add-App 'Microsoft PowerToys'  'PowerToys'     'Microsoft.PowerToys'                 'chk_PowerToys'     @('Everyday user','Power user','SysAdmin')  'Power-user utilities (FancyZones, launcher, etc.).'        'powertoys'
Add-App 'Windows Terminal'     'Terminal'      'Microsoft.WindowsTerminal'           'chk_Terminal'      @('Power user','SysAdmin')                  'Modern multi-tab terminal.'                                'microsoft-windows-terminal'
Add-App 'PuTTY'                'Putty'         'SimonTatham.Putty'                   'chk_Putty'         @('Power user','SysAdmin')                  'Classic SSH/Telnet client.'                                'putty'
Add-App 'WinSCP'               'WinSCP'        'WinSCP.WinSCP'                       'chk_WinSCP'        @('Power user','SysAdmin')                  'SFTP/SCP/FTP client with scripting.'                       'winscp'
Add-App 'Docker Desktop'       'DockerDesktop' 'Docker.DockerDesktop'                'chk_DockerDesktop' @('Power user','SysAdmin')                  'Local Docker runtime and GUI.'                             'docker-desktop'
Add-App 'Postman'              'Postman'       'Postman.Postman'                     'chk_Postman'       @('Power user','SysAdmin')                  'API testing and automation client.'                        'postman'
Add-App 'AutoHotkey'           'AutoHotkey'    'AutoHotkey.AutoHotkey'               'chk_AutoHotkey'    @('Power user')                             'Macro and automation scripting language.'                  'autohotkey'
Add-App 'Python 3'             'Python3'       'Python.Python.3'                     'chk_Python'        @('Power user','SysAdmin')                  'Python runtime for scripting and apps.'                    'python'
Add-App 'Node.js LTS'          'NodeLTS'       'OpenJS.NodeJS.LTS'                   'chk_NodeLTS'       @('Power user')                             'LTS Node.js + npm for tooling.'                            'nodejs-lts'

# Gaming / Social
Add-App 'Steam'                'Steam'         'Valve.Steam'                         'chk_Steam'         @('Gamer')                                  'Primary PC game platform and store.'                       'steam'
Add-App 'Epic Games Launcher'  'Epic'          'EpicGames.EpicGamesLauncher'         'chk_Epic'          @('Gamer')                                  'Epic store/launcher (free weekly games).'                  'epicgameslauncher'
Add-App 'Discord'              'Discord'       'Discord.Discord'                     'chk_Discord'       @('Gamer','Everyday user','Power user')     'Voice/text chat for gaming & communities.'                 'discord'
Add-App 'Battle.net'           'Battlenet'     'Blizzard.BattleNet'                  'chk_Battlenet'     @('Gamer')                                  'Launcher for Blizzard games.'                              'battlenet'
Add-App 'EA App / Origin'      'Origin'        'ElectronicArts.EADesktop'            'chk_Origin'        @('Gamer')                                  'Electronic Arts launcher.'
Add-App 'GOG Galaxy'           'GOG'           'GOG.Galaxy'                          'chk_GOG'           @('Gamer')                                  'GOG game library and store.'                               'goggalaxy'
Add-App 'OBS Studio'           'OBS'           'OBSProject.OBSStudio'                'chk_OBS'           @('Gamer','Power user')                     'Streaming/recording suite.'                                'obs-studio'
Add-App 'FanControl'           'FanControl'    'Rem0o.FanControl'                    'chk_FanControl'    @('Gamer','Power user')                     'Fan curve controller for temperature tuning.'
Add-App 'HWiNFO64'             'HWiNFO'        'REALiX.HWiNFO'                       'chk_HWiNFO'        @('Gamer','Power user','SysAdmin')          'Detailed hardware sensor monitoring.'
Add-App 'MSI Afterburner'      'MSIAB'         'Guru3D.Afterburner'                  'chk_MSIAB'         @('Gamer')                                  'GPU overclocking and fan tuning.'

# Sysadmin / Diagnostics
Add-App 'Sysinternals Suite'   'SysInternals'  'Microsoft.SysinternalsSuite'         'chk_SysInternals'  @('SysAdmin')                               'Process Explorer, ProcMon, PsExec, and more.'              'sysinternals'
Add-App 'Wireshark'            'Wireshark'     'WiresharkFoundation.Wireshark'       'chk_Wireshark'     @('Power user','SysAdmin')                  'Network protocol analyzer.'                                'wireshark'
Add-App 'Remote Desktop Client' 'RDCMan'       'Microsoft.RemoteDesktopClient'       'chk_RDCMan'        @('SysAdmin')                               'RDP client for managing remote sessions.'                  'microsoft-remote-desktop'
Add-App 'VS Build Tools'       'VSBuildTools'  'Microsoft.VisualStudio.2022.BuildTools' 'chk_VSBuildTools' @('SysAdmin','Power user')                 'Visual Studio build chain without full IDE.'              'visualstudio2022buildtools'
Add-App 'Notepad3'             'Notepad3'      'Rizonesoft.Notepad3'                 'chk_Notepad3'      @('Power user','SysAdmin')                  'Notepad replacement with syntax highlighting.'            'notepad3'
Add-App 'PowerShell 7'         'PowerShell7'   'Microsoft.PowerShell'                'chk_PowerShell7'   @('Power user','SysAdmin')                  'Modern PowerShell Core runtime.'                           'powershell-core'
Add-App 'Azure CLI'            'AzureCLI'      'Microsoft.AzureCLI'                  'chk_AzureCLI'      @('SysAdmin')                               'CLI for Azure resources.'                                  'azure-cli'
Add-App 'AWS CLI'              'AWSCLI'        'Amazon.AWSCLI'                       'chk_AWSCLI'        @('SysAdmin')                               'CLI for AWS resources.'                                    'awscli'
Add-App 'Terraform'            'Terraform'     'Hashicorp.Terraform'                 'chk_Terraform'     @('SysAdmin')                               'Infrastructure-as-code tool.'                              'terraform'

# Extra utilities / disk / security / comms
Add-App 'Malwarebytes'         'Malwarebytes'  'Malwarebytes.Malwarebytes'           'chk_Malwarebytes'  @('Everyday user','Power user','SysAdmin')  'On-demand malware scanner and real-time protection.'       'malwarebytes'
Add-App 'Rufus'                'Rufus'         'Rufus.Rufus'                         'chk_Rufus'         @('Power user','SysAdmin')                  'Create bootable USB drives from ISO images.'               'rufus'
Add-App 'CrystalDiskInfo'      'CrystalDiskInfo' 'CrystalDewWorld.CrystalDiskInfo'   'chk_CrystalDiskInfo' @('Power user','SysAdmin')                'Monitor SMART health of HDDs/SSDs.'                        'crystaldiskinfo'
Add-App 'CrystalDiskMark'      'CrystalDiskMark' 'CrystalDewWorld.CrystalDiskMark'   'chk_CrystalDiskMark' @('Gamer','Power user','SysAdmin')        'Benchmark sequential/random disk performance.'             'crystaldiskmark'
Add-App 'TreeSize Free'        'TreeSizeFree'  'JAMSoftware.TreeSize.Free'           'chk_TreeSizeFree'  @('Power user','SysAdmin')                  'Visualize disk usage to find large folders.'               'treesizefree'
Add-App 'CPU-Z'                'CPUZ'          'CPUID.CPU-Z'                         'chk_CPUZ'          @('Gamer','Power user','SysAdmin')          'Detailed CPU, memory and board info.'                      'cpu-z.install'
Add-App 'Oracle VirtualBox'    'VirtualBox'    'Oracle.VirtualBox'                   'chk_VirtualBox'    @('Power user','SysAdmin')                  'Local virtualization platform for running VMs.'            'virtualbox'
Add-App 'Microsoft Teams'      'Teams'         'Microsoft.Teams'                     'chk_Teams'         @('Everyday user','Power user','SysAdmin')  'Collaboration, meetings, and chat client.'                 'microsoft-teams'
Add-App 'Slack'                'Slack'         'SlackTechnologies.Slack'             'chk_Slack'         @('Everyday user','Power user','SysAdmin')  'Workspace chat and collaboration tool.'                    'slack'
Add-App 'Telegram Desktop'     'Telegram'      'Telegram.TelegramDesktop'            'chk_Telegram'      @('Everyday user','Power user')             'Desktop client for Telegram messaging.'                    'telegram'

# --- attach process/service/cleanup metadata for apps ---
$meta = @{
    'Chrome' = @{
        Processes     = @('chrome')
        Services      = @()
        CleanupTokens = @('Google\Chrome', 'Chrome')
    }
    'Firefox' = @{
        Processes     = @('firefox')
        Services      = @()
        CleanupTokens = @('Mozilla\Firefox', 'Firefox')
    }
    'Edge' = @{
        Processes     = @('msedge')
        Services      = @()
        CleanupTokens = @('Microsoft\Edge')
    }
    'VSCode' = @{
        Processes     = @('Code')
        Services      = @()
        CleanupTokens = @('Microsoft VS Code', 'VSCode', 'Code')
    }
    'NotepadPP' = @{
        Processes     = @('notepad++')
        Services      = @()
        CleanupTokens = @('Notepad++')
    }
    'VLC' = @{
        Processes     = @('vlc')
        Services      = @()
        CleanupTokens = @('VideoLAN\VLC', 'VLC')
    }
    '7zip' = @{
        Processes     = @('7zFM','7zG')
        Services      = @()
        CleanupTokens = @('7-Zip')
    }
    'ShareX' = @{
        Processes     = @('ShareX')
        Services      = @()
        CleanupTokens = @('ShareX')
    }
    'Everything' = @{
        Processes     = @('Everything')
        Services      = @('Everything')
        CleanupTokens = @('Everything')
    }
    'LibreOffice' = @{
        Processes     = @('soffice','soffice.bin')
        Services      = @()
        CleanupTokens = @('LibreOffice')
    }
    'AcrobatReader' = @{
        Processes     = @('AcroRd32')
        Services      = @()
        CleanupTokens = @('Adobe\Acrobat Reader','Acrobat Reader')
    }
    'Spotify' = @{
        Processes     = @('Spotify')
        Services      = @()
        CleanupTokens = @('Spotify')
    }
    'Zoom' = @{
        Processes     = @('Zoom')
        Services      = @()
        CleanupTokens = @('Zoom')
    }

    'Git' = @{
        Processes     = @()
        Services      = @()
        CleanupTokens = @('Git','GitHub\Git')
    }
    'GitHubDesktop' = @{
        Processes     = @('GitHubDesktop')
        Services      = @()
        CleanupTokens = @('GitHub Desktop')
    }
    'PowerToys' = @{
        Processes     = @('PowerToys','PowerToys.Settings')
        Services      = @()
        CleanupTokens = @('Microsoft\PowerToys')
    }
    'Terminal' = @{
        Processes     = @('WindowsTerminal')
        Services      = @()
        CleanupTokens = @('Microsoft\Windows Terminal')
    }
    'Putty' = @{
        Processes     = @('putty','pageant','plink')
        Services      = @()
        CleanupTokens = @('PuTTY')
    }
    'WinSCP' = @{
        Processes     = @('WinSCP')
        Services      = @()
        CleanupTokens = @('WinSCP')
    }
    'DockerDesktop' = @{
        Processes     = @('Docker Desktop','com.docker.backend')
        Services      = @('com.docker.service')
        CleanupTokens = @('Docker','Docker Desktop')
    }
    'Postman' = @{
        Processes     = @('Postman')
        Services      = @()
        CleanupTokens = @('Postman')
    }
    'AutoHotkey' = @{
        Processes     = @('AutoHotkey','AutoHotkeyU64')
        Services      = @()
        CleanupTokens = @('AutoHotkey')
    }
    'Python3' = @{
        Processes     = @('python','pythonw')
        Services      = @()
        CleanupTokens = @('Python','Python Software Foundation')
    }
    'NodeLTS' = @{
        Processes     = @('node')
        Services      = @()
        CleanupTokens = @('nodejs')
    }

    'Steam' = @{
        Processes     = @('steam')
        Services      = @()
        CleanupTokens = @('Valve\Steam','Steam')
    }
    'Epic' = @{
        Processes     = @('EpicGamesLauncher')
        Services      = @()
        CleanupTokens = @('Epic Games','EpicGamesLauncher')
    }
    'Battlenet' = @{
        Processes     = @('Battle.net','Agent')
        Services      = @()
        CleanupTokens = @('Battle.net','Blizzard Entertainment')
    }
    'Origin' = @{
        Processes     = @('EADesktop','Origin')
        Services      = @()
        CleanupTokens = @('Electronic Arts','Origin','EA Desktop')
    }
    'GOG' = @{
        Processes     = @('GalaxyClient','GOG Galaxy')
        Services      = @()
        CleanupTokens = @('GOG Galaxy')
    }
    'Discord' = @{
        Processes     = @('Discord')
        Services      = @()
        CleanupTokens = @('Discord')
    }
    'OBS' = @{
        Processes     = @('obs64')
        Services      = @()
        CleanupTokens = @('OBS Studio','obs-studio')
    }
    'FanControl' = @{
        Processes     = @('FanControl')
        Services      = @()
        CleanupTokens = @('FanControl')
    }
    'HWiNFO' = @{
        Processes     = @('HWiNFO64')
        Services      = @()
        CleanupTokens = @('HWiNFO64')
    }
    'MSIAB' = @{
        Processes     = @('MSIAfterburner')
        Services      = @()
        CleanupTokens = @('MSI Afterburner')
    }

    'SysInternals' = @{
        Processes     = @('procexp','procmon','PsExec')
        Services      = @()
        CleanupTokens = @('Sysinternals','SysInternals')
    }
    'Wireshark' = @{
        Processes     = @('Wireshark')
        Services      = @('Wireshark')
        CleanupTokens = @('Wireshark')
    }
    'RDCMan' = @{
        Processes     = @('mstsc')
        Services      = @()
        CleanupTokens = @('Remote Desktop','Terminal Services Client')
    }
    'VSBuildTools' = @{
        Processes     = @('MSBuild','devenv')
        Services      = @()
        CleanupTokens = @('Microsoft Visual Studio\2022\BuildTools','VisualStudio\17.0')
    }
    'Notepad3' = @{
        Processes     = @('Notepad3')
        Services      = @()
        CleanupTokens = @('Notepad3')
    }
    'PowerShell7' = @{
        Processes     = @('pwsh')
        Services      = @()
        CleanupTokens = @('PowerShell\7','PowerShell-7')
    }
    'AzureCLI' = @{
        Processes     = @('az')
        Services      = @()
        CleanupTokens = @('Microsoft SDKs\Azure','AzureCLI')
    }
    'AWSCLI' = @{
        Processes     = @('aws')
        Services      = @()
        CleanupTokens = @('Amazon\AWSCLI','AWSCLI')
    }
    'Terraform' = @{
        Processes     = @('terraform')
        Services      = @()
        CleanupTokens = @('Hashicorp\Terraform','terraform')
    }

    'Malwarebytes' = @{
        Processes     = @('mbam','mbamtray')
        Services      = @('MBAMService')
        CleanupTokens = @('Malwarebytes','MBAMService')
    }
    'Rufus' = @{
        Processes     = @('rufus')
        Services      = @()
        CleanupTokens = @('Rufus')
    }
    'CrystalDiskInfo' = @{
        Processes     = @('DiskInfo64','DiskInfo32')
        Services      = @()
        CleanupTokens = @('CrystalDiskInfo')
    }
    'CrystalDiskMark' = @{
        Processes     = @('DiskMark64','DiskMark32')
        Services      = @()
        CleanupTokens = @('CrystalDiskMark')
    }
    'TreeSizeFree' = @{
        Processes     = @('TreeSizeFree')
        Services      = @()
        CleanupTokens = @('TreeSize Free','JAM Software\TreeSize Free')
    }
    'CPUZ' = @{
        Processes     = @('cpuz')
        Services      = @()
        CleanupTokens = @('CPUID\CPU-Z')
    }
    'VirtualBox' = @{
        Processes     = @('VirtualBox','VBoxSVC')
        Services      = @('VBoxDrv','VBoxSup','VBoxUSBMon','VBoxNetLwf')
        CleanupTokens = @('Oracle\VirtualBox','VirtualBox')
    }
    'Teams' = @{
        Processes     = @('ms-teams','Teams')
        Services      = @()
        CleanupTokens = @('Microsoft\Teams','Teams')
    }
    'Slack' = @{
        Processes     = @('slack')
        Services      = @()
        CleanupTokens = @('Slack')
    }
    'Telegram' = @{
        Processes     = @('Telegram')
        Services      = @()
        CleanupTokens = @('Telegram Desktop','Telegram')
    }
}

foreach ($app in $script:Apps) {
    if ($meta.ContainsKey($app.Key)) {
        $m = $meta[$app.Key]
        if ($m.ContainsKey('Processes'))     { $app.Processes     = $m.Processes }
        if ($m.ContainsKey('Services'))      { $app.Services      = $m.Services }
        if ($m.ContainsKey('CleanupTokens')) { $app.CleanupTokens = $m.CleanupTokens }
    }
}

# Tooltips
foreach ($app in $script:Apps) {
    if ($app.Checkbox -and $app.Description) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = $app.Description
        $tb.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $tb.MaxWidth = 320

        $tooltip = New-Object System.Windows.Controls.ToolTip
        $tooltip.Content = $tb

        $app.Checkbox.ToolTip = $tooltip
    }
}

# Warn if any checkbox missing
$missingControls = $script:Apps | Where-Object { -not $_.Checkbox }
if ($missingControls) {
    Write-Host "WARNING: The following app entries do not have matching controls in XAML:" -ForegroundColor Yellow
    $missingControls | ForEach-Object { Write-Host " - $($_.Name) (Key=$($_.Key))" -ForegroundColor Yellow }
}

# ---------------------------
# Helpers: selection & visibility
# ---------------------------
function Clear-VisibleSelections {
    $script:Apps | Where-Object { $_.Checkbox -and $_.Checkbox.Visibility -eq 'Visible' } |
        ForEach-Object { $_.Checkbox.IsChecked = $false }
}

function Apply-Profile {
    param([string]$ProfileName)

    foreach ($app in $script:Apps) {
        if (-not $app.Checkbox) { continue }

        if ($app.Profiles -contains $ProfileName) {
            $app.Checkbox.Visibility = 'Visible'
        } else {
            $app.Checkbox.Visibility = 'Collapsed'
            $app.Checkbox.IsChecked  = $false
        }
    }

    Write-Log "Profile applied: $ProfileName (visibility updated, selections cleared)."
    Clear-VisibleSelections
}

function Get-SelectedApps {
    $script:Apps | Where-Object {
        $_.Checkbox -and $_.Checkbox.Visibility -eq 'Visible' -and $_.Checkbox.IsChecked -eq $true
    }
}

# ---------------------------
# Backend availability helpers
# ---------------------------
function Ensure-Winget {
    $cmd = Get-Command winget -ErrorAction SilentlyContinue
    if ($cmd) { return $true }

    if (-not $global:WingetPromptShown) {
        $global:WingetPromptShown = $true

        $msg = @"
winget.exe (Windows Package Manager) was not found on this system.

winget is provided by the ""App Installer"" package from the Microsoft Store.

Do you want to open the Microsoft Store page now to install or repair it?
"@

        $result = [System.Windows.MessageBox]::Show(
            $msg,
            "winget not found",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Start-Process "ms-windows-store://pdp/?productid=9NBLGGH4NNS1" -ErrorAction Stop
                Write-Log "Opened Microsoft Store for App Installer (winget)."
            } catch {
                Write-Log "Failed to open Microsoft Store for App Installer: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show(
                    "Could not open Microsoft Store automatically.`r`n`r`nSearch for ""App Installer"" in the Store and install it manually.",
                    "Store launch failed",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
            }
        } else {
            Write-Log "User chose not to open Microsoft Store for winget."
        }
    }

    return $false
}

function Ensure-Choco {
    # Have we already checked/updated this run?
    if ($script:ChocoChecked) {
        return [bool](Get-Command choco -ErrorAction SilentlyContinue)
    }

    $script:ChocoChecked = $true

    $cmd = Get-Command choco -ErrorAction SilentlyContinue

    if ($cmd) {
        try {
            Write-Log "Chocolatey detected. Checking for core updates (choco upgrade chocolatey -y)..."
            $proc = Start-Process -FilePath "choco.exe" -ArgumentList "upgrade chocolatey -y" -NoNewWindow -PassThru -Wait -ErrorAction Stop
            if ($proc.ExitCode -eq 0) {
                Write-Log "Chocolatey core package is present and up to date."
            } else {
                Write-Log "Chocolatey upgrade returned ExitCode=$($proc.ExitCode). Continuing with existing version."
            }
        } catch {
            Write-Log "Chocolatey upgrade check failed: $($_.Exception.Message)"
        }
        return $true
    }

    $msg = @"
Chocolatey (choco.exe) was not found on this system.

Chocolatey is a third-party package manager often used on servers and power-user builds.

Do you want to install Chocolatey automatically now?
(Requires administrator rights and internet access.)
"@

    $result = [System.Windows.MessageBox]::Show(
        $msg,
        "Chocolatey not found",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Log "User declined automatic Chocolatey install."
        return $false
    }

    $installCmd = @"
Set-ExecutionPolicy Bypass -Scope Process -Force;
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072;
iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
"@.Trim()

    try {
        if ($script:IsAdmin) {
            Write-Log "Running Chocolatey install in current elevated session..."
            powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $installCmd
        } else {
            Write-Log "Launching elevated PowerShell window to install Chocolatey..."
            Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-Command",$installCmd
        }
    } catch {
        Write-Log "Failed to start Chocolatey installer: $($_.Exception.Message)"
        return $false
    }

    Start-Sleep -Seconds 5
    $cmd = Get-Command choco -ErrorAction SilentlyContinue
    if ($cmd) {
        Write-Log "Chocolatey installation appears to be present."
        return $true
    } else {
        Write-Log "Chocolatey still not detected after install attempt."
        return $false
    }
}

function Ensure-Backend {
    $backend = Get-Backend
    switch ($backend) {
        'Winget'     { return (Ensure-Winget) }
        'Chocolatey' { return (Ensure-Choco) }
        default      { return $false }
    }
}

# ---------------------------
# Installed-status detection (winget only) – no hard failure if JSON missing
# ---------------------------
function Detect-InstalledApps {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "Skipping installed detection: winget not found."
        return
    }

    try {
        Write-Log "Detecting installed apps via winget list..."
        $jsonText = winget list --source winget --accept-source-agreements --output json 2>$null

        if (-not $jsonText) {
            Write-Log "No output from winget list; skipping detection."
            return
        }

        $trim = $jsonText.TrimStart()
        if (-not ($trim.StartsWith('[') -or $trim.StartsWith('{'))) {
            Write-Log "This winget build does not appear to support JSON output; skipping installed-status detection."
            return
        }

        $data = $jsonText | ConvertFrom-Json

        $ids = @()
        if ($data.PSObject.Properties.Name -contains 'Sources') {
            foreach ($src in $data.Sources) {
                if ($src.Packages) {
                    $ids += $src.Packages.Id
                }
            }
        } elseif ($data.Packages) {
            $ids = $data.Packages.Id
        }

        $hash = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($id in $ids) {
            if ($id) { [void]$hash.Add($id) }
        }

        foreach ($app in $script:Apps) {
            if ($app.WingetId -and $hash.Contains($app.WingetId)) {
                $app.Installed = $true
                if ($app.Checkbox) {
                    $content = $app.Checkbox.Content.ToString()
                    if ($content -notlike '*Installed*') {
                        $app.Checkbox.Content = "$content (Installed)"
                    }
                    $app.Checkbox.FontWeight = 'Bold'
                }
            }
        }

        Write-Log "Installed-status detection complete."
    } catch {
        Write-Log "Installed detection skipped (JSON parse not supported on this winget build)."
    }
}

# Confirm an individual winget registration after successful install
function Test-WingetRegistered {
    param([string]$Id)

    if (-not $Id) { return $false }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return $false }

    try {
        $output = winget list --id "$Id" --source winget --accept-source-agreements 2>$null
        if ($LASTEXITCODE -eq 0 -and $output -match [regex]::Escape($Id)) {
            return $true
        }
    } catch {
        return $false
    }

    return $false
}

# ---------------------------
# Runtime stop/restart + cleanup helpers
# ---------------------------
function Stop-AppRuntime {
    param(
        [Parameter(Mandatory=$true)] [psobject]$App,
        [ValidateSet('update','uninstall')] [string]$Reason
    )

    $runningProcs = @()
    foreach ($p in $App.Processes) {
        try {
            $runningProcs += Get-Process -Name $p -ErrorAction SilentlyContinue
        } catch {}
    }

    $runningSvcs = @()
    foreach ($s in $App.Services) {
        try {
            $svc = Get-Service -Name $s -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -eq 'Running') { $runningSvcs += $svc }
        } catch {}
    }

    if (-not $runningProcs -and -not $runningSvcs) {
        return $null
    }

    $names = @()
    if ($runningProcs) { $names += ("Processes: " + ($runningProcs.Name | Sort-Object -Unique -join ', ')) }
    if ($runningSvcs)  { $names += ("Services: "  + ($runningSvcs.Name  | Sort-Object -Unique -join ', ')) }
    $namesText = $names -join "`r`n"

    if ($Reason -eq 'update') {
        $msg = "The following for $($App.Name) appear to be running:`r`n`r`n$namesText`r`n`r`n" +
               "Stop them so the update can proceed?"
        $res = [System.Windows.MessageBox]::Show(
            $msg,
            "Stop $($App.Name) for update?",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($res -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-Log "User chose not to stop $($App.Name) before update."
            return $null
        }
    } else {
        Write-Log "Stopping $($App.Name) runtime for uninstall."
    }

    $stoppedProcs = @()
    foreach ($p in $runningProcs) {
        try {
            Write-Log "Stopping process $($p.Name) (Id=$($p.Id)) for $Reason..."
            Stop-Process -Id $p.Id -Force -ErrorAction Stop
            $stoppedProcs += $p.Name
        } catch {
            Write-Log "Failed to stop process $($p.Name): $($_.Exception.Message)"
        }
    }

    $stoppedSvcs = @()
    foreach ($svc in $runningSvcs) {
        try {
            Write-Log "Stopping service $($svc.Name) for $Reason..."
            Stop-Service -Name $svc.Name -Force -ErrorAction Stop
            $stoppedSvcs += $svc.Name
        } catch {
            Write-Log "Failed to stop service $($svc.Name): $($_.Exception.Message)"
        }
    }

    if ($Reason -eq 'update' -and ($stoppedProcs -or $stoppedSvcs)) {
        return [pscustomobject]@{
            App            = $App
            Processes      = $stoppedProcs
            Services       = $stoppedSvcs
        }
    }

    return $null
}

function Restart-AppRuntime {
    param(
        [Parameter(Mandatory=$true)] [psobject]$StoppedInfo
    )

    $app = $StoppedInfo.App

    foreach ($svcName in $StoppedInfo.Services) {
        try {
            Write-Log "Restarting service $svcName for $($app.Name)..."
            Start-Service -Name $svcName -ErrorAction Stop
        } catch {
            Write-Log "Failed to restart service ${svcName}: $($_.Exception.Message)"
        }
    }

    foreach ($procName in $StoppedInfo.Processes | Sort-Object -Unique) {
        try {
            Write-Log "Relaunching process $procName for $($app.Name)..."
            Start-Process -FilePath "$procName.exe" -ErrorAction Stop
        } catch {
            try {
                Start-Process -FilePath $procName -ErrorAction Stop
            } catch {
                Write-Log "Failed to relaunch process ${procName}: $($_.Exception.Message)"
            }
        }
    }
}

function Cleanup-AppAfterUninstall {
    param(
        [Parameter(Mandatory=$true)] [psobject]$App
    )

    if (-not $App.CleanupTokens -or $App.CleanupTokens.Count -eq 0) {
        return
    }

    if (-not $script:IsAdmin) {
        Write-Log "Deep cleanup for $($App.Name) may be limited (not running as administrator)."
    }

    $dirRoots = @(
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)},
        $env:ProgramData,
        $env:LOCALAPPDATA,
        $env:APPDATA
    ) | Where-Object { $_ }

    $regRoots = @(
        'HKCU:\Software',
        'HKLM:\Software',
        'HKLM:\Software\WOW6432Node'
    )

    foreach ($token in $App.CleanupTokens) {
        foreach ($root in $dirRoots) {
            $path = Join-Path $root $token
            if (Test-Path $path) {
                try {
                    Write-Log "Removing folder $path for $($App.Name)..."
                    Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                } catch {
                    Write-Log "Failed to remove folder ${path}: $($_.Exception.Message)"
                }
            }
        }

        foreach ($root in $regRoots) {
            $regPath = Join-Path $root $token
            if (Test-Path $regPath) {
                try {
                    Write-Log "Removing registry key $regPath for $($App.Name)..."
                    Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
                } catch {
                    Write-Log "Failed to remove registry key ${regPath}: $($_.Exception.Message)"
                }
            }
        }
    }
}

# ---------------------------
# Invoke actions via backend
# ---------------------------
function Invoke-AppAction {
    param(
        [ValidateSet('install','update','uninstall')]
        [string]$Action,
        [object[]]$SelectedApps
    )

    if (-not (Ensure-Backend)) {
        Write-Log "Backend unavailable. Cannot perform $Action."
        return
    }

    if (-not $SelectedApps -or $SelectedApps.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Select at least one tool first.",
            "No selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $backend     = Get-Backend
    $names       = $SelectedApps.Name -join "`r`n - "
    $actionLabel = switch ($Action) {
        'install'   { 'Install' }
        'update'    { 'Update' }
        'uninstall' { 'Uninstall' }
    }

    $result = [System.Windows.MessageBox]::Show(
        "Using backend: $backend`r`n`r`nYou are about to $($actionLabel.ToLower()) the following:`r`n`r`n - $names`r`n`r`nContinue?",
        "Confirm $actionLabel",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Log "$actionLabel cancelled by user."
        return
    }

    Show-Busy "$actionLabel in progress..."
    Write-Log "$actionLabel started for $($SelectedApps.Count) app(s) using backend: $backend."

    $success = @()
    $failed  = @()

    $total = $SelectedApps.Count
    $index = 0

    foreach ($app in $SelectedApps) {
        $index++
if ($txtBusy) {
    $txtBusy.Text = "{0} {1} of {2}: {3}" -f $actionLabel, $index, $total, $app.Name
    $Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

        if (-not $app.Checkbox) { continue }

        $verb = $actionLabel
        $id   = $null
        $file = $null
        $args = $null

        # Pre-actions for update/uninstall
        $stoppedInfo = $null
        if ($Action -eq 'update' -or $Action -eq 'uninstall') {
            $reason = $Action
            $stoppedInfo = Stop-AppRuntime -App $app -Reason $reason
        }

        switch ($backend) {
            'Winget' {
                $id = $app.WingetId
                if (-not $id) {
                    Write-Log "No winget ID for $($app.Name); skipping."
                    $failed += $app.Name
                    continue
                }
                $file = "winget.exe"
                $args = switch ($Action) {
                    'install'   { "install --id `"$id`" --silent --accept-package-agreements --accept-source-agreements" }
                    'update'    { "upgrade --id `"$id`" --silent --accept-package-agreements --accept-source-agreements" }
                    'uninstall' { "uninstall --id `"$id`" --silent" }
                }
            }

            'Chocolatey' {
                $id = $app.ChocoId
                if (-not $id) {
                    Write-Log "No Chocolatey ID for $($app.Name); skipping for choco backend."
                    $failed += $app.Name
                    continue
                }
                $file = "choco.exe"
                $args = switch ($Action) {
                    'install'   { "install $id -y" }
                    'update'    { "upgrade $id -y" }
                    'uninstall' { "uninstall $id -y" }
                }
            }

            default {
                Write-Log "Unknown backend: $backend."
                $failed += $app.Name
                continue
            }
        }

        Write-Log ("{0}: {1} [{2}] via {3}" -f $verb, $app.Name, $id, $backend)

        $backendOk = $false
        try {
            $proc = Start-Process -FilePath $file -ArgumentList $args -NoNewWindow -PassThru -Wait -ErrorAction Stop

            switch ($proc.ExitCode) {
                0 {
                    Write-Log "$verb OK: $($app.Name)"
                    $backendOk = $true
                    $success += $app.Name

                    if ($Action -eq 'install' -and $backend -eq 'Winget') {
                        if (-not (Test-WingetRegistered -Id $id)) {
                            Write-Log "NOTE: $($app.Name) installed via winget, but winget list does not show it yet. Shortcuts/registration may rely on the vendor installer."
                        }
                    }
                }

                -1978335212 {   # 0x8A150014 APPINSTALLER_CLI_ERROR_NO_APPLICATIONS_FOUND
                    Write-Log "$verb FAILED: $($app.Name) is not registered with $backend (no matching installed package found by $backend)."
                    Write-Log "Tip: Uninstall it once manually, then reinstall using $backend so future updates/uninstalls can be managed here."
                    $failed += $app.Name
                }

                default {
                    Write-Log "$verb FAILED (ExitCode=$($proc.ExitCode)): $($app.Name)"
                    $failed += $app.Name
                }
            }
        } catch {
            Write-Log "$verb ERROR for $($app.Name): $($_.Exception.Message)"
            $failed += $app.Name
        }

        # Post-actions
        if ($Action -eq 'update' -and $backendOk -and $stoppedInfo) {
            Restart-AppRuntime -StoppedInfo $stoppedInfo
        } elseif ($Action -eq 'uninstall' -and $backendOk) {
            Cleanup-AppAfterUninstall -App $app
        }
    }

    Write-Log "$actionLabel completed."
    Hide-Busy

    # Summary popup
    $msgLines = @()
    if ($success.Count -gt 0) {
        $msgLines += "Succeeded ($($success.Count)):"
        $msgLines += " - " + ($success -join "`r`n - ")
    }
    if ($failed.Count -gt 0) {
        if ($msgLines.Count -gt 0) { $msgLines += "" }
        $msgLines += "Failed or skipped ($($failed.Count)):"
        $msgLines += " - " + ($failed -join "`r`n - ")
    }
    if ($msgLines.Count -eq 0) {
        $msgLines = @("$actionLabel finished. No changes were made.")
    }

    [System.Windows.MessageBox]::Show(
        ($msgLines -join "`r`n"),
        "$actionLabel summary",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
}

# ---------------------------
# JSON Profile save/load
# ---------------------------
function Save-CurrentProfileToFile {
    $profileName = Get-ProfileName
    $backend     = Get-Backend
    $selected    = @(Get-SelectedApps)

    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Select at least one app before saving a profile.",
            "Nothing selected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $keys = $selected.Key
    $obj  = [pscustomobject]@{
        ProfileName = $profileName
        Backend     = $backend
        Apps        = $keys
    }

    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "JSON Files|*.json|All Files|*.*"
    $safeName = ($profileName -replace '\s','_')
    if (-not [string]::IsNullOrWhiteSpace($safeName)) {
        $dlg.FileName = "$safeName.json"
    }

    if ($dlg.ShowDialog()) {
        $json = $obj | ConvertTo-Json -Depth 3
        Set-Content -Path $dlg.FileName -Value $json -Encoding UTF8
        Write-Log "Profile saved to $($dlg.FileName)."
    }
}

function Set-ComboSelectionByContent {
    param(
        [System.Windows.Controls.ComboBox]$Combo,
        [string]$Content
    )
    if (-not $Combo) { return }
    for ($i = 0; $i -lt $Combo.Items.Count; $i++) {
        $item = $Combo.Items[$i]
        if ($item.Content -and ($item.Content.ToString() -eq $Content)) {
            $Combo.SelectedIndex = $i
            return
        }
    }
}

function Load-ProfileFromFile {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "JSON Files|*.json|All Files|*.*"

    if (-not $dlg.ShowDialog()) { return }

    try {
        $json = Get-Content -Path $dlg.FileName -Raw
        $profile = $json | ConvertFrom-Json

        if ($profile.ProfileName) {
            Set-ComboSelectionByContent -Combo $cmbProfile -Content $profile.ProfileName
        }
        if ($profile.Backend) {
            Set-ComboSelectionByContent -Combo $cmbBackend -Content $profile.Backend
        }

        $profileName = Get-ProfileName
        Apply-Profile -ProfileName $profileName

        if ($profile.Apps) {
            foreach ($app in $script:Apps) {
                if ($app.Checkbox -and ($profile.Apps -contains $app.Key)) {
                    if ($app.Checkbox.Visibility -eq 'Visible') {
                        $app.Checkbox.IsChecked = $true
                    }
                }
            }
        }

        Write-Log "Profile loaded from $($dlg.FileName)."
    } catch {
        Write-Log "Failed to load profile: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to load profile.`r`n`r`n$($_.Exception.Message)",
            "Profile load error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
}

# ---------------------------
# Profile handling
# ---------------------------
$cmbProfile.Add_SelectionChanged({
    $profileName = Get-ProfileName
    Apply-Profile -ProfileName $profileName
})

# ---------------------------
# Button handlers
# ---------------------------
$btnSelectAll.Add_Click({
    $script:Apps | Where-Object { $_.Checkbox -and $_.Checkbox.Visibility -eq 'Visible' } |
        ForEach-Object { $_.Checkbox.IsChecked = $true }
})

$btnClearAll.Add_Click({
    Clear-VisibleSelections
})

$btnInstall.Add_Click({
    $sel = @(Get-SelectedApps)
    Invoke-AppAction -Action 'install' -SelectedApps $sel
})

$btnUpdate.Add_Click({
    $sel = @(Get-SelectedApps)
    Invoke-AppAction -Action 'update' -SelectedApps $sel
})

$btnUninstall.Add_Click({
    $sel = @(Get-SelectedApps)
    Invoke-AppAction -Action 'uninstall' -SelectedApps $sel
})

$btnClose.Add_Click({
    $Window.Close()
})

$btnSaveProfile.Add_Click({
    Save-CurrentProfileToFile
})

$btnLoadProfile.Add_Click({
    Load-ProfileFromFile
})

# ---------------------------
# Initial state
# ---------------------------
$initialProfileName = Get-ProfileName
Apply-Profile -ProfileName $initialProfileName
Write-Log "MicroDeploy Toolbox UI loaded. Profile: $initialProfileName. Backend: $(Get-Backend)."
Write-Log "All checkboxes are unchecked by default. Select apps, or use Select All / Clear, or load a JSON profile."

# Run installed detection AFTER window is rendered so we can show a loading overlay
$Window.Add_ContentRendered({
    Show-Busy "Initializing (scanning installed apps)..."
    Detect-InstalledApps
    Hide-Busy
})

# ---------------------------
# Run window
# ---------------------------
[void]$Window.ShowDialog()
