#Requires -Version 5.1

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$Script:MainXaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Print Deployment Maker"
    Width="840" Height="760"
    MinWidth="700" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    FontFamily="Segoe UI" FontSize="13"
    Background="{DynamicResource BrushWinBg}">

    <Window.Resources>
        <!-- Theme brushes (light-mode defaults) -->
        <SolidColorBrush x:Key="BrushWinBg"          Color="#F0F0F0"/>
        <SolidColorBrush x:Key="BrushPanelBg"         Color="#FFFFFF"/>
        <SolidColorBrush x:Key="BrushPanelBorder"     Color="#D0D0D0"/>
        <SolidColorBrush x:Key="BrushBarBg"           Color="#E8E8E8"/>
        <SolidColorBrush x:Key="BrushBorder"          Color="#CCCCCC"/>
        <SolidColorBrush x:Key="BrushControlBg"       Color="#FFFFFF"/>
        <SolidColorBrush x:Key="BrushTextHeader"      Color="#222222"/>
        <SolidColorBrush x:Key="BrushTextBody"        Color="#444444"/>
        <SolidColorBrush x:Key="BrushTextMuted"       Color="#555555"/>
        <SolidColorBrush x:Key="BrushTextFaint"       Color="#888888"/>
        <SolidColorBrush x:Key="BrushListHover"       Color="#EBF4FF"/>
        <SolidColorBrush x:Key="BrushListSelected"    Color="#CCE4FF"/>
        <SolidColorBrush x:Key="BrushListSelectedFg"  Color="#000000"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}"                      Color="#CCE4FF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}"                  Color="#000000"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightBrushKey}"     Color="#E0EEFF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightTextBrushKey}" Color="#000000"/>

        <!-- Flat button style: opacity-only hover/press so any background colour works -->
        <Style x:Key="FlatBtn" TargetType="Button">
            <Setter Property="Foreground"      Value="White"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="3"
                                Padding="{TemplateBinding Padding}"
                                SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.82"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.65"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.38"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Implicit styles -->
        <Style TargetType="TextBox">
            <Setter Property="Background"       Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"       Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="CaretBrush"       Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush"      Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"          Value="4,3"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background"  Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"     Value="4,3"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground"   Value="{DynamicResource BrushTextHeader}"/>
            <Setter Property="BorderBrush"  Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"      Value="6"/>
        </Style>
        <Style TargetType="ListView">
            <Setter Property="Background"  Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextBody}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource BrushListHover}"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource BrushListSelected}"/>
                    <Setter Property="Foreground" Value="{DynamicResource BrushListSelectedFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Background"      Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"      Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush"     Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"         Value="10,4"/>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background"  Value="{DynamicResource BrushBarBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextMuted}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"     Value="6,3"/>
            <Setter Property="FontSize"    Value="11"/>
        </Style>
    </Window.Resources>

    <DockPanel>

        <!-- ── Header bar ── -->
        <Border DockPanel.Dock="Top" Background="#0078D4" Padding="12,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Logo placeholder (96×96 spec, visually contained in header) -->
                <Border Grid.Column="0" Width="46" Height="46" CornerRadius="6"
                        Background="#1A5FA8" VerticalAlignment="Center">
                    <TextBlock Text="PDM" Foreground="White" FontWeight="Bold" FontSize="13"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>

                <StackPanel Grid.Column="1" Orientation="Horizontal"
                            VerticalAlignment="Center" Margin="12,0,0,0">
                    <TextBlock x:Name="TitleText" Text="Print Deployment Maker"
                               Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                    <TextBlock x:Name="VersionText" Text=""
                               Foreground="#AADCFF" FontSize="12"
                               Margin="10,0,0,0" VerticalAlignment="Center"/>
                </StackPanel>

                <Button x:Name="ThemeBtn" Grid.Column="2" Content="☾  Dark"
                        Style="{StaticResource FlatBtn}"
                        Background="#1A5FA8" Foreground="White" Padding="10,5"/>
            </Grid>
        </Border>

        <!-- ── Log pane (bottom) ── -->
        <Border DockPanel.Dock="Bottom" BorderThickness="0,1,0,0"
                BorderBrush="{DynamicResource BrushBorder}"
                Background="{DynamicResource BrushBarBg}">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="130"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Text="Log" FontSize="11" FontWeight="SemiBold"
                           Foreground="{DynamicResource BrushTextMuted}" Margin="8,4,0,2"/>
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Disabled">
                    <TextBox x:Name="LogBox" IsReadOnly="True" TextWrapping="Wrap"
                             FontFamily="Consolas" FontSize="11" BorderThickness="0"
                             Background="{DynamicResource BrushBarBg}"
                             Foreground="{DynamicResource BrushTextBody}"
                             Padding="8,4"/>
                </ScrollViewer>
            </Grid>
        </Border>

        <!-- ── Main form ── -->
        <ScrollViewer VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled">
            <StackPanel Margin="12,10,12,10">

                <!-- Driver section -->
                <GroupBox Header="Driver" Margin="0,0,0,10">
                    <Grid Margin="4,6,4,4">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="8"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="4"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="10"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="4"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="4"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <!-- .inf browse -->
                        <TextBox x:Name="InfPathBox" Grid.Row="0" Grid.Column="0"
                                 IsReadOnly="True" Text="No .inf file selected"
                                 Foreground="{DynamicResource BrushTextFaint}"/>
                        <Button x:Name="BrowseInfBtn" Grid.Row="0" Grid.Column="1"
                                Content="Browse .inf…" Margin="8,0,0,0"/>

                        <!-- Driver model -->
                        <TextBlock Grid.Row="2" Grid.ColumnSpan="2"
                                   Text="Driver Model" FontSize="11"
                                   Foreground="{DynamicResource BrushTextMuted}"/>
                        <ComboBox x:Name="DriverModelCombo" Grid.Row="4" Grid.ColumnSpan="2"/>

                        <!-- Separator -->
                        <Separator Grid.Row="5" Grid.ColumnSpan="2"
                                   Background="{DynamicResource BrushBorder}" Margin="0,2"/>

                        <!-- Manual driver name for Queue Only -->
                        <TextBlock Grid.Row="6" Grid.ColumnSpan="2"
                                   Text="Manual Driver Name  (Print Queue Only)"
                                   FontSize="11" Foreground="{DynamicResource BrushTextMuted}"/>
                        <TextBox x:Name="ManualDriverBox" Grid.Row="8" Grid.ColumnSpan="2"/>
                        <TextBlock Grid.Row="10" Grid.ColumnSpan="2" FontSize="10"
                                   FontStyle="Italic" Foreground="{DynamicResource BrushTextFaint}"
                                   Text="Type the exact driver name already installed — used only by 'Package Print Queue Only'"/>
                    </Grid>
                </GroupBox>

                <!-- Deployment name -->
                <GroupBox Header="Deployment" Margin="0,0,0,10">
                    <Grid Margin="4,6,4,4">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="4"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Deployment Name" FontSize="11"
                                   Foreground="{DynamicResource BrushTextMuted}"/>
                        <TextBox x:Name="DeploymentNameBox" Grid.Row="2"/>
                    </Grid>
                </GroupBox>

                <!-- Print queues -->
                <GroupBox Header="Print Queues" Margin="0,0,0,10">
                    <Grid Margin="4,6,4,4">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="6"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="6"/>
                            <RowDefinition Height="130"/>
                            <RowDefinition Height="6"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <!-- Column headers hint -->
                        <Grid Grid.Row="0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="8"/>
                                <ColumnDefinition Width="160"/>
                                <ColumnDefinition Width="8"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="Printer Name" FontSize="11"
                                       Foreground="{DynamicResource BrushTextMuted}"/>
                            <TextBlock Grid.Column="2" Text="IP Address" FontSize="11"
                                       Foreground="{DynamicResource BrushTextMuted}"/>
                        </Grid>

                        <!-- Add row -->
                        <Grid Grid.Row="2">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="8"/>
                                <ColumnDefinition Width="160"/>
                                <ColumnDefinition Width="8"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="NewPrinterNameBox" Grid.Column="0"/>
                            <TextBox x:Name="NewPrinterIPBox"   Grid.Column="2"/>
                            <Button  x:Name="AddQueueBtn"       Grid.Column="4"
                                     Content="Add Queue" Padding="10,4"/>
                        </Grid>

                        <!-- Queue list -->
                        <ListView x:Name="QueueListView" Grid.Row="4">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Printer Name" Width="380"
                                                    DisplayMemberBinding="{Binding Content}"/>
                                    <GridViewColumn Header="IP Address" Width="160"
                                                    DisplayMemberBinding="{Binding Tag}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>

                        <Button x:Name="RemoveQueueBtn" Grid.Row="6"
                                Content="Remove Selected" HorizontalAlignment="Right"/>
                    </Grid>
                </GroupBox>

                <!-- Actions -->
                <GroupBox Header="Actions" Margin="0,0,0,10">
                    <UniformGrid Margin="4,6,4,4" Rows="2" Columns="2">
                        <Button x:Name="CreateBtn"
                                Content="Create" Margin="0,0,4,4" Padding="12,10"
                                Style="{StaticResource FlatBtn}" Background="#0078D4"/>
                        <Button x:Name="CreatePackageBtn"
                                Content="Create and Package" Margin="4,0,0,4" Padding="12,10"
                                Style="{StaticResource FlatBtn}" Background="#107C10"/>
                        <Button x:Name="DriverOnlyBtn"
                                Content="Package Driver Only" Margin="0,4,4,0" Padding="12,10"
                                Style="{StaticResource FlatBtn}" Background="#5C2D91"/>
                        <Button x:Name="QueueOnlyBtn"
                                Content="Package Print Queue Only" Margin="4,4,0,0" Padding="12,10"
                                Style="{StaticResource FlatBtn}" Background="#D83B01"/>
                    </UniformGrid>
                </GroupBox>

            </StackPanel>
        </ScrollViewer>

    </DockPanel>
</Window>
'@

# ── Script-scope state ───────────────────────────────────────────────────────

$Script:IsDarkMode       = $false
$Script:InfPath          = ''
$Script:InfSourceDir     = ''
$Script:DriverFolderName = ''
$Script:InfFileName      = ''
$Script:ScriptRoot       = ''
$Script:UI               = @{}

# ── Theme ────────────────────────────────────────────────────────────────────

function Set-Theme {
    param([bool]$Dark)
    $palette = if ($Dark) {
        @{
            BrushWinBg          = '#1E1E1E'
            BrushPanelBg        = '#252526'
            BrushPanelBorder    = '#3C3C3C'
            BrushBarBg          = '#252526'
            BrushBorder         = '#555555'
            BrushControlBg      = '#3C3C3C'
            BrushTextHeader     = '#E8E8E8'
            BrushTextBody       = '#CCCCCC'
            BrushTextMuted      = '#AAAAAA'
            BrushTextFaint      = '#777777'
            BrushListHover      = '#2D4B6E'
            BrushListSelected   = '#094771'
            BrushListSelectedFg = '#FFFFFF'
            HighlightColor        = '#094771'
            HighlightText         = '#FFFFFF'
            InactiveHighlight     = '#1C3A5C'
            InactiveHighlightText = '#FFFFFF'
        }
    } else {
        @{
            BrushWinBg          = '#F0F0F0'
            BrushPanelBg        = '#FFFFFF'
            BrushPanelBorder    = '#D0D0D0'
            BrushBarBg          = '#E8E8E8'
            BrushBorder         = '#CCCCCC'
            BrushControlBg      = '#FFFFFF'
            BrushTextHeader     = '#222222'
            BrushTextBody       = '#444444'
            BrushTextMuted      = '#555555'
            BrushTextFaint      = '#888888'
            BrushListHover      = '#EBF4FF'
            BrushListSelected   = '#CCE4FF'
            BrushListSelectedFg = '#000000'
            HighlightColor        = '#CCE4FF'
            HighlightText         = '#000000'
            InactiveHighlight     = '#E0EEFF'
            InactiveHighlightText = '#000000'
        }
    }

    $r    = $Script:UI.Window.Resources
    $conv = [System.Windows.Media.ColorConverter]
    foreach ($key in ($palette.Keys | Where-Object { $_ -like 'Brush*' })) {
        $r[$key] = [System.Windows.Media.SolidColorBrush]::new($conv::ConvertFromString($palette[$key]))
    }
    $r[[System.Windows.SystemColors]::HighlightBrushKey]                      = [System.Windows.Media.SolidColorBrush]::new($conv::ConvertFromString($palette.HighlightColor))
    $r[[System.Windows.SystemColors]::HighlightTextBrushKey]                  = [System.Windows.Media.SolidColorBrush]::new($conv::ConvertFromString($palette.HighlightText))
    $r[[System.Windows.SystemColors]::InactiveSelectionHighlightBrushKey]     = [System.Windows.Media.SolidColorBrush]::new($conv::ConvertFromString($palette.InactiveHighlight))
    $r[[System.Windows.SystemColors]::InactiveSelectionHighlightTextBrushKey] = [System.Windows.Media.SolidColorBrush]::new($conv::ConvertFromString($palette.InactiveHighlightText))

    $Script:UI.ThemeBtn.Content = if ($Dark) { '☀  Light' } else { '☾  Dark' }
    $Script:IsDarkMode = $Dark
}

# ── Logging ──────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    $Script:UI.LogBox.Dispatcher.Invoke([action]{
        $Script:UI.LogBox.AppendText("[$ts] $Message`n")
        $Script:UI.LogBox.ScrollToEnd()
    })
}

# ── INF parser ───────────────────────────────────────────────────────────────

function Parse-InfDriverModels {
    param([string]$InfPath)

    try {
        $lines = Get-Content -Path $InfPath -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Log "ERROR reading .inf file: $_"
        return @()
    }

    # Build [Strings] lookup table
    $strings  = @{}
    $inTarget = $false
    foreach ($line in $lines) {
        $clean = ($line -replace ';.*$', '').Trim()
        if ($clean -match '^\[(.+)\]') {
            $inTarget = $Matches[1] -ieq 'Strings'
            continue
        }
        if ($inTarget -and $clean -match '^(\w+)\s*=\s*"(.+)"') {
            $strings[$Matches[1].ToLower()] = $Matches[2]
        }
    }

    # Find manufacturer section names from [Manufacturer]
    $mfrSections = [System.Collections.Generic.List[string]]::new()
    $inMfr = $false
    foreach ($line in $lines) {
        $clean = ($line -replace ';.*$', '').Trim()
        if ($clean -match '^\[(.+)\]') {
            $inMfr = $Matches[1] -ieq 'Manufacturer'
            continue
        }
        if ($inMfr -and $clean -match '=\s*(.+)') {
            foreach ($part in ($Matches[1] -split ',')) {
                $p = $part.Trim()
                # Skip platform decorators (NTamd64, NTx86, etc.) and empty
                if ($p -and $p -notmatch '^NT') { $mfrSections.Add($p) }
            }
        }
    }

    # Build list of target section name variants (base + NTamd64 + NTamd64.10.0)
    $targets = [System.Collections.Generic.List[string]]::new()
    foreach ($s in ($mfrSections | Select-Object -Unique)) {
        $targets.Add($s)
        $targets.Add("$s.NTamd64")
        $targets.Add("$s.NTamd64.10.0")
        $targets.Add("$s.NTamd64.10.0.0")
    }

    # Collect model names from target sections
    $models   = [System.Collections.Generic.List[string]]::new()
    $inTarget = $false
    foreach ($line in $lines) {
        $clean = ($line -replace ';.*$', '').Trim()
        if ($clean -match '^\[(.+)\]') {
            $sName    = $Matches[1].Trim()
            $inTarget = ($null -ne ($targets | Where-Object { $_ -ieq $sName }))
            continue
        }
        if (-not $inTarget -or -not ($clean -match '^(.+?)\s*=')) { continue }

        $token = $Matches[1].Trim()
        if ($token -match '^"(.+)"$') {
            $models.Add($Matches[1])
        } elseif ($token -match '^%(.+)%$') {
            $key = $Matches[1].ToLower()
            if ($strings.ContainsKey($key) -and $strings[$key]) {
                $models.Add($strings[$key])
            }
        }
    }

    return @($models | Where-Object { $_ } | Sort-Object -Unique)
}

# ── Validation ───────────────────────────────────────────────────────────────

function Test-DeploymentName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        Write-Log 'ERROR: Deployment Name is required.'
        return $false
    }
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($Name.IndexOf($ch) -ge 0) {
            Write-Log "ERROR: Deployment Name contains an invalid character."
            return $false
        }
    }
    return $true
}

function Test-ForFullOrDriverOnly {
    if (-not $Script:InfPath) {
        Write-Log 'ERROR: Browse for a .inf file first.'
        return $false
    }
    if ($Script:UI.DriverModelCombo.SelectedIndex -lt 0) {
        Write-Log 'ERROR: Select a driver model from the dropdown.'
        return $false
    }
    if (-not (Test-DeploymentName $Script:UI.DeploymentNameBox.Text.Trim())) { return $false }
    return $true
}

function Test-ForQueuesPresent {
    if ($Script:UI.QueueListView.Items.Count -eq 0) {
        Write-Log 'ERROR: Add at least one print queue.'
        return $false
    }
    return $true
}

function Test-ForQueueOnly {
    if ([string]::IsNullOrWhiteSpace($Script:UI.ManualDriverBox.Text)) {
        Write-Log 'ERROR: Manual Driver Name is required for Print Queue Only.'
        return $false
    }
    if (-not (Test-DeploymentName $Script:UI.DeploymentNameBox.Text.Trim())) { return $false }
    if (-not (Test-ForQueuesPresent)) { return $false }
    return $true
}

# ── Script generation ────────────────────────────────────────────────────────

function ConvertTo-PrinterArrayBlock {
    param([System.Windows.Controls.ListView]$ListView)
    $lines = foreach ($item in $ListView.Items) {
        $n = $item.Content -replace "'", "''"
        $i = $item.Tag     -replace "'", "''"
        "    @{ Name = '$n'; IP = '$i' }"
    }
    return $lines -join "`n"
}

function ConvertTo-NamesBlock {
    param([System.Windows.Controls.ListView]$ListView)
    $quoted = foreach ($item in $ListView.Items) {
        $n = $item.Content -replace "'", "''"
        "'$n'"
    }
    return $quoted -join ', '
}

function New-FullDeployScript {
    param(
        [string]$DriverName,
        [string]$DriverFolder,
        [string]$InfFileName,
        [string]$PrintersBlock
    )
    $template = @'
param([string]$Action = 'Install')

$DriverName     = '__DRIVER_NAME__'
$DriverFolder   = '__DRIVER_FOLDER__'
$InfFileName    = '__INF_FILE__'
$DriverStorePath = "C:\ProgramData\AutoPilotConfig\Printers\$DriverFolder"
$Printers = @(
__PRINTERS_BLOCK__
)

if ($Action -eq 'Install') {
    New-Item -ItemType Directory -Path 'C:\ProgramData\AutoPilotConfig\Printers' -Force | Out-Null
    New-Item -ItemType Directory -Path $DriverStorePath -Force | Out-Null
    Copy-Item -Path "$PSScriptRoot\$DriverFolder\*" -Destination $DriverStorePath -Recurse -Force
    & cscript 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs' `
        -a -m $DriverName -h "$DriverStorePath\" -i "$DriverStorePath\$InfFileName"
    foreach ($p in $Printers) {
        $portName = "TCPPort:$($p.IP)"
        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $p.IP
        }
        if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
            Add-Printer -Name $p.Name -PortName $portName -DriverName $DriverName
        } else {
            Write-Warning "Driver '$DriverName' not found — '$($p.Name)' not added"
        }
    }
} elseif ($Action -eq 'Uninstall') {
    foreach ($p in $Printers) {
        Get-Printer -Name $p.Name -ErrorAction SilentlyContinue | Remove-Printer
    }
}
'@
    $dn = $DriverName   -replace "'", "''"
    $df = $DriverFolder -replace "'", "''"
    $inf = $InfFileName  -replace "'", "''"
    return $template `
        -replace '__DRIVER_NAME__',    $dn `
        -replace '__DRIVER_FOLDER__',  $df `
        -replace '__INF_FILE__',       $inf `
        -replace '__PRINTERS_BLOCK__', $PrintersBlock
}

function New-DriverOnlyDeployScript {
    param(
        [string]$DriverName,
        [string]$DriverFolder,
        [string]$InfFileName
    )
    $template = @'
param([string]$Action = 'Install')

$DriverName     = '__DRIVER_NAME__'
$DriverFolder   = '__DRIVER_FOLDER__'
$InfFileName    = '__INF_FILE__'
$DriverStorePath = "C:\ProgramData\AutoPilotConfig\Printers\$DriverFolder"

if ($Action -eq 'Install') {
    New-Item -ItemType Directory -Path 'C:\ProgramData\AutoPilotConfig\Printers' -Force | Out-Null
    New-Item -ItemType Directory -Path $DriverStorePath -Force | Out-Null
    Copy-Item -Path "$PSScriptRoot\$DriverFolder\*" -Destination $DriverStorePath -Recurse -Force
    & cscript 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs' `
        -a -m $DriverName -h "$DriverStorePath\" -i "$DriverStorePath\$InfFileName"
} elseif ($Action -eq 'Uninstall') {
    & cscript 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs' `
        -d -m $DriverName
}
'@
    $dn  = $DriverName   -replace "'", "''"
    $df  = $DriverFolder -replace "'", "''"
    $inf = $InfFileName  -replace "'", "''"
    return $template `
        -replace '__DRIVER_NAME__',   $dn `
        -replace '__DRIVER_FOLDER__', $df `
        -replace '__INF_FILE__',      $inf
}

function New-QueueOnlyDeployScript {
    param(
        [string]$DriverName,
        [string]$PrintersBlock
    )
    $template = @'
param([string]$Action = 'Install')

$DriverName = '__DRIVER_NAME__'
$Printers = @(
__PRINTERS_BLOCK__
)

if ($Action -eq 'Install') {
    foreach ($p in $Printers) {
        $portName = "TCPPort:$($p.IP)"
        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $p.IP
        }
        if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
            Add-Printer -Name $p.Name -PortName $portName -DriverName $DriverName
        } else {
            Write-Warning "Driver '$DriverName' not found — '$($p.Name)' not added"
        }
    }
} elseif ($Action -eq 'Uninstall') {
    foreach ($p in $Printers) {
        Get-Printer -Name $p.Name -ErrorAction SilentlyContinue | Remove-Printer
    }
}
'@
    $dn = $DriverName -replace "'", "''"
    return $template `
        -replace '__DRIVER_NAME__',    $dn `
        -replace '__PRINTERS_BLOCK__', $PrintersBlock
}

function New-PrinterDetectScript {
    param([string]$NamesBlock)
    $template = @'
$printerNames = @(__NAMES_BLOCK__)
$missing = $printerNames | Where-Object { -not (Get-Printer -Name $_ -ErrorAction SilentlyContinue) }
if ($missing.Count -eq 0) { exit 0 } else { exit 1 }
'@
    return $template -replace '__NAMES_BLOCK__', $NamesBlock
}

function New-DriverDetectScript {
    param([string]$DriverName)
    $dn = $DriverName -replace "'", "''"
    $template = @'
$driverName = '__DRIVER_NAME__'
if (Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }
'@
    return $template -replace '__DRIVER_NAME__', $dn
}

function Invoke-Package {
    param([string]$PackageFolder)
    $util = Join-Path $Script:ScriptRoot 'IntuneWinAppUtil.exe'
    if (-not (Test-Path $util)) {
        Write-Log "ERROR: IntuneWinAppUtil.exe not found at: $util"
        return $false
    }
    Write-Log 'Packaging with IntuneWinAppUtil...'
    $out = & $util -c $PackageFolder -s 'deploy.ps1' -o $PackageFolder -q 2>&1
    foreach ($line in $out) { if ("$line".Trim()) { Write-Log "  $line" } }
    $pkg = Get-ChildItem -Path $PackageFolder -Filter '*.intunewin' -ErrorAction SilentlyContinue |
           Select-Object -First 1
    if ($pkg) { Write-Log "Package created: $($pkg.FullName)" }
    return $true
}

function Write-IntuneCmdHint {
    Write-Log '--- Intune commands ---'
    Write-Log 'Install:   powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Install'
    Write-Log 'Uninstall: powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Uninstall'
}

# ── Main window ───────────────────────────────────────────────────────────────

function Show-MainWindow {
    param([string]$AppVersion = '')

    $Script:ScriptRoot = $PSScriptRoot

    $reader = [System.Xml.XmlReader]::Create(
        [System.IO.StringReader]::new($Script:MainXaml)
    )
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $Script:UI = @{
        Window           = $window
        LogBox           = $window.FindName('LogBox')
        InfPathBox       = $window.FindName('InfPathBox')
        BrowseInfBtn     = $window.FindName('BrowseInfBtn')
        DriverModelCombo = $window.FindName('DriverModelCombo')
        ManualDriverBox  = $window.FindName('ManualDriverBox')
        DeploymentNameBox= $window.FindName('DeploymentNameBox')
        NewPrinterNameBox= $window.FindName('NewPrinterNameBox')
        NewPrinterIPBox  = $window.FindName('NewPrinterIPBox')
        AddQueueBtn      = $window.FindName('AddQueueBtn')
        QueueListView    = $window.FindName('QueueListView')
        RemoveQueueBtn   = $window.FindName('RemoveQueueBtn')
        CreateBtn        = $window.FindName('CreateBtn')
        CreatePackageBtn = $window.FindName('CreatePackageBtn')
        DriverOnlyBtn    = $window.FindName('DriverOnlyBtn')
        QueueOnlyBtn     = $window.FindName('QueueOnlyBtn')
        ThemeBtn         = $window.FindName('ThemeBtn')
        VersionText      = $window.FindName('VersionText')
    }

    if ($AppVersion) { $Script:UI.VersionText.Text = "v$AppVersion" }

    Set-Theme -Dark $false

    # ── Browse .inf ──
    $Script:UI.BrowseInfBtn.Add_Click({
        $dlg = [System.Windows.Forms.OpenFileDialog]::new()
        $dlg.Title  = 'Select driver .inf file'
        $dlg.Filter = 'INF files (*.inf)|*.inf|All files (*.*)|*.*'
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $Script:InfPath          = $dlg.FileName
        $Script:InfSourceDir     = Split-Path $Script:InfPath -Parent
        $Script:DriverFolderName = Split-Path $Script:InfSourceDir -Leaf
        $Script:InfFileName      = Split-Path $Script:InfPath -Leaf

        $Script:UI.InfPathBox.Text       = $Script:InfPath
        $Script:UI.InfPathBox.Foreground = $Script:UI.Window.Resources['BrushTextBody']

        $models = Parse-InfDriverModels -InfPath $Script:InfPath
        $Script:UI.DriverModelCombo.Items.Clear()
        foreach ($m in $models) { [void]$Script:UI.DriverModelCombo.Items.Add($m) }
        if ($models.Count -gt 0) {
            $Script:UI.DriverModelCombo.SelectedIndex = 0
            Write-Log "Loaded $($models.Count) driver model(s) from $Script:InfFileName"
        } else {
            Write-Log "WARNING: No driver models found in $Script:InfFileName"
        }
    })

    # ── Add queue ──
    $Script:UI.AddQueueBtn.Add_Click({
        $name = $Script:UI.NewPrinterNameBox.Text.Trim()
        $ip   = $Script:UI.NewPrinterIPBox.Text.Trim()

        if ([string]::IsNullOrEmpty($name)) { Write-Log 'ERROR: Printer Name is required.'; return }
        if ([string]::IsNullOrEmpty($ip))   { Write-Log 'ERROR: IP Address is required.'; return }

        $parsed = $null
        if (-not [System.Net.IPAddress]::TryParse($ip, [ref]$parsed)) {
            Write-Log "ERROR: '$ip' is not a valid IP address."
            return
        }

        foreach ($item in $Script:UI.QueueListView.Items) {
            if ($item.Content -ieq $name) {
                Write-Log "ERROR: A queue named '$name' already exists."
                return
            }
        }

        $lvi         = [System.Windows.Controls.ListViewItem]::new()
        $lvi.Content = $name
        $lvi.Tag     = $ip
        [void]$Script:UI.QueueListView.Items.Add($lvi)

        $Script:UI.NewPrinterNameBox.Text = ''
        $Script:UI.NewPrinterIPBox.Text   = ''
        $Script:UI.NewPrinterNameBox.Focus() | Out-Null
    })

    # ── Remove queue ──
    $Script:UI.RemoveQueueBtn.Add_Click({
        $sel = $Script:UI.QueueListView.SelectedItem
        if ($null -eq $sel) { Write-Log 'Select a queue in the list first.'; return }
        $Script:UI.QueueListView.Items.Remove($sel)
    })

    # ── Create (no package) ──
    $Script:UI.CreateBtn.Add_Click({
        $Script:UI.LogBox.Clear()
        if (-not (Test-ForFullOrDriverOnly))  { return }
        if (-not (Test-ForQueuesPresent))     { return }

        $deployName    = $Script:UI.DeploymentNameBox.Text.Trim()
        $driverName    = $Script:UI.DriverModelCombo.SelectedItem.ToString()
        $outFolder     = Join-Path $Script:ScriptRoot "Packages\$deployName"
        $driverDest    = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }

        Write-Log "Creating deployment: $deployName"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied to $driverDest"

        $pb = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy  = New-FullDeployScript  -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName -PrintersBlock $pb
        $detect  = New-PrinterDetectScript -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView)
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1')  -Value $deploy  -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1')  -Value $detect  -Encoding UTF8
        Write-Log "Scripts written."
        Write-Log "Output: $outFolder"
        Write-IntuneCmdHint
    })

    # ── Create and Package ──
    $Script:UI.CreatePackageBtn.Add_Click({
        $Script:UI.LogBox.Clear()
        if (-not (Test-ForFullOrDriverOnly))  { return }
        if (-not (Test-ForQueuesPresent))     { return }

        $deployName    = $Script:UI.DeploymentNameBox.Text.Trim()
        $driverName    = $Script:UI.DriverModelCombo.SelectedItem.ToString()
        $outFolder     = Join-Path $Script:ScriptRoot "Packages\$deployName"
        $driverDest    = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }

        Write-Log "Creating deployment: $deployName"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied."

        $pb = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-FullDeployScript   -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName -PrintersBlock $pb
        $detect = New-PrinterDetectScript -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView)
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-Log "Scripts written."

        Invoke-Package -PackageFolder $outFolder | Out-Null
        Write-Log "Output: $outFolder"
        Write-IntuneCmdHint
    })

    # ── Package Driver Only ──
    $Script:UI.DriverOnlyBtn.Add_Click({
        $Script:UI.LogBox.Clear()
        if (-not (Test-ForFullOrDriverOnly)) { return }

        $deployName = $Script:UI.DeploymentNameBox.Text.Trim()
        $driverName = $Script:UI.DriverModelCombo.SelectedItem.ToString()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName-Driver"
        $driverDest = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }

        Write-Log "Creating driver-only deployment: $deployName-Driver"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied."

        $deploy = New-DriverOnlyDeployScript -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName
        $detect = New-DriverDetectScript     -DriverName $driverName
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-Log "Scripts written."

        Invoke-Package -PackageFolder $outFolder | Out-Null
        Write-Log "Output: $outFolder"
        Write-IntuneCmdHint
    })

    # ── Package Print Queue Only ──
    $Script:UI.QueueOnlyBtn.Add_Click({
        $Script:UI.LogBox.Clear()
        if (-not (Test-ForQueueOnly)) { return }

        $deployName  = $Script:UI.DeploymentNameBox.Text.Trim()
        $driverName  = $Script:UI.ManualDriverBox.Text.Trim()
        $outFolder   = Join-Path $Script:ScriptRoot "Packages\$deployName-QueueOnly"

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }

        Write-Log "Creating print-queue-only deployment: $deployName-QueueOnly"
        New-Item -ItemType Directory -Path $outFolder -Force | Out-Null

        $pb     = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-QueueOnlyDeployScript   -DriverName $driverName -PrintersBlock $pb
        $detect = New-PrinterDetectScript     -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView)
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-Log "Scripts written."

        Invoke-Package -PackageFolder $outFolder | Out-Null
        Write-Log "Output: $outFolder"
        Write-IntuneCmdHint
    })

    # ── Theme toggle ──
    $Script:UI.ThemeBtn.Add_Click({
        Set-Theme -Dark (-not $Script:IsDarkMode)
    })

    $window.ShowDialog() | Out-Null
}
