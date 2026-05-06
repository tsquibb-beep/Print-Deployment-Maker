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
    MinWidth="700" MinHeight="600"
    WindowStartupLocation="CenterScreen"
    FontFamily="Segoe UI" FontSize="13"
    Background="{DynamicResource BrushWinBg}">

    <Window.Resources>

        <!-- ── Theme brushes (light defaults) ── -->
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
        <!-- System highlight overrides so ListBox/ComboBox selection respects theme -->
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}"                      Color="#CCE4FF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}"                  Color="#000000"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightBrushKey}"     Color="#E0EEFF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightTextBrushKey}" Color="#000000"/>

        <!-- ── Flat button: opacity-only hover/press so any background colour works ── -->
        <Style x:Key="FlatBtn" TargetType="Button">
            <Setter Property="Foreground"               Value="White"/>
            <Setter Property="FontWeight"               Value="SemiBold"/>
            <Setter Property="BorderThickness"          Value="0"/>
            <Setter Property="Cursor"                   Value="Hand"/>
            <Setter Property="HorizontalContentAlignment" Value="Center"/>
            <Setter Property="VerticalContentAlignment"   Value="Center"/>
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
                            <ContentPresenter
                                HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
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

        <!-- ── TextBox: custom template forces Background from theme, bypasses Aero ── -->
        <Style TargetType="TextBox">
            <Setter Property="Background"  Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="CaretBrush"  Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"     Value="4,3"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                SnapsToDevicePixels="True">
                            <ScrollViewer x:Name="PART_ContentHost" Focusable="False"
                                          HorizontalScrollBarVisibility="Hidden"
                                          VerticalScrollBarVisibility="Hidden"
                                          Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.56"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush"
                                        Value="{DynamicResource BrushTextMuted}"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#0078D4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ── ListBox / ListBoxItem ── -->
        <Style TargetType="ListBox">
            <Setter Property="Background"  Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding"    Value="6,3"/>
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

        <!-- ── ListView / ListViewItem ── -->
        <Style TargetType="ListView">
            <Setter Property="Background"  Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="Background" Value="Transparent"/>
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

        <!-- ── Default button ── -->
        <Style TargetType="Button">
            <Setter Property="Cursor"      Value="Hand"/>
            <Setter Property="Background"  Value="{DynamicResource BrushControlBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextBody}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"     Value="10,4"/>
        </Style>

        <!-- ── GroupBox ── -->
        <Style TargetType="GroupBox">
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextHeader}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"     Value="6"/>
        </Style>

        <!-- ── GridViewColumnHeader ── -->
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background"  Value="{DynamicResource BrushBarBg}"/>
            <Setter Property="Foreground"  Value="{DynamicResource BrushTextMuted}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushBorder}"/>
            <Setter Property="Padding"     Value="6,3"/>
            <Setter Property="FontSize"    Value="11"/>
        </Style>

        <!-- ── TabControl: removes OS chrome so we own all tab rendering ── -->
        <Style TargetType="TabControl">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid KeyboardNavigation.TabNavigation="Local"
                              KeyboardNavigation.DirectionalNavigation="Contained">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0"
                                    Background="{DynamicResource BrushWinBg}"
                                    BorderBrush="{DynamicResource BrushBorder}"
                                    BorderThickness="0,0,0,1">
                                <TabPanel IsItemsHost="True" Background="Transparent"/>
                            </Border>
                            <ContentPresenter x:Name="PART_SelectedContentHost"
                                              Grid.Row="1"
                                              ContentSource="SelectedContent"
                                              HorizontalAlignment="Stretch"
                                              VerticalAlignment="Stretch"
                                              SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ── TabItem: themed header with selection accent line ── -->
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding"    Value="14,7"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                BorderThickness="0"
                                Padding="{TemplateBinding Padding}"
                                SnapsToDevicePixels="True">
                            <ContentPresenter ContentSource="Header"
                                              TextElement.Foreground="{TemplateBinding Foreground}"
                                              RecognizesAccessKey="True"
                                              SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Panel.ZIndex"                    Value="1"/>
                                <Setter TargetName="Bd" Property="Background"      Value="{DynamicResource BrushPanelBg}"/>
                                <Setter TargetName="Bd" Property="BorderThickness" Value="0,2,0,0"/>
                                <Setter TargetName="Bd" Property="BorderBrush"     Value="#0078D4"/>
                                <Setter Property="Foreground"                       Value="{DynamicResource BrushTextBody}"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected"  Value="False"/>
                                </MultiTrigger.Conditions>
                                <MultiTrigger.Setters>
                                    <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BrushListHover}"/>
                                </MultiTrigger.Setters>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>   <!-- header -->
            <RowDefinition Height="*"/>      <!-- shared form + tab area -->
            <RowDefinition Height="Auto"/>   <!-- collapsible log -->
        </Grid.RowDefinitions>

        <!-- ══ Header ══ -->
        <Border Grid.Row="0" Background="#0078D4" Padding="12,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Logo placeholder (96×96 spec; scaled in header) -->
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

        <!-- ══ Main area: shared form (scrollable, top) + tab strip (pinned, bottom) ══ -->
        <Grid Grid.Row="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>      <!-- shared form (scrollable, fills space) -->
                <RowDefinition Height="Auto"/>   <!-- tab control (sized to content) -->
            </Grid.RowDefinitions>

            <!-- ── Shared form ── -->
            <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto"
                          HorizontalScrollBarVisibility="Disabled">
                <StackPanel Margin="12,10,12,4">

                    <!-- Deployment -->
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

                    <!-- Driver -->
                    <GroupBox Header="Driver" Margin="0,0,0,10">
                        <Grid Margin="4,6,4,4">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="8"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="4"/>
                                <RowDefinition Height="90"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <TextBox x:Name="InfPathBox" Grid.Row="0" Grid.Column="0"
                                     IsReadOnly="True" Text="No .inf file selected"
                                     Foreground="{DynamicResource BrushTextFaint}"/>
                            <Button x:Name="BrowseInfBtn" Grid.Row="0" Grid.Column="1"
                                    Content="Browse .inf…" Margin="8,0,0,0"/>

                            <TextBlock Grid.Row="2" Grid.ColumnSpan="2"
                                       Text="Driver Model" FontSize="11"
                                       Foreground="{DynamicResource BrushTextMuted}"/>
                            <ListBox x:Name="DriverModelList" Grid.Row="4" Grid.ColumnSpan="2"
                                     ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                        </Grid>
                    </GroupBox>

                    <!-- Print Queues -->
                    <GroupBox Header="Print Queues" Margin="0,0,0,4">
                        <Grid Margin="4,6,4,4">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="4"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="6"/>
                                <RowDefinition Height="90"/>
                                <RowDefinition Height="6"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Column labels -->
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

                            <!-- Queue list — binds to PSCustomObject .Name / .IP properties -->
                            <ListView x:Name="QueueListView" Grid.Row="4">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Printer Name" Width="380">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding Name}"
                                                               Foreground="{DynamicResource BrushTextBody}"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                        <GridViewColumn Header="IP Address" Width="155">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding IP}"
                                                               Foreground="{DynamicResource BrushTextBody}"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                    </GridView>
                                </ListView.View>
                            </ListView>

                            <Button x:Name="RemoveQueueBtn" Grid.Row="6"
                                    Content="Remove Selected" HorizontalAlignment="Right"/>
                        </Grid>
                    </GroupBox>

                </StackPanel>
            </ScrollViewer><!-- end shared form -->

            <!-- ── Action tab strip (always visible at bottom of main area) ── -->
            <TabControl Grid.Row="1" Margin="12,0,12,8" MinHeight="160"
                        Background="{DynamicResource BrushPanelBg}">

                <!-- Tab 1: Create and Package -->
                <TabItem Header="Create &amp; Package">
                    <Grid Margin="12,10,12,10" VerticalAlignment="Top">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="8"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" TextWrapping="Wrap" FontSize="11"
                                   Foreground="{DynamicResource BrushTextMuted}"
                                   Text="Generates scripts, copies driver files, and packages everything as .intunewin for direct upload to Intune."/>
                        <Button x:Name="CreatePackageBtn" Grid.Row="2"
                                Content="Create and Package"
                                Style="{StaticResource FlatBtn}" Background="#107C10"
                                HorizontalAlignment="Stretch" Padding="12,10"/>
                    </Grid>
                </TabItem>

                <!-- Tab 2: Create Only -->
                <TabItem Header="Create Only">
                    <Grid Margin="12,10,12,10" VerticalAlignment="Top">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="8"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" TextWrapping="Wrap" FontSize="11"
                                   Foreground="{DynamicResource BrushTextMuted}"
                                   Text="Generates scripts and copies driver files into the output folder — no .intunewin packaging. Use this to test a deployment before packaging."/>
                        <Button x:Name="CreateBtn" Grid.Row="2"
                                Content="Create Only"
                                Style="{StaticResource FlatBtn}" Background="#0078D4"
                                HorizontalAlignment="Stretch" Padding="12,10"/>
                    </Grid>
                </TabItem>

                <!-- Tab 3: Driver Only -->
                <TabItem Header="Driver Only">
                    <Grid Margin="12,10,12,10" VerticalAlignment="Top">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="8"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" TextWrapping="Wrap" FontSize="11"
                                   Foreground="{DynamicResource BrushTextMuted}"
                                   Text="Packages the driver installation only as .intunewin. Print queues are not included — use when deploying the driver separately from any printers."/>
                        <Button x:Name="DriverOnlyBtn" Grid.Row="2"
                                Content="Package Driver Only"
                                Style="{StaticResource FlatBtn}" Background="#5C2D91"
                                HorizontalAlignment="Stretch" Padding="12,10"/>
                    </Grid>
                </TabItem>

                <!-- Tab 4: Queue Only -->
                <TabItem Header="Queue Only">
                    <Grid Margin="12,10,12,10" VerticalAlignment="Top">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="10"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="4"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="4"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="10"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" TextWrapping="Wrap" FontSize="11"
                                   Foreground="{DynamicResource BrushTextMuted}"
                                   Text="Packages print queue creation only as .intunewin. The driver must already be installed on the target device. No driver files are copied."/>
                        <TextBlock Grid.Row="2" Text="Driver Name (exact, as installed on target device)"
                                   FontSize="11" Foreground="{DynamicResource BrushTextMuted}"/>
                        <TextBox x:Name="ManualDriverBox" Grid.Row="4"/>
                        <TextBlock Grid.Row="6" FontSize="10" FontStyle="Italic"
                                   Foreground="{DynamicResource BrushTextFaint}"
                                   Text="Must match exactly — e.g. 'TOSHIBA Universal Printer 2'"/>
                        <Button x:Name="QueueOnlyBtn" Grid.Row="8"
                                Content="Package Print Queue Only"
                                Style="{StaticResource FlatBtn}" Background="#D83B01"
                                HorizontalAlignment="Stretch" Padding="12,10"/>
                    </Grid>
                </TabItem>

            </TabControl>

        </Grid><!-- end main area -->

        <!-- ══ Collapsible log ══ -->
        <Border Grid.Row="2" BorderThickness="0,1,0,0"
                BorderBrush="{DynamicResource BrushBorder}"
                Background="{DynamicResource BrushBarBg}">
            <StackPanel>
                <Button x:Name="LogToggleBtn"
                        Style="{StaticResource FlatBtn}"
                        Background="{DynamicResource BrushBarBg}"
                        Foreground="{DynamicResource BrushTextMuted}"
                        HorizontalAlignment="Stretch"
                        HorizontalContentAlignment="Left"
                        BorderThickness="0"
                        Padding="10,5"
                        FontSize="11" FontWeight="SemiBold"
                        Content="▸  Log"/>
                <ScrollViewer x:Name="LogScrollViewer" Height="110"
                              Visibility="Collapsed"
                              VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Disabled">
                    <TextBox x:Name="LogBox" IsReadOnly="True" TextWrapping="Wrap"
                             FontFamily="Consolas" FontSize="11" BorderThickness="0"
                             Background="{DynamicResource BrushBarBg}"
                             Foreground="{DynamicResource BrushTextBody}"
                             Padding="8,4"/>
                </ScrollViewer>
            </StackPanel>
        </Border>

    </Grid>
</Window>
'@

# ── Script-scope state ────────────────────────────────────────────────────────

$Script:IsDarkMode       = $false
$Script:InfPath          = ''
$Script:InfSourceDir     = ''
$Script:DriverFolderName = ''
$Script:InfFileName      = ''
$Script:ScriptRoot       = ''
$Script:UI               = @{}

# ── Theme ─────────────────────────────────────────────────────────────────────

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

# ── Logging ───────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message)
    $Script:UI.LogBox.Dispatcher.Invoke([action]{
        $Script:UI.LogBox.AppendText("$Message`n")
        $Script:UI.LogBox.ScrollToEnd()
    })
}

# ── INF parser ────────────────────────────────────────────────────────────────

function Parse-InfDriverModels {
    param([string]$InfPath)
    try {
        $lines = Get-Content -Path $InfPath -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Log "ERROR reading .inf file: $_"
        return @()
    }

    # [Strings] lookup table
    $strings  = @{}
    $inTarget = $false
    foreach ($line in $lines) {
        $clean = ($line -replace ';.*$', '').Trim()
        if ($clean -match '^\[(.+)\]') { $inTarget = $Matches[1] -ieq 'Strings'; continue }
        if ($inTarget -and $clean -match '^(\w+)\s*=\s*"(.+)"') {
            $strings[$Matches[1].ToLower()] = $Matches[2]
        }
    }

    # Manufacturer section names from [Manufacturer]
    $mfrSections = [System.Collections.Generic.List[string]]::new()
    $inMfr = $false
    foreach ($line in $lines) {
        $clean = ($line -replace ';.*$', '').Trim()
        if ($clean -match '^\[(.+)\]') { $inMfr = $Matches[1] -ieq 'Manufacturer'; continue }
        if ($inMfr -and $clean -match '=\s*(.+)') {
            foreach ($part in ($Matches[1] -split ',')) {
                $p = $part.Trim()
                if ($p -and $p -notmatch '^NT') { $mfrSections.Add($p) }
            }
        }
    }

    # Build target section name variants
    $targets = [System.Collections.Generic.List[string]]::new()
    foreach ($s in ($mfrSections | Select-Object -Unique)) {
        $targets.Add($s)
        $targets.Add("$s.NTamd64")
        $targets.Add("$s.NTamd64.10.0")
        $targets.Add("$s.NTamd64.10.0.0")
    }

    # Extract model names
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
            if ($strings.ContainsKey($key) -and $strings[$key]) { $models.Add($strings[$key]) }
        }
    }

    return @($models | Where-Object { $_ } | Sort-Object -Unique)
}

# ── Validation ────────────────────────────────────────────────────────────────

function Test-DeploymentName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { Write-Log 'ERROR: Deployment Name is required.'; return $false }
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($Name.IndexOf($ch) -ge 0) { Write-Log 'ERROR: Deployment Name contains an invalid character.'; return $false }
    }
    return $true
}

function Test-ForFullOrDriverOnly {
    if (-not $Script:InfPath) { Write-Log 'ERROR: Browse for a .inf file first.'; return $false }
    if ($null -eq $Script:UI.DriverModelList.SelectedItem) { Write-Log 'ERROR: Select a driver model from the list.'; return $false }
    if (-not (Test-DeploymentName $Script:UI.DeploymentNameBox.Text.Trim())) { return $false }
    return $true
}

function Test-ForQueuesPresent {
    if ($Script:UI.QueueListView.Items.Count -eq 0) { Write-Log 'ERROR: Add at least one print queue.'; return $false }
    return $true
}

function Test-ForQueueOnly {
    if ([string]::IsNullOrWhiteSpace($Script:UI.ManualDriverBox.Text)) {
        Write-Log 'ERROR: Manual Driver Name is required for Print Queue Only.'; return $false
    }
    if (-not (Test-DeploymentName $Script:UI.DeploymentNameBox.Text.Trim())) { return $false }
    if (-not (Test-ForQueuesPresent)) { return $false }
    return $true
}

# ── Script generation ─────────────────────────────────────────────────────────

function ConvertTo-PrinterArrayBlock {
    param([System.Windows.Controls.ListView]$ListView)
    $lines = foreach ($item in $ListView.Items) {
        $n = $item.Name -replace "'", "''"
        $i = $item.IP   -replace "'", "''"
        "    @{ Name = '$n'; IP = '$i' }"
    }
    return $lines -join "`n"
}

function ConvertTo-NamesBlock {
    param([System.Windows.Controls.ListView]$ListView)
    $quoted = foreach ($item in $ListView.Items) {
        $n = $item.Name -replace "'", "''"
        "'$n'"
    }
    return $quoted -join ', '
}

function New-FullDeployScript {
    param([string]$DriverName, [string]$DriverFolder, [string]$InfFileName, [string]$PrintersBlock)
    $template = @'
param([string]$Action = 'Install')

$DriverName      = '__DRIVER_NAME__'
$DriverFolder    = '__DRIVER_FOLDER__'
$InfFileName     = '__INF_FILE__'
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
    return $template `
        -replace '__DRIVER_NAME__',    ($DriverName   -replace "'","''") `
        -replace '__DRIVER_FOLDER__',  ($DriverFolder -replace "'","''") `
        -replace '__INF_FILE__',       ($InfFileName  -replace "'","''") `
        -replace '__PRINTERS_BLOCK__', $PrintersBlock
}

function New-DriverOnlyDeployScript {
    param([string]$DriverName, [string]$DriverFolder, [string]$InfFileName)
    $template = @'
param([string]$Action = 'Install')

$DriverName      = '__DRIVER_NAME__'
$DriverFolder    = '__DRIVER_FOLDER__'
$InfFileName     = '__INF_FILE__'
$DriverStorePath = "C:\ProgramData\AutoPilotConfig\Printers\$DriverFolder"

if ($Action -eq 'Install') {
    New-Item -ItemType Directory -Path 'C:\ProgramData\AutoPilotConfig\Printers' -Force | Out-Null
    New-Item -ItemType Directory -Path $DriverStorePath -Force | Out-Null
    Copy-Item -Path "$PSScriptRoot\$DriverFolder\*" -Destination $DriverStorePath -Recurse -Force
    & cscript 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs' `
        -a -m $DriverName -h "$DriverStorePath\" -i "$DriverStorePath\$InfFileName"
} elseif ($Action -eq 'Uninstall') {
    & cscript 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs' -d -m $DriverName
}
'@
    return $template `
        -replace '__DRIVER_NAME__',   ($DriverName   -replace "'","''") `
        -replace '__DRIVER_FOLDER__', ($DriverFolder -replace "'","''") `
        -replace '__INF_FILE__',      ($InfFileName  -replace "'","''")
}

function New-QueueOnlyDeployScript {
    param([string]$DriverName, [string]$PrintersBlock)
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
    return $template `
        -replace '__DRIVER_NAME__',    ($DriverName -replace "'","''") `
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
    $template = @'
$driverName = '__DRIVER_NAME__'
if (Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }
'@
    return $template -replace '__DRIVER_NAME__', ($DriverName -replace "'","''")
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
    if ($pkg) { Write-Log "Package: $($pkg.FullName)" }
    return $true
}

function Write-DeploymentInstructions {
    param(
        [string]$OutFolder,
        [string]$DeploymentName,
        [string]$DeploymentType,
        [string]$DriverName,
        [System.Windows.Controls.ListView]$QueueListView
    )
    $lines = @(
        "Deployment : $DeploymentName",
        "Type       : $DeploymentType",
        "Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm')",
        "Driver     : $DriverName",
        ''
    )
    if ($null -ne $QueueListView -and $QueueListView.Items.Count -gt 0) {
        $lines += 'Print Queues:'
        foreach ($item in $QueueListView.Items) {
            $lines += "  $($item.Name)  ($($item.IP))"
        }
        $lines += ''
    }
    $lines += @(
        'Intune Commands',
        '  Install  : powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Install',
        '  Uninstall: powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Uninstall'
    )
    Set-Content -Path (Join-Path $OutFolder 'deployment-info.txt') -Value $lines -Encoding UTF8
}

function Write-IntuneCmdHint {
    Write-Log '--- Intune commands ---'
    Write-Log 'Install:   powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Install'
    Write-Log 'Uninstall: powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Uninstall'
}

# ── Main window ───────────────────────────────────────────────────────────────

function Show-MainWindow {
    param(
        [string]$AppVersion = '',
        [string]$ScriptRoot = ''
    )

    # ScriptRoot must come from Start.ps1 (caller's $PSScriptRoot = project root).
    # Fallback: go up two levels from src\UI\ to the project root.
    $Script:ScriptRoot = if ([string]::IsNullOrWhiteSpace($ScriptRoot)) {
        Split-Path (Split-Path $PSScriptRoot)
    } else { $ScriptRoot }

    $reader = [System.Xml.XmlReader]::Create(
        [System.IO.StringReader]::new($Script:MainXaml)
    )
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $Script:UI = @{
        Window           = $window
        LogBox           = $window.FindName('LogBox')
        LogScrollViewer  = $window.FindName('LogScrollViewer')
        LogToggleBtn     = $window.FindName('LogToggleBtn')
        InfPathBox       = $window.FindName('InfPathBox')
        BrowseInfBtn     = $window.FindName('BrowseInfBtn')
        DriverModelList  = $window.FindName('DriverModelList')
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
        $Script:UI.DriverModelList.Items.Clear()
        foreach ($m in $models) { [void]$Script:UI.DriverModelList.Items.Add($m) }
        if ($models.Count -gt 0) {
            $Script:UI.DriverModelList.SelectedIndex = 0
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
            Write-Log "ERROR: '$ip' is not a valid IP address."; return
        }
        foreach ($item in $Script:UI.QueueListView.Items) {
            if ($item.Name -ieq $name) { Write-Log "ERROR: A queue named '$name' already exists."; return }
        }
        [void]$Script:UI.QueueListView.Items.Add([PSCustomObject]@{ Name = $name; IP = $ip })
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

    # ── Log collapse/expand ──
    $Script:UI.LogToggleBtn.Add_Click({
        if ($Script:UI.LogScrollViewer.Visibility -eq [System.Windows.Visibility]::Visible) {
            $Script:UI.LogScrollViewer.Visibility = [System.Windows.Visibility]::Collapsed
            $Script:UI.LogToggleBtn.Content       = '▸  Log'
        } else {
            $Script:UI.LogScrollViewer.Visibility = [System.Windows.Visibility]::Visible
            $Script:UI.LogToggleBtn.Content       = '▾  Log'
        }
    })

    # ── Create Only ──
    $Script:UI.CreateBtn.Add_Click({
        $Script:UI.LogBox.Clear()
        if (-not (Test-ForFullOrDriverOnly)) { return }
        if (-not (Test-ForQueuesPresent))    { return }

        $deployName = $Script:UI.DeploymentNameBox.Text.Trim()
        $driverName = $Script:UI.DriverModelList.SelectedItem.ToString()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName"
        $driverDest = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }
        Write-Log "Creating deployment: $deployName"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied."
        $pb     = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-FullDeployScript   -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName -PrintersBlock $pb
        $detect = New-PrinterDetectScript -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView)
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName $deployName `
            -DeploymentType 'Full (driver + print queues)' `
            -DriverName $driverName -QueueListView $Script:UI.QueueListView
        Write-Log "Scripts written."
        Write-Log "Output: $outFolder"
        Write-IntuneCmdHint
    })

    # ── Create and Package ──
    $Script:UI.CreatePackageBtn.Add_Click({
        $Script:UI.LogBox.Clear()
        if (-not (Test-ForFullOrDriverOnly)) { return }
        if (-not (Test-ForQueuesPresent))    { return }

        $deployName = $Script:UI.DeploymentNameBox.Text.Trim()
        $driverName = $Script:UI.DriverModelList.SelectedItem.ToString()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName"
        $driverDest = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }
        Write-Log "Creating deployment: $deployName"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied."
        $pb     = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-FullDeployScript   -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName -PrintersBlock $pb
        $detect = New-PrinterDetectScript -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView)
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName $deployName `
            -DeploymentType 'Full (driver + print queues)' `
            -DriverName $driverName -QueueListView $Script:UI.QueueListView
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
        $driverName = $Script:UI.DriverModelList.SelectedItem.ToString()
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
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName "$deployName-Driver" `
            -DeploymentType 'Driver only' `
            -DriverName $driverName -QueueListView $null
        Write-Log "Scripts written."
        Invoke-Package -PackageFolder $outFolder | Out-Null
        Write-Log "Output: $outFolder"
        Write-IntuneCmdHint
    })

    # ── Package Print Queue Only ──
    $Script:UI.QueueOnlyBtn.Add_Click({
        $Script:UI.LogBox.Clear()
        if (-not (Test-ForQueueOnly)) { return }

        $deployName = $Script:UI.DeploymentNameBox.Text.Trim()
        $driverName = $Script:UI.ManualDriverBox.Text.Trim()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName-QueueOnly"

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }
        Write-Log "Creating print-queue-only deployment: $deployName-QueueOnly"
        New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
        $pb     = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-QueueOnlyDeployScript   -DriverName $driverName -PrintersBlock $pb
        $detect = New-PrinterDetectScript     -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView)
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName "$deployName-QueueOnly" `
            -DeploymentType 'Print queues only' `
            -DriverName $driverName -QueueListView $Script:UI.QueueListView
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
