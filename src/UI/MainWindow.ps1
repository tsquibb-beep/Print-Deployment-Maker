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

                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button x:Name="ResetBtn" Content="Reset"
                            Style="{StaticResource FlatBtn}"
                            Background="#8B0000" Foreground="White" Padding="10,5"
                            Margin="0,0,8,0"/>
                    <Button x:Name="ThemeBtn" Content="☾  Dark"
                            Style="{StaticResource FlatBtn}"
                            Background="#1A5FA8" Foreground="White" Padding="10,5"/>
                </StackPanel>
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

                    <!-- Reopen existing deployment -->
                    <GroupBox Header="Reopen Existing Deployment" Margin="0,0,0,10">
                        <Grid Margin="4,6,4,4">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="8"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <ComboBox x:Name="ReopenCombo" Grid.Column="0" VerticalAlignment="Center"/>
                            <Button   x:Name="ReopenBtn"   Grid.Column="2"
                                      Content="Load" Padding="14,4"/>
                        </Grid>
                    </GroupBox>

                    <!-- Deployment -->
                    <GroupBox Header="Deployment" Margin="0,0,0,10">
                        <Grid Margin="4,6,4,4">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="4"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="10"/>
                                <ColumnDefinition Width="90"/>
                            </Grid.ColumnDefinitions>

                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Deployment Name" FontSize="11"
                                       Foreground="{DynamicResource BrushTextMuted}"/>
                            <TextBlock Grid.Row="0" Grid.Column="2" Text="Version" FontSize="11"
                                       Foreground="{DynamicResource BrushTextMuted}"/>
                            <TextBox x:Name="DeploymentNameBox"    Grid.Row="2" Grid.Column="0"/>
                            <TextBox x:Name="DeploymentVersionBox" Grid.Row="2" Grid.Column="2" Text="1"/>
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

                            <!-- Queue list — binds to PSCustomObject .Name / .IP / .SettingsSummary properties -->
                            <ListView x:Name="QueueListView" Grid.Row="4">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Printer Name" Width="240">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding Name}"
                                                               Foreground="{DynamicResource BrushTextBody}"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                        <GridViewColumn Header="IP Address" Width="130">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding IP}"
                                                               Foreground="{DynamicResource BrushTextBody}"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                        <GridViewColumn Header="Set" Width="36">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding SettingsApplied}"
                                                               HorizontalAlignment="Center"
                                                               FontWeight="Bold"
                                                               Foreground="#1B8A3A"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                        <GridViewColumn Header="Settings" Width="125">
                                            <GridViewColumn.CellTemplate>
                                                <DataTemplate>
                                                    <TextBlock Text="{Binding SettingsSummary}"
                                                               Foreground="{DynamicResource BrushTextBody}"/>
                                                </DataTemplate>
                                            </GridViewColumn.CellTemplate>
                                        </GridViewColumn>
                                    </GridView>
                                </ListView.View>
                            </ListView>

                            <!-- Printing defaults (staging printer) + remove -->
                            <Grid Grid.Row="6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <StackPanel Orientation="Horizontal">
                                        <Button x:Name="StageSettingsBtn"
                                                Content="Install staging printer &amp; open settings"
                                                Padding="10,4"/>
                                        <Button x:Name="CaptureSettingsBtn" Margin="6,0,0,0"
                                                Content="Capture to selected queue" Padding="10,4"
                                                IsEnabled="False"/>
                                    </StackPanel>
                                    <CheckBox x:Name="DevmodeCheck" Margin="2,7,0,0"
                                              Content="Capture vendor-specific driver settings (full DEVMODE)"
                                              Foreground="{DynamicResource BrushTextBody}">
                                        <CheckBox.ToolTip>
                                            <TextBlock MaxWidth="320" TextWrapping="Wrap">
                                                Use this for driver-private options the standard capture cannot see -
                                                e.g. Toshiba "Print Job" modes (Private, Hold, Scheduled print).
                                                Captures the full driver DEVMODE. Less portable: target devices must
                                                run the SAME driver version. Set these under the printer's Printing Defaults.
                                            </TextBlock>
                                        </CheckBox.ToolTip>
                                    </CheckBox>
                                    <TextBlock Margin="2,2,0,0" FontSize="10" TextWrapping="Wrap"
                                               Foreground="{DynamicResource BrushTextFaint}"
                                               Text="Tick for vendor-only modes (Private / Hold / Scheduled print). Otherwise the standard capture (duplex, color, paper) is more portable."/>
                                </StackPanel>
                                <Button x:Name="RemoveQueueBtn" Grid.Column="1" VerticalAlignment="Top"
                                        Content="Remove Selected" HorizontalAlignment="Right"/>
                            </Grid>
                        </Grid>
                    </GroupBox>

                </StackPanel>
            </ScrollViewer><!-- end shared form -->

            <!-- ── Action tab strip (always visible at bottom of main area) ── -->
            <TabControl x:Name="ActionTabs" Grid.Row="1" Margin="12,0,12,8" MinHeight="160"
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
$Script:StagingPrinterName = ''
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

function Test-DeploymentVersion {
    $v = $Script:UI.DeploymentVersionBox.Text.Trim()
    $n = 0
    if (-not [int]::TryParse($v, [ref]$n) -or $n -lt 1) {
        Write-Log 'ERROR: Version must be a whole number 1 or greater.'; return $false
    }
    return $true
}

function Test-ForFullOrDriverOnly {
    if (-not $Script:InfPath) { Write-Log 'ERROR: Browse for a .inf file first.'; return $false }
    if ($null -eq $Script:UI.DriverModelList.SelectedItem) { Write-Log 'ERROR: Select a driver model from the list.'; return $false }
    if (-not (Test-DeploymentName $Script:UI.DeploymentNameBox.Text.Trim())) { return $false }
    if (-not (Test-DeploymentVersion)) { return $false }
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
    if (-not (Test-DeploymentVersion)) { return $false }
    if (-not (Test-ForQueuesPresent)) { return $false }
    return $true
}

# ── Printer settings (staging) ────────────────────────────────────────────────

# Remove the throwaway staging printer (if any) and clear the tracked name.
function Remove-StagingPrinter {
    if ($Script:StagingPrinterName) {
        Get-Printer -Name $Script:StagingPrinterName -ErrorAction SilentlyContinue | Remove-Printer -ErrorAction SilentlyContinue
        $Script:StagingPrinterName = ''
    }
}

# Read the staging printer's current driver settings as a PrintTicket XML string.
# Uses the .NET UserPrintTicket so it reflects exactly what the printui /e
# (Printing Preferences) dialog set. On the target, Set-PrintConfiguration
# -PrintTicketXml applies the same ticket as the queue default.
function Get-StagingPrintTicket {
    param([string]$PrinterName)
    # System.Printing.LocalPrintServer/PrintQueue live in System.Printing.dll
    # (which depends on PresentationCore + ReachFramework). Load all three.
    foreach ($asm in 'PresentationCore', 'ReachFramework', 'System.Printing') {
        try { Add-Type -AssemblyName $asm -ErrorAction Stop } catch {}
    }
    if (-not ('System.Printing.LocalPrintServer' -as [type])) {
        throw 'System.Printing assembly could not be loaded.'
    }
    $server = New-Object System.Printing.LocalPrintServer
    try {
        $queue = $server.GetPrintQueue($PrinterName)
        $queue.Refresh()
        $ticket = $queue.UserPrintTicket

        $ms = New-Object System.IO.MemoryStream
        $ticket.SaveTo($ms)
        $xml = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
        $xml = $xml.TrimStart([char]0xFEFF)   # SaveTo emits a BOM; strip it so the
        $ms.Dispose()                          # XML string is clean for the target

        $parts = @()
        if ($null -ne $ticket.Duplexing) {
            switch ($ticket.Duplexing.ToString()) {
                'OneSided'           { $parts += '1-sided' }
                'TwoSidedLongEdge'   { $parts += '2-sided' }
                'TwoSidedShortEdge'  { $parts += '2-sided (short)' }
            }
        }
        if ($null -ne $ticket.OutputColor) {
            switch ($ticket.OutputColor.ToString()) {
                'Color'      { $parts += 'Color' }
                'Grayscale'  { $parts += 'Grayscale' }
                'Monochrome' { $parts += 'Mono' }
            }
        }
        $summary = if ($parts.Count) { $parts -join ', ' } else { 'captured' }
        return [PSCustomObject]@{ Blob = $xml; Summary = $summary }
    } finally {
        $server.Dispose()
    }
}

# Capture the staging printer's full default DEVMODE straight from the registry
# ('Default DevMode' = the printer's global/default settings, including the driver's
# private region). Returns base64 so it can ride on the queue item until
# Export-QueueSettingsFiles writes it out. This carries vendor-private job settings
# (Private/Hold/Scheduled print, account codes) that the PrintTicket omits, at the
# cost of portability (target must run the same driver version). Reading the registry
# is silent and instant -- printui.dll /Ss pops a dialog on some drivers and hangs.
function Get-StagingDevmode {
    param([string]$PrinterName)
    $key = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$PrinterName"
    $bytes = (Get-ItemProperty -Path $key -Name 'Default DevMode' -ErrorAction Stop).'Default DevMode'
    if (-not $bytes -or $bytes.Length -eq 0) {
        throw "No 'Default DevMode' found for '$PrinterName'."
    }
    $b64 = [System.Convert]::ToBase64String($bytes)

    # Best-effort human summary from the default print configuration.
    $summary = 'vendor (full)'
    $cfg = Get-PrintConfiguration -PrinterName $PrinterName -ErrorAction SilentlyContinue
    if ($cfg) {
        $parts = @()
        switch ("$($cfg.DuplexingMode)") {
            'OneSided'          { $parts += '1-sided' }
            'TwoSidedLongEdge'  { $parts += '2-sided' }
            'TwoSidedShortEdge' { $parts += '2-sided (short)' }
        }
        if ($null -ne $cfg.Color) { $parts += (&{ if ($cfg.Color) { 'Color' } else { 'Mono' } }) }
        if ($parts.Count) { $summary = ($parts -join ', ') + ' +vendor' }
    }
    return [PSCustomObject]@{ Blob = $b64; Summary = $summary }
}

# Write each queue's captured PrintTicket XML to <OutFolder>\settings\queueN.xml and
# stamp a transient .SettingsFile (relative path) onto the item. Queues with no
# captured settings get .SettingsFile = ''. Keeps the XML out of deploy.ps1 so that
# script stays plain ASCII.
function Export-QueueSettingsFiles {
    param([string]$OutFolder, [System.Windows.Controls.ListView]$ListView)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $i = 0
    foreach ($item in $ListView.Items) {
        $i++
        $rel = ''
        if (-not [string]::IsNullOrWhiteSpace($item.SettingsBlob)) {
            $settingsDir = Join-Path $OutFolder 'settings'
            New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
            if ($item.SettingsKind -eq 'devmode') {
                $rel = "settings\queue$i.dat"
                [System.IO.File]::WriteAllBytes((Join-Path $OutFolder $rel),
                    [System.Convert]::FromBase64String($item.SettingsBlob))
            } else {
                $rel = "settings\queue$i.xml"
                [System.IO.File]::WriteAllText((Join-Path $OutFolder $rel), $item.SettingsBlob, $utf8NoBom)
            }
        }
        $item | Add-Member -NotePropertyName SettingsFile -NotePropertyValue $rel -Force
    }
}

# ── Script generation ─────────────────────────────────────────────────────────

function ConvertTo-PrinterArrayBlock {
    param([System.Windows.Controls.ListView]$ListView)
    $lines = foreach ($item in $ListView.Items) {
        $n = $item.Name -replace "'", "''"
        $i = $item.IP   -replace "'", "''"
        $sf = ''
        if ($item.PSObject.Properties['SettingsFile']) { $sf = [string]$item.SettingsFile }
        $sf = $sf -replace "'", "''"
        $sk = ''
        if ($item.PSObject.Properties['SettingsKind']) { $sk = [string]$item.SettingsKind }
        $sk = $sk -replace "'", "''"
        "    @{ Name = '$n'; IP = '$i'; SettingsFile = '$sf'; SettingsKind = '$sk' }"
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
    param([string]$DriverName, [string]$DriverFolder, [string]$InfFileName, [string]$PrintersBlock,
          [string]$DeploymentKey, [string]$Version)
    $template = @'
param([string]$Action = 'Install')

$DriverName      = '__DRIVER_NAME__'
$DriverFolder    = '__DRIVER_FOLDER__'
$InfFileName     = '__INF_FILE__'
$DeploymentKey   = '__DEPLOY_KEY__'
$DeploymentVer   = '__VERSION__'
$VersionDir      = 'C:\ProgramData\AutoPilotConfig\PrintDeployments'
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
    $restartSpooler = $false
    foreach ($p in $Printers) {
        $portName = "TCPPort:$($p.IP)"
        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $p.IP
        }
        if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
            Add-Printer -Name $p.Name -PortName $portName -DriverName $DriverName
            if ($p.SettingsFile -and (Test-Path "$PSScriptRoot\$($p.SettingsFile)")) {
                $settingsPath = "$PSScriptRoot\$($p.SettingsFile)"
                try {
                    if ($p.SettingsKind -eq 'devmode') {
                        $regKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($p.Name)"
                        Set-ItemProperty -Path $regKey -Name 'Default DevMode' `
                            -Value ([System.IO.File]::ReadAllBytes($settingsPath))
                        $restartSpooler = $true
                    } else {
                        Set-PrintConfiguration -PrinterName $p.Name `
                            -PrintTicketXml (Get-Content $settingsPath -Raw)
                    }
                } catch {
                    Write-Warning "Could not apply settings to '$($p.Name)' - $_"
                }
            }
        } else {
            Write-Warning "Driver '$DriverName' not found - '$($p.Name)' not added"
        }
    }
    if ($restartSpooler) {
        Restart-Service -Name Spooler -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $VersionDir -Force | Out-Null
    Set-Content -Path "$VersionDir\$DeploymentKey.txt" -Value $DeploymentVer -Encoding ASCII
} elseif ($Action -eq 'Uninstall') {
    foreach ($p in $Printers) {
        Get-Printer -Name $p.Name -ErrorAction SilentlyContinue | Remove-Printer
    }
    Remove-Item "$VersionDir\$DeploymentKey.txt" -Force -ErrorAction SilentlyContinue
}
'@
    return $template `
        -replace '__DRIVER_NAME__',    ($DriverName   -replace "'","''") `
        -replace '__DRIVER_FOLDER__',  ($DriverFolder -replace "'","''") `
        -replace '__INF_FILE__',       ($InfFileName  -replace "'","''") `
        -replace '__DEPLOY_KEY__',     ($DeploymentKey -replace "'","''") `
        -replace '__VERSION__',        ($Version      -replace "'","''") `
        -replace '__PRINTERS_BLOCK__', $PrintersBlock
}

function New-DriverOnlyDeployScript {
    param([string]$DriverName, [string]$DriverFolder, [string]$InfFileName,
          [string]$DeploymentKey, [string]$Version)
    $template = @'
param([string]$Action = 'Install')

$DriverName      = '__DRIVER_NAME__'
$DriverFolder    = '__DRIVER_FOLDER__'
$InfFileName     = '__INF_FILE__'
$DeploymentKey   = '__DEPLOY_KEY__'
$DeploymentVer   = '__VERSION__'
$VersionDir      = 'C:\ProgramData\AutoPilotConfig\PrintDeployments'
$DriverStorePath = "C:\ProgramData\AutoPilotConfig\Printers\$DriverFolder"

if ($Action -eq 'Install') {
    New-Item -ItemType Directory -Path 'C:\ProgramData\AutoPilotConfig\Printers' -Force | Out-Null
    New-Item -ItemType Directory -Path $DriverStorePath -Force | Out-Null
    Copy-Item -Path "$PSScriptRoot\$DriverFolder\*" -Destination $DriverStorePath -Recurse -Force
    & cscript 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs' `
        -a -m $DriverName -h "$DriverStorePath\" -i "$DriverStorePath\$InfFileName"
    New-Item -ItemType Directory -Path $VersionDir -Force | Out-Null
    Set-Content -Path "$VersionDir\$DeploymentKey.txt" -Value $DeploymentVer -Encoding ASCII
} elseif ($Action -eq 'Uninstall') {
    & cscript 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs' -d -m $DriverName
    Remove-Item "$VersionDir\$DeploymentKey.txt" -Force -ErrorAction SilentlyContinue
}
'@
    return $template `
        -replace '__DRIVER_NAME__',   ($DriverName   -replace "'","''") `
        -replace '__DRIVER_FOLDER__', ($DriverFolder -replace "'","''") `
        -replace '__INF_FILE__',      ($InfFileName  -replace "'","''") `
        -replace '__DEPLOY_KEY__',    ($DeploymentKey -replace "'","''") `
        -replace '__VERSION__',       ($Version      -replace "'","''")
}

function New-QueueOnlyDeployScript {
    param([string]$DriverName, [string]$PrintersBlock, [string]$DeploymentKey, [string]$Version)
    $template = @'
param([string]$Action = 'Install')

$DriverName    = '__DRIVER_NAME__'
$DeploymentKey = '__DEPLOY_KEY__'
$DeploymentVer = '__VERSION__'
$VersionDir    = 'C:\ProgramData\AutoPilotConfig\PrintDeployments'
$Printers = @(
__PRINTERS_BLOCK__
)

if ($Action -eq 'Install') {
    $restartSpooler = $false
    foreach ($p in $Printers) {
        $portName = "TCPPort:$($p.IP)"
        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $p.IP
        }
        if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
            Add-Printer -Name $p.Name -PortName $portName -DriverName $DriverName
            if ($p.SettingsFile -and (Test-Path "$PSScriptRoot\$($p.SettingsFile)")) {
                $settingsPath = "$PSScriptRoot\$($p.SettingsFile)"
                try {
                    if ($p.SettingsKind -eq 'devmode') {
                        $regKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($p.Name)"
                        Set-ItemProperty -Path $regKey -Name 'Default DevMode' `
                            -Value ([System.IO.File]::ReadAllBytes($settingsPath))
                        $restartSpooler = $true
                    } else {
                        Set-PrintConfiguration -PrinterName $p.Name `
                            -PrintTicketXml (Get-Content $settingsPath -Raw)
                    }
                } catch {
                    Write-Warning "Could not apply settings to '$($p.Name)' - $_"
                }
            }
        } else {
            Write-Warning "Driver '$DriverName' not found - '$($p.Name)' not added"
        }
    }
    if ($restartSpooler) {
        Restart-Service -Name Spooler -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $VersionDir -Force | Out-Null
    Set-Content -Path "$VersionDir\$DeploymentKey.txt" -Value $DeploymentVer -Encoding ASCII
} elseif ($Action -eq 'Uninstall') {
    foreach ($p in $Printers) {
        Get-Printer -Name $p.Name -ErrorAction SilentlyContinue | Remove-Printer
    }
    Remove-Item "$VersionDir\$DeploymentKey.txt" -Force -ErrorAction SilentlyContinue
}
'@
    return $template `
        -replace '__DRIVER_NAME__',    ($DriverName -replace "'","''") `
        -replace '__DEPLOY_KEY__',     ($DeploymentKey -replace "'","''") `
        -replace '__VERSION__',        ($Version    -replace "'","''") `
        -replace '__PRINTERS_BLOCK__', $PrintersBlock
}

function New-PrinterDetectScript {
    param([string]$NamesBlock, [string]$DeploymentKey, [string]$Version)
    $template = @'
$printerNames = @(__NAMES_BLOCK__)
$missingPrinters = @()

foreach ($printerName in $printerNames) {
    $printer = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
    if ($printer -eq $null) {
        Write-Host "Printer '$printerName' not found."
        $missingPrinters += $printerName
    }
}

$versionFile = 'C:\ProgramData\AutoPilotConfig\PrintDeployments\__DEPLOY_KEY__.txt'
$expectedVersion = __VERSION__
$versionOk = $false
if (Test-Path $versionFile) {
    $current = 0
    if ([int]::TryParse((Get-Content $versionFile -Raw).Trim(), [ref]$current)) {
        if ($current -ge $expectedVersion) { $versionOk = $true }
        Write-Host "Installed version: $current (need >= $expectedVersion)."
    }
} else {
    Write-Host "Version marker not found (need >= $expectedVersion)."
}

if ($missingPrinters.Count -eq 0 -and $versionOk) {
    Write-Host "All printers found and version is current."
    Exit 0
} else {
    if ($missingPrinters.Count -gt 0) { Write-Host "Some printers are missing: $($missingPrinters -join ', ')" }
    if (-not $versionOk) { Write-Host "Version is missing or older than $expectedVersion." }
    Exit 1
}
'@
    return $template `
        -replace '__NAMES_BLOCK__',  $NamesBlock `
        -replace '__DEPLOY_KEY__',   ($DeploymentKey -replace "'","''") `
        -replace '__VERSION__',      ($Version -replace "'","''")
}

function New-DriverDetectScript {
    param([string]$DriverName, [string]$DeploymentKey, [string]$Version)
    $template = @'
$driverName = '__DRIVER_NAME__'
$versionFile = 'C:\ProgramData\AutoPilotConfig\PrintDeployments\__DEPLOY_KEY__.txt'
$expectedVersion = __VERSION__
$versionOk = $false
if (Test-Path $versionFile) {
    $current = 0
    if ([int]::TryParse((Get-Content $versionFile -Raw).Trim(), [ref]$current)) {
        if ($current -ge $expectedVersion) { $versionOk = $true }
    }
}
if ((Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue) -and $versionOk) { exit 0 } else { exit 1 }
'@
    return $template `
        -replace '__DRIVER_NAME__',  ($DriverName -replace "'","''") `
        -replace '__DEPLOY_KEY__',   ($DeploymentKey -replace "'","''") `
        -replace '__VERSION__',      ($Version -replace "'","''")
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
        [System.Windows.Controls.ListView]$QueueListView,
        [string]$Version = '1'
    )
    $lines = @(
        "Deployment : $DeploymentName",
        "Type       : $DeploymentType",
        "Version    : $Version",
        "Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm')",
        "Driver     : $DriverName",
        ''
    )
    if ($null -ne $QueueListView -and $QueueListView.Items.Count -gt 0) {
        $lines += 'Print Queues:'
        foreach ($item in $QueueListView.Items) {
            $settings = if ($item.SettingsSummary -and $item.SettingsSummary -ne 'default') {
                "  [$($item.SettingsSummary)]"
            } else { '' }
            $lines += "  $($item.Name)  ($($item.IP))$settings"
        }
        $lines += ''
    }
    $lines += @(
        'Intune Commands',
        '  Install  : powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Install',
        '  Uninstall: powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Uninstall',
        '',
        'Detection : printer/driver present AND on-target version marker >= this version.',
        '  Marker  : C:\ProgramData\AutoPilotConfig\PrintDeployments\<deployment>.txt'
    )
    Set-Content -Path (Join-Path $OutFolder 'deployment-info.txt') -Value $lines -Encoding UTF8
}

# ── Deployment manifest (machine-readable; powers Reopen/Edit) ─────────────────
# Written into every package folder as deployment.json. Captures all form state
# needed to repopulate the UI later. Call AFTER Export-QueueSettingsFiles so each
# queue carries its transient .SettingsFile (relative path to settings\queueN.*).
function Write-DeploymentManifest {
    param(
        [string]$OutFolder,
        [string]$Name,
        [string]$Version,
        [string]$Type,                 # Full | DriverOnly | QueueOnly
        [string]$DriverModel,
        [string]$DriverFolderName,
        [string]$InfFileName,
        [string]$ManualDriverName,
        [System.Windows.Controls.ListView]$QueueListView
    )
    $queues = @()
    if ($null -ne $QueueListView) {
        foreach ($item in $QueueListView.Items) {
            $sf = ''
            if ($item.PSObject.Properties['SettingsFile']) { $sf = [string]$item.SettingsFile }
            $queues += [PSCustomObject]@{
                Name            = [string]$item.Name
                IP              = [string]$item.IP
                SettingsKind    = [string]$item.SettingsKind
                SettingsSummary = [string]$item.SettingsSummary
                SettingsFile    = $sf
            }
        }
    }
    $manifest = [PSCustomObject]@{
        Name             = $Name
        Version          = $Version
        Type             = $Type
        DriverModel      = $DriverModel
        DriverFolderName = $DriverFolderName
        InfFileName      = $InfFileName
        ManualDriverName = $ManualDriverName
        Queues           = $queues
    }
    $json = $manifest | ConvertTo-Json -Depth 6
    Set-Content -Path (Join-Path $OutFolder 'deployment.json') -Value $json -Encoding UTF8
}

# List Packages\ subfolders that contain a deployment.json (reopen candidates).
function Get-ReopenableDeployments {
    $pkgRoot = Join-Path $Script:ScriptRoot 'Packages'
    if (-not (Test-Path $pkgRoot)) { return @() }
    return @(Get-ChildItem -Path $pkgRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName 'deployment.json') } |
        Select-Object -ExpandProperty Name | Sort-Object)
}

# Repopulate the whole form from a package's deployment.json. Version is auto-bumped
# by +1 (you reopen to amend + redeploy). Driver files live inside the package, so we
# point InfSourceDir at the copied folder and re-parse the copied .inf for the model list.
function Import-Deployment {
    param([string]$PackageFolder)
    $manifestPath = Join-Path $PackageFolder 'deployment.json'
    if (-not (Test-Path $manifestPath)) {
        Write-Log "ERROR: No deployment.json in '$PackageFolder'."; return
    }
    try {
        $m = Get-Content -Path $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "ERROR: Could not read deployment.json - $($_.Exception.Message)"; return
    }

    $Script:UI.DeploymentNameBox.Text = [string]$m.Name
    $nextVer = 1
    [int]::TryParse([string]$m.Version, [ref]$nextVer) | Out-Null
    $Script:UI.DeploymentVersionBox.Text = ($nextVer + 1).ToString()

    # Reset driver/inf state; restore per type.
    $Script:InfPath = ''; $Script:InfSourceDir = ''; $Script:DriverFolderName = ''; $Script:InfFileName = ''
    $Script:UI.DriverModelList.Items.Clear()
    $Script:UI.ManualDriverBox.Text = ''

    switch ([string]$m.Type) {
        'QueueOnly' {
            $Script:UI.ManualDriverBox.Text = [string]$m.ManualDriverName
            $Script:UI.ActionTabs.SelectedIndex = 3
            $Script:UI.InfPathBox.Text       = 'No .inf file selected'
            $Script:UI.InfPathBox.Foreground = $Script:UI.Window.Resources['BrushTextFaint']
        }
        default {
            # Full or DriverOnly: driver files were copied into <pkg>\<DriverFolderName>.
            $Script:DriverFolderName = [string]$m.DriverFolderName
            $Script:InfFileName      = [string]$m.InfFileName
            $Script:InfSourceDir     = Join-Path $PackageFolder $Script:DriverFolderName
            $Script:InfPath          = Join-Path $Script:InfSourceDir $Script:InfFileName
            if (Test-Path $Script:InfPath) {
                $models = Parse-InfDriverModels -InfPath $Script:InfPath
                foreach ($mod in $models) { [void]$Script:UI.DriverModelList.Items.Add($mod) }
                $Script:UI.DriverModelList.SelectedItem = [string]$m.DriverModel
                if ($null -eq $Script:UI.DriverModelList.SelectedItem -and $models.Count) {
                    $Script:UI.DriverModelList.SelectedIndex = 0
                }
                $Script:UI.InfPathBox.Text       = $Script:InfPath
                $Script:UI.InfPathBox.Foreground = $Script:UI.Window.Resources['BrushTextBody']
            } else {
                Write-Log "WARNING: copied .inf not found at $Script:InfPath"
                $Script:UI.DriverModelList.Items.Add([string]$m.DriverModel) | Out-Null
                $Script:UI.DriverModelList.SelectedIndex = 0
            }
            $Script:UI.ActionTabs.SelectedIndex = if ([string]$m.Type -eq 'DriverOnly') { 2 } else { 0 }
        }
    }

    # Rebuild queue list, reading captured settings files back into SettingsBlob.
    $Script:UI.QueueListView.Items.Clear()
    foreach ($q in @($m.Queues)) {
        $blob = ''
        $kind = [string]$q.SettingsKind
        $applied = ''
        $sf = [string]$q.SettingsFile
        if ($sf) {
            $full = Join-Path $PackageFolder $sf
            if (Test-Path $full) {
                if ($kind -eq 'devmode') {
                    $blob = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($full))
                } else {
                    $blob = [System.IO.File]::ReadAllText($full)
                }
                $applied = [char]0x2713   # check mark
            } else {
                Write-Log "WARNING: settings file missing: $sf"
            }
        }
        [void]$Script:UI.QueueListView.Items.Add([PSCustomObject]@{
            Name = [string]$q.Name; IP = [string]$q.IP
            SettingsBlob = $blob; SettingsKind = $kind
            SettingsSummary = [string]$q.SettingsSummary; SettingsApplied = $applied })
    }
    $Script:UI.QueueListView.Items.Refresh()
    Write-Log "Reopened '$($m.Name)' (was v$($m.Version), now editing as v$($nextVer + 1))."
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
        DeploymentVersionBox = $window.FindName('DeploymentVersionBox')
        ReopenCombo      = $window.FindName('ReopenCombo')
        ReopenBtn        = $window.FindName('ReopenBtn')
        ActionTabs       = $window.FindName('ActionTabs')
        NewPrinterNameBox= $window.FindName('NewPrinterNameBox')
        NewPrinterIPBox  = $window.FindName('NewPrinterIPBox')
        AddQueueBtn      = $window.FindName('AddQueueBtn')
        QueueListView    = $window.FindName('QueueListView')
        RemoveQueueBtn   = $window.FindName('RemoveQueueBtn')
        StageSettingsBtn = $window.FindName('StageSettingsBtn')
        CaptureSettingsBtn = $window.FindName('CaptureSettingsBtn')
        DevmodeCheck     = $window.FindName('DevmodeCheck')
        CreateBtn        = $window.FindName('CreateBtn')
        CreatePackageBtn = $window.FindName('CreatePackageBtn')
        DriverOnlyBtn    = $window.FindName('DriverOnlyBtn')
        QueueOnlyBtn     = $window.FindName('QueueOnlyBtn')
        ThemeBtn         = $window.FindName('ThemeBtn')
        ResetBtn         = $window.FindName('ResetBtn')
        VersionText      = $window.FindName('VersionText')
    }

    if ($AppVersion) { $Script:UI.VersionText.Text = "v$AppVersion" }

    Set-Theme -Dark $false

    # ── Reopen existing deployment ──
    $refreshReopen = {
        $sel = $Script:UI.ReopenCombo.SelectedItem
        $Script:UI.ReopenCombo.Items.Clear()
        foreach ($name in (Get-ReopenableDeployments)) { [void]$Script:UI.ReopenCombo.Items.Add($name) }
        if ($sel -and $Script:UI.ReopenCombo.Items.Contains($sel)) { $Script:UI.ReopenCombo.SelectedItem = $sel }
    }
    & $refreshReopen
    $Script:UI.ReopenCombo.Add_DropDownOpened($refreshReopen)
    $Script:UI.ReopenBtn.Add_Click({
        $name = $Script:UI.ReopenCombo.SelectedItem
        if (-not $name) { Write-Log 'Select a deployment to reopen from the list.'; return }
        $folder = Join-Path (Join-Path $Script:ScriptRoot 'Packages') $name
        Import-Deployment -PackageFolder $folder
    })

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
        [void]$Script:UI.QueueListView.Items.Add(
            [PSCustomObject]@{ Name = $name; IP = $ip; SettingsBlob = ''; SettingsKind = ''; SettingsSummary = 'default'; SettingsApplied = '' })
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

    # ── Install staging printer & open its driver settings ──
    # Installs the selected driver locally (if needed) and creates a throwaway queue
    # on the built-in FILE: port purely so the driver's settings dialog can be opened.
    # Name/IP are irrelevant: the captured PrintTicket is driver-schema only.
    $Script:UI.StageSettingsBtn.Add_Click({
        if ($null -eq $Script:UI.DriverModelList.SelectedItem) {
            Write-Log 'ERROR: Select a driver model first.'; return
        }
        if (-not $Script:InfPath) { Write-Log 'ERROR: Browse for a .inf file first.'; return }
        $driverName = $Script:UI.DriverModelList.SelectedItem.ToString()
        $stageName  = 'PDM-Staging-' + ($Script:DriverFolderName -replace '[\\/:*?"<>|]', '_')

        try {
            Remove-StagingPrinter   # clear any leftover from a previous attempt

            if (-not (Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue)) {
                Write-Log "Installing driver locally: $driverName"
                # Use the same prndrvr.vbs method the generated deploy.ps1 uses (proven).
                $prndrvr = 'C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs'
                & cscript.exe '//nologo' $prndrvr -a -m $driverName `
                    -h "$Script:InfSourceDir\" `
                    -i (Join-Path $Script:InfSourceDir $Script:InfFileName) 2>&1 |
                    ForEach-Object { Write-Log "  $_" }
                if (-not (Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue)) {
                    throw "prndrvr.vbs did not register '$driverName' (see log above). Run as Administrator if this is an access error."
                }
            }

            Add-Printer -Name $stageName -DriverName $driverName -PortName 'FILE:' -ErrorAction Stop
            $Script:StagingPrinterName = $stageName
            Write-Log "Staging printer created: $stageName"

            if ($Script:UI.DevmodeCheck.IsChecked) {
                # DEVMODE capture reads the printer's *default* (global) DevMode, so the
                # user must set options under Printing Defaults (Properties > Advanced).
                Start-Process -FilePath 'rundll32.exe' `
                    -ArgumentList 'printui.dll,PrintUIEntry', '/p', '/n', $stageName
                Write-Log 'DEVMODE mode: in Properties go to Advanced > Printing Defaults, set your vendor options (e.g. Print Job > Private/Hold), click OK, then "Capture to selected queue".'
            } else {
                # PrintTicket capture reads the user print ticket the preferences dialog sets.
                Start-Process -FilePath 'rundll32.exe' `
                    -ArgumentList 'printui.dll,PrintUIEntry', '/e', '/n', $stageName
                Write-Log 'Set your defaults in the dialog, click OK, then click "Capture to selected queue".'
            }
            $Script:UI.CaptureSettingsBtn.IsEnabled = $true
        } catch {
            Write-Log "ERROR: Could not create staging printer - $($_.Exception.Message)"
            Write-Log 'If this is an access error, run the app as Administrator.'
            Remove-StagingPrinter
        }
    })

    # ── Capture staging printer settings into the selected queue ──
    $Script:UI.CaptureSettingsBtn.Add_Click({
        if (-not $Script:StagingPrinterName) {
            Write-Log 'ERROR: Install a staging printer first.'; return
        }
        $sel = $Script:UI.QueueListView.SelectedItem
        if ($null -eq $sel) { Write-Log 'ERROR: Select the queue to apply these settings to.'; return }

        try {
            if ($Script:UI.DevmodeCheck.IsChecked) {
                $cap = Get-StagingDevmode -PrinterName $Script:StagingPrinterName
                $sel.SettingsKind = 'devmode'
            } else {
                $cap = Get-StagingPrintTicket -PrinterName $Script:StagingPrinterName
                $sel.SettingsKind = 'printticket'
            }
            $sel.SettingsBlob    = $cap.Blob
            $sel.SettingsSummary = $cap.Summary
            $sel.SettingsApplied = '✓'
            $Script:UI.QueueListView.Items.Refresh()
            Write-Log "Captured settings for '$($sel.Name)': $($cap.Summary)"
        } catch {
            Write-Log "ERROR: Could not read staging printer settings - $($_.Exception.Message)"
            return
        } finally {
            Remove-StagingPrinter
            $Script:UI.CaptureSettingsBtn.IsEnabled = $false
        }
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
        $version    = $Script:UI.DeploymentVersionBox.Text.Trim()
        $driverName = $Script:UI.DriverModelList.SelectedItem.ToString()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName"
        $markerKey  = $deployName
        $driverDest = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }
        Write-Log "Creating deployment: $deployName (v$version)"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied."
        Export-QueueSettingsFiles -OutFolder $outFolder -ListView $Script:UI.QueueListView
        $pb     = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-FullDeployScript   -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName -PrintersBlock $pb -DeploymentKey $markerKey -Version $version
        $detect = New-PrinterDetectScript -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView) -DeploymentKey $markerKey -Version $version
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName $deployName `
            -DeploymentType 'Full (driver + print queues)' -Version $version `
            -DriverName $driverName -QueueListView $Script:UI.QueueListView
        Write-DeploymentManifest -OutFolder $outFolder -Name $deployName -Version $version `
            -Type 'Full' -DriverModel $driverName -DriverFolderName $Script:DriverFolderName `
            -InfFileName $Script:InfFileName -ManualDriverName '' -QueueListView $Script:UI.QueueListView
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
        $version    = $Script:UI.DeploymentVersionBox.Text.Trim()
        $driverName = $Script:UI.DriverModelList.SelectedItem.ToString()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName"
        $markerKey  = $deployName
        $driverDest = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }
        Write-Log "Creating deployment: $deployName (v$version)"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied."
        Export-QueueSettingsFiles -OutFolder $outFolder -ListView $Script:UI.QueueListView
        $pb     = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-FullDeployScript   -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName -PrintersBlock $pb -DeploymentKey $markerKey -Version $version
        $detect = New-PrinterDetectScript -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView) -DeploymentKey $markerKey -Version $version
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName $deployName `
            -DeploymentType 'Full (driver + print queues)' -Version $version `
            -DriverName $driverName -QueueListView $Script:UI.QueueListView
        Write-DeploymentManifest -OutFolder $outFolder -Name $deployName -Version $version `
            -Type 'Full' -DriverModel $driverName -DriverFolderName $Script:DriverFolderName `
            -InfFileName $Script:InfFileName -ManualDriverName '' -QueueListView $Script:UI.QueueListView
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
        $version    = $Script:UI.DeploymentVersionBox.Text.Trim()
        $driverName = $Script:UI.DriverModelList.SelectedItem.ToString()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName-Driver"
        $markerKey  = "$deployName-Driver"
        $driverDest = Join-Path $outFolder $Script:DriverFolderName

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }
        Write-Log "Creating driver-only deployment: $deployName-Driver (v$version)"
        New-Item -ItemType Directory -Path $driverDest -Force | Out-Null
        Copy-Item -Path (Join-Path $Script:InfSourceDir '*') -Destination $driverDest -Recurse -Force
        Write-Log "Driver files copied."
        $deploy = New-DriverOnlyDeployScript -DriverName $driverName -DriverFolder $Script:DriverFolderName -InfFileName $Script:InfFileName -DeploymentKey $markerKey -Version $version
        $detect = New-DriverDetectScript     -DriverName $driverName -DeploymentKey $markerKey -Version $version
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName "$deployName-Driver" `
            -DeploymentType 'Driver only' -Version $version `
            -DriverName $driverName -QueueListView $null
        Write-DeploymentManifest -OutFolder $outFolder -Name $deployName -Version $version `
            -Type 'DriverOnly' -DriverModel $driverName -DriverFolderName $Script:DriverFolderName `
            -InfFileName $Script:InfFileName -ManualDriverName '' -QueueListView $null
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
        $version    = $Script:UI.DeploymentVersionBox.Text.Trim()
        $driverName = $Script:UI.ManualDriverBox.Text.Trim()
        $outFolder  = Join-Path $Script:ScriptRoot "Packages\$deployName-QueueOnly"
        $markerKey  = "$deployName-QueueOnly"

        if (Test-Path $outFolder) { Write-Log "WARNING: Output folder exists — contents will be overwritten." }
        Write-Log "Creating print-queue-only deployment: $deployName-QueueOnly (v$version)"
        New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
        Export-QueueSettingsFiles -OutFolder $outFolder -ListView $Script:UI.QueueListView
        $pb     = ConvertTo-PrinterArrayBlock $Script:UI.QueueListView
        $deploy = New-QueueOnlyDeployScript   -DriverName $driverName -PrintersBlock $pb -DeploymentKey $markerKey -Version $version
        $detect = New-PrinterDetectScript     -NamesBlock (ConvertTo-NamesBlock $Script:UI.QueueListView) -DeploymentKey $markerKey -Version $version
        Set-Content -Path (Join-Path $outFolder 'deploy.ps1') -Value $deploy -Encoding UTF8
        Set-Content -Path (Join-Path $outFolder 'detect.ps1') -Value $detect -Encoding UTF8
        Write-DeploymentInstructions -OutFolder $outFolder -DeploymentName "$deployName-QueueOnly" `
            -DeploymentType 'Print queues only' -Version $version `
            -DriverName $driverName -QueueListView $Script:UI.QueueListView
        Write-DeploymentManifest -OutFolder $outFolder -Name $deployName -Version $version `
            -Type 'QueueOnly' -DriverModel '' -DriverFolderName '' `
            -InfFileName '' -ManualDriverName $driverName -QueueListView $Script:UI.QueueListView
        Write-Log "Scripts written."
        Invoke-Package -PackageFolder $outFolder | Out-Null
        Write-Log "Output: $outFolder"
        Write-IntuneCmdHint
    })

    # ── Reset ──
    $Script:UI.ResetBtn.Add_Click({
        $confirm = [System.Windows.MessageBox]::Show(
            'Clear all fields and start a new deployment?',
            'Reset',
            [System.Windows.MessageBoxButton]::OKCancel,
            [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::OK) { return }

        $Script:InfPath          = ''
        $Script:InfSourceDir     = ''
        $Script:DriverFolderName = ''
        $Script:InfFileName      = ''

        $Script:UI.DeploymentNameBox.Text     = ''
        $Script:UI.DeploymentVersionBox.Text  = '1'
        $Script:UI.InfPathBox.Text         = 'No .inf file selected'
        $Script:UI.InfPathBox.Foreground   = $Script:UI.Window.Resources['BrushTextFaint']
        $Script:UI.DriverModelList.Items.Clear()
        $Script:UI.QueueListView.Items.Clear()
        $Script:UI.NewPrinterNameBox.Text  = ''
        $Script:UI.NewPrinterIPBox.Text    = ''
        $Script:UI.ManualDriverBox.Text    = ''
        Remove-StagingPrinter
        $Script:UI.CaptureSettingsBtn.IsEnabled = $false
        Write-Log "Form reset."
    })

    # ── Theme toggle ──
    $Script:UI.ThemeBtn.Add_Click({
        Set-Theme -Dark (-not $Script:IsDarkMode)
    })

    $window.ShowDialog() | Out-Null
}
