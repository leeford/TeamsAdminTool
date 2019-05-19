# Add required assemblies
Add-Type -AssemblyName System.Web, PresentationFramework, PresentationCore

function LoadMainWindow {

    param (
        
    )

    # Declare Objects
    $script:mainWindow = @{ }

    # Load XAML
    [xml]$xaml = @"
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="MainWindow"
        Title="TeamsAdminTool" SizeToContent="WidthAndHeight" FontSize="13.5" Background="#F3F2F1"
    >
    <Window.Resources>
        <!-- Inspired from https://stackoverflow.com/questions/22673483/how-do-i-create-a-flat-combo-box-using-wpf -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="{x:Type ToggleButton}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="32" />
                </Grid.ColumnDefinitions>
                <Border
                  x:Name="Border" 
                  Grid.ColumnSpan="2"
                  Height="30"
                  CornerRadius="3"
                  BorderBrush="#DFDEDE"
                  Background="White"
                  BorderThickness="1" />
                <Border 
                  Grid.Column="0"
                  CornerRadius="3,0,0,3" 
                  Margin="1"
                 
                  />
                <Path 
                  x:Name="Arrow"
                  Grid.Column="1"     
                  Fill="#6264A7"
                  HorizontalAlignment="Center"
                  VerticalAlignment="Center"
                  Data="M7.41,8.58L12,13.17L16.59,8.58L18,10L12,16L6,10L7.41,8.58Z"
                    Width="24"
                    Height="24"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter TargetName="Border" Property="Background" Value="#F3F2F1" />
                    <Setter TargetName="Border" Property="BorderBrush" Value="#DFDEDE" />
                    <Setter Property="Foreground" Value="#252424"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="#DFDEDE" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="{x:Type TextBox}">
            <Border x:Name="PART_ContentHost" Focusable="False" Background="{TemplateBinding Background}" />
        </ControlTemplate>
        <Style x:Key="{x:Type ComboBox}" TargetType="{x:Type ComboBox}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="true"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="MinHeight" Value="20"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBox}">
                        <Grid>
                            <ToggleButton 
                            Name="ToggleButton" 
                            Template="{StaticResource ComboBoxToggleButton}" 
                            Grid.Column="2" 
                            Focusable="false"
                            IsChecked="{Binding Path=IsDropDownOpen,Mode=TwoWay,RelativeSource={RelativeSource TemplatedParent}}"
                            ClickMode="Press">
                            </ToggleButton>
                            <ContentPresenter
                            Name="ContentSite"
                            IsHitTestVisible="False" 
                            Content="{TemplateBinding SelectionBoxItem}"
                            ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                            ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                            Margin="8,3,28,3"
                            VerticalAlignment="Center"
                            HorizontalAlignment="Left" />
                            <TextBox x:Name="PART_EditableTextBox"
                            Style="{x:Null}" 
                            Template="{StaticResource ComboBoxTextBox}" 
                            HorizontalAlignment="Left" 
                            VerticalAlignment="Center"
                            Focusable="True" 
                            Background="Transparent"
                            Visibility="Hidden"
                            />
                            <Popup 
                            Name="Popup"
                            Placement="Bottom"
                            IsOpen="{TemplateBinding IsDropDownOpen}"
                            AllowsTransparency="True" 
                            Focusable="False"
                            PopupAnimation="Slide">
                                <Grid 
                                  Name="DropDown"
                                  SnapsToDevicePixels="True"                
                                  MinWidth="{TemplateBinding ActualWidth}"
                                  MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border 
                                    x:Name="DropDownBorder"
                                    Background="White"
                                    BorderThickness="1"
                                    BorderBrush="#DFDEDE"
                                    CornerRadius="3"    
                                        />
                                    <ScrollViewer Margin="0,5" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="95"/>
                            </Trigger>
                            <Trigger Property="IsGrouping" Value="true">
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </Trigger>
                            <Trigger Property="IsEditable" Value="true">
                                <Setter Property="IsTabStop" Value="false"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
            </Style.Triggers>
        </Style>
        <Style x:Key="{x:Type ComboBoxItem}" TargetType="{x:Type ComboBoxItem}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBoxItem}">
                        <Border 
                      Name="Border"
                      Padding="8,5"
                      SnapsToDevicePixels="true">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="Border" Property="Background" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Foreground" Value="#252424"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!---->
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#6264A7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="Height" Value="26" />
            <Setter Property="MinWidth" Value="100" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="SnapsToDevicePixels" Value="True" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border CornerRadius="3" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter Content="{TemplateBinding Content}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#464775" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#33344A" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#BDBDBD" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="Foreground" Value="#252424" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="2,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                        >
                            <ScrollViewer x:Name="PART_ContentHost" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#6264A7" />
                                <Setter Property="BorderThickness" Value="1" />
                            </Trigger>
                            <Trigger Property="IsReadOnly" Value="True">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type PasswordBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="2,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type PasswordBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                        >
                            <ScrollViewer x:Name="PART_ContentHost" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#6264A7" />
                                <Setter Property="BorderThickness" Value="1" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type ListBoxItem}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ListBoxItem}">
                        <Border BorderBrush="{TemplateBinding BorderBrush}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#E2E2F6" />
                                <Setter Property="Foreground" Value="#252424" />
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#6264A7" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type ListBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="Padding" Value="0,5" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ListBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                Background="{TemplateBinding Background}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ScrollViewer Margin="0" Focusable="false">
                                <StackPanel IsItemsHost="True" />
                            </ScrollViewer>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type RadioButton}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type RadioButton}">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Grid Height="20" Width="20">
                                    <Border Name="Outer" Background="Transparent" BorderBrush="#DFDEDE" BorderThickness="2" CornerRadius="3" />
                                    <Border Name="Mark" CornerRadius="2" Margin="5" Visibility="Hidden" />
                                </Grid>
                            </BulletDecorator.Bullet>
                            <TextBlock Margin="5,0" Foreground="#252424" VerticalAlignment="Center">
                            <ContentPresenter />
                            </TextBlock>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Mark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="Outer" Property="BorderBrush" Value="#E2E2F6" />
                                <Setter TargetName="Mark" Property="Background" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Mark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="Outer" Property="BorderBrush" Value="#6264A7" />
                                <Setter TargetName="Mark" Property="Background" Value="#6264A7" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Outer" Property="BorderBrush" Value="#DFDEDE" />
                                <Setter TargetName="Mark" Property="Background" Value="#DFDEDE" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type CheckBox}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type CheckBox}">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Grid Height="24" Width="24">
                                    <Border Name="Outer" Background="Transparent" BorderBrush="#DFDEDE" BorderThickness="2" CornerRadius="3" Margin="2"/>
                                    <Border Name="CheckMark" Visibility="Hidden" >
                                        <Path
                                        x:Name="CheckMarkPath"
                                        Width="24" Height="24"
                                        SnapsToDevicePixels="False"
                                        StrokeThickness="2"
                                        Data="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" />
                                    </Border>
                                </Grid>
                            </BulletDecorator.Bullet>
                            <TextBlock Margin="5,0" Foreground="#252424" VerticalAlignment="Center">
                            <ContentPresenter />
                            </TextBlock>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#6264A7" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#DFDEDE" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto" />
                                <RowDefinition Height="*" />
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0">
                                <Label Foreground="#252424"
                                       FontSize="16">
                                    <ContentPresenter ContentSource="Header"/>
                                </Label>
                            </Border>
                            <Border Grid.Row="1"
                            BorderThickness="1"
                            BorderBrush="#DFDEDE"
                            CornerRadius="3"
                            >
                                <ContentPresenter />
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Grid>
                            <Border 
                                CornerRadius="3,3,0,0"
                                BorderThickness="0.5"
                                Background="{TemplateBinding Background}"
                                >
                                <ContentPresenter x:Name="TabItemCP" ContentSource="Header" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter Property="Background" Value="White"/>
                                <Setter Property="TextElement.Foreground" Value="#333333" TargetName="TabItemCP" />
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#E2E2F6" />
                                <Setter Property="TextElement.Foreground" Value="#333333" TargetName="TabItemCP" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#BDBDBD" />
                                <Setter Property="TextElement.Foreground" Value="White" TargetName="TabItemCP" />
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#6264A7"/>
                                <Setter Property="TextElement.Foreground" Value="White" TargetName="TabItemCP" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="BorderThickness" Value="1" />
        </Style>
        <Style TargetType="{x:Type DataGridRow}">
            <Setter Property="Foreground" Value="#252424"/>
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="Height" Value="30"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E2E2F6"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#6264A7"/>
                    <Setter Property="Foreground" Value="White" />
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="{x:Type DataGridCell}">
            <Setter Property="Foreground" Value="#252424"/>
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type DataGridCell}">
                        <Border Padding="10,0">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#6264A7"/>
                    <Setter Property="Foreground" Value="White" />
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="{x:Type DataGridColumnHeader}">
            <Setter Property="Foreground" Value="#252424"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type DataGridColumnHeader}">
                        <Border BorderBrush="#DFDEDE" BorderThickness="1,0,1,1">
                            <ContentPresenter Margin="10" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type DataGrid}">
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="Background" Value="White" />
            <Setter Property="SelectionUnit" Value="FullRow" />
            <Setter Property="GridLinesVisibility" Value="None" />
            <Setter Property="HeadersVisibility" Value="Column" />
        </Style>
    </Window.Resources>
    <StackPanel>
        <TabControl x:Name="mainTabControl" Margin="10">
            <TabItem x:Name="connectTabItem" Header="Connect" Width="100" Height="30">
                <StackPanel Orientation="Vertical" Margin="10">
                    <StackPanel Orientation="Vertical" Margin="0,0,0,10">
                        <TextBlock Text="TeamsAdminTool" FontSize="20" FontWeight="SemiBold" Foreground="#6264A7"/>
                        <TextBlock Text="https://github.com/leeford/TeamsAdminTool"/>
                    </StackPanel>
                    <StackPanel Orientation="Vertical" x:Name="connectionSettingsStackPanel">
                        <StackPanel Orientation="Vertical" x:Name="sharedAzureADAppStackPanel" Margin="10">
                            <RadioButton x:Name="useSharedAzureADApplicationRadioButton" Content="Use Shared Azure AD Application"  GroupName="azureADApplication" IsChecked="True"/>
                            <TextBlock Text="Use shared Azure AD Application (no further setup required)." FontStyle="Italic" />
                        </StackPanel>
                        <StackPanel Orientation="Vertical" Margin="10">
                        <RadioButton x:Name="useCustomAzureADApplicationRadioButton" Content="Use Custom Azure AD Application"  GroupName="azureADApplication" IsChecked="False"/>
                        <TextBlock Text="Only required if you don't want to use the shared Azure AD Application." FontStyle="Italic" />
                        <StackPanel Orientation="Vertical" x:Name="customAzureADAppStackPanel" IsEnabled="False">
                            <GroupBox Header="Azure AD Application" Margin="10,5">
                                <StackPanel Orientation="Vertical" Margin="5">
                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                        <TextBlock Width="150" TextAlignment="Right" VerticalAlignment="Center" Text="Application (Client) ID:"/>
                                        <TextBox x:Name="clientIdTextBox" Margin="5,0" Width="400" HorizontalAlignment="Left"/>
                                    </StackPanel>
                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                        <TextBlock Width="150" TextAlignment="Right" VerticalAlignment="Center" Text="Tenant (Directory) ID:"/>
                                        <TextBox x:Name="tenantIdTextBox" Margin="5,0" Width="400" HorizontalAlignment="Left"/>
                                    </StackPanel>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Permission Type" Margin="10,5">
                                <StackPanel Orientation="Vertical" Margin="5">
                                    <RadioButton x:Name="applicationPermissionsRadioButton" Content="Application Permissions" GroupName="Permissions"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,10">
                                        <TextBlock Width="150" TextAlignment="Right" VerticalAlignment="Center" Text="Client Secret ID:"/>
                                        <PasswordBox x:Name="clientSecretPasswordBox" Margin="5,0" Width="400" HorizontalAlignment="Left"/>
                                    </StackPanel>
                                        <RadioButton x:Name="userPermissionsRadioButton" Content="(Delegated) User Permissions" GroupName="Permissions" IsChecked="True"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,10">
                                        <TextBlock Width="150" TextAlignment="Right" VerticalAlignment="Center" Text="Redirect URI:"/>
                                        <TextBox x:Name="redirectUriTextBox" Margin="5,0" Width="400" HorizontalAlignment="Left" Text="https://login.microsoftonline.com/common/oauth2/nativeclient"/>
                                    </StackPanel>

                                </StackPanel>
                            </GroupBox>
                        </StackPanel>
                    </StackPanel>
                </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                        <Button x:Name="connectButton" Content="Connect"/>
                        <Button x:Name="disconnectButton" Content="Disconnect" IsEnabled="False"/>
                    </StackPanel>

                </StackPanel>
            </TabItem>
            <TabItem x:Name="teamsTabItem" Header="Teams" IsEnabled="False" Width="100" Height="30">
                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal" Margin="10">
                        <Button x:Name="teamsRefreshButton" Content="Refresh Teams"/>
                        <Button x:Name="addTeamButton" Content="Add Team"/>
                        <Button x:Name="teamsReportButton" Content="Teams Report"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="10,0,10,10">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="totalTeamsTextBlock" Text="Teams (0):" FontSize="20" Margin="5,0,0,5"/>
                            <TextBox x:Name="teamsFilterTextBox"/>
                            <ListBox x:Name="teamsListBox" MinWidth="220" MaxHeight="640" Margin="4,10"/>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                            <StackPanel Orientation="Vertical">
                                <TextBlock x:Name="teamDisplayNameTextBlock" FontSize="24" FontWeight="SemiBold" />
                                <TextBlock x:Name="teamDescriptionTextBlock" FontSize="16" />
                            </StackPanel>
                            <TabControl x:Name="teamTabControl" IsEnabled="False" Margin="0,10,0,0" MinWidth="1050" MinHeight="660">
                                <TabItem x:Name="teamOverviewTabItem" Header="Overview" Width="100" Height="30">
                                    <StackPanel Orientation="Vertical"  Margin="10">
                                        <GroupBox Header="Team Overview">
                                            <StackPanel Orientation="Vertical">
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Visbility:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <TextBlock x:Name="teamVisibilityTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                </StackPanel>
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Creation Date:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <TextBlock x:Name="teamCreatedDateTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                </StackPanel>
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Expiration Date:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <TextBlock x:Name="teamExpirationDateTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                </StackPanel>
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Mail:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <TextBlock x:Name="teamMailTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                </StackPanel>
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Archived:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <CheckBox x:Name="teamArchivedCheckBox" Margin="5,0" IsHitTestVisible="False" Focusable="False" VerticalAlignment="Center"/>
                                                </StackPanel>
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Members:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <TextBlock x:Name="totalTeamMembersTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                </StackPanel>
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Owners:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <TextBlock x:Name="totalTeamOwnersTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                </StackPanel>
                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                    <TextBlock Text="Channels:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                    <TextBlock x:Name="totalTeamChannelsTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                </StackPanel>
                                            </StackPanel>
                                        </GroupBox>
                                    </StackPanel>
                                </TabItem>
                                <TabItem x:Name="teamMembersTabItem" Header="Membership" Width="100" Height="30">
                                    <StackPanel Orientation="Vertical" Margin="10">
                                        <TextBlock x:Name="teamMembersTextBlock" Text="Members (0):" FontSize="16"/>
                                        <DataGrid x:Name="teamMembersDataGrid" AutoGenerateColumns="False" Margin="0,10" Height="220" SelectionMode="Single" Width="1025">
                                            <DataGrid.Columns>
                                                <DataGridTextColumn Binding="{Binding id}" IsReadOnly="True" Visibility="Hidden"/>
                                                <DataGridTextColumn Header="Name" Binding="{Binding displayName}" IsReadOnly="True" MinWidth="100" Width="Auto"/>
                                                <DataGridTextColumn Header="Title" Binding="{Binding jobTitle}" IsReadOnly="True" MinWidth="100" Width="Auto" />
                                                <DataGridTextColumn Header="User Name" Binding="{Binding userPrincipalName}" IsReadOnly="True" MinWidth="100" Width="Auto" />
                                                <DataGridTextColumn Header="Location" Binding="{Binding officeLocation}" IsReadOnly="True" MinWidth="100" Width="*" />
                                            </DataGrid.Columns>
                                        </DataGrid>
                                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                            <Button x:Name="addTeamMemberButton" Content="Add Member"/>
                                            <Button x:Name="removeTeamMemberButton" Content="Remove Member"/>
                                        </StackPanel>
                                        <TextBlock x:Name="teamOwnersTextBlock" Text="Owners (0):" Margin="0,5,0,0" FontSize="16"/>
                                        <DataGrid x:Name="teamOwnersDataGrid" AutoGenerateColumns="False" Margin="0,10" Height="220" SelectionMode="Single" Width="1025">
                                            <DataGrid.Columns>
                                                <DataGridTextColumn Binding="{Binding id}" IsReadOnly="True" Visibility="Hidden"/>
                                                <DataGridTextColumn Header="Name" Binding="{Binding displayName}" IsReadOnly="True" MinWidth="100" Width="Auto"/>
                                                <DataGridTextColumn Header="Title" Binding="{Binding jobTitle}" IsReadOnly="True" MinWidth="100" Width="Auto"/>
                                                <DataGridTextColumn Header="User Name" Binding="{Binding userPrincipalName}" IsReadOnly="True" MinWidth="100" Width="Auto"/>
                                                <DataGridTextColumn Header="Location" Binding="{Binding officeLocation}" IsReadOnly="True" MinWidth="100" Width="*"/>
                                            </DataGrid.Columns>
                                        </DataGrid>
                                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                            <Button x:Name="addTeamOwnerButton" Content="Add Owner"/>
                                            <Button x:Name="removeTeamOwnerButton" Content="Remove Owner"/>
                                        </StackPanel>
                                    </StackPanel>
                                </TabItem>
                                <TabItem x:Name="teamChannelsTabItem" Header="Channels" Width="100" Height="30">
                                    <StackPanel Orientation="Vertical" x:Name="channelsStackPanel" Margin="10">
                                        <StackPanel Orientation="Horizontal">
                                            <TextBlock x:Name="teamChannelsTextBlock" Text="Channels (0):" FontSize="16"/>
                                        </StackPanel>
                                        <StackPanel Orientation="Horizontal" Margin="0,10">
                                            <ComboBox x:Name="channelsComboBox" MinWidth="150" VerticalAlignment="Center" />
                                            <Button x:Name="addChannelButton" Content="Add Channel" VerticalAlignment="Center" />
                                        </StackPanel>
                                        <StackPanel>
                                            <TabControl x:Name="channelTabControl" IsEnabled="False" Margin="0,10,0,0">
                                                <TabItem x:Name="channelOverviewTabItem" Header="Overview" Width="100" Height="25">
                                                    <StackPanel Orientation="Vertical" Margin="10">
                                                        <GroupBox Header="Channel Overview">
                                                            <StackPanel Orientation="Vertical" Margin="5">
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Name:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <TextBlock x:Name="channelDisplayNameTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                                </StackPanel>
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Description:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <TextBlock x:Name="channelDescriptionTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                                </StackPanel>
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Mail:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <TextBlock x:Name="channelMailTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                                </StackPanel>
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Tabs:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <TextBlock x:Name="totalChannelTabsTextBlock" Width="500" TextWrapping="Wrap" Margin="5,0"/>
                                                                </StackPanel>
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Favourite:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <CheckBox x:Name="channelIsFavouriteByDefault1CheckBox" Margin="5,0" IsHitTestVisible="False" Focusable="False" VerticalAlignment="Center"/>
                                                                </StackPanel>
                                                            </StackPanel>
                                                        </GroupBox>
                                                    </StackPanel>
                                                </TabItem>
                                                <TabItem x:Name="channelTabsTabItem" Header="Tabs" Width="100" Height="25">
                                                    <StackPanel Orientation="Vertical" Margin="10">
                                                        <StackPanel Orientation="Horizontal">
                                                            <TextBlock x:Name="channelTabsTextBlock" Text="Tabs (0):" FontSize="16"/>
                                                        </StackPanel>
                                                        <StackPanel Orientation="Horizontal" Margin="0,10">
                                                            <ComboBox x:Name="channelTabsComboBox" MinWidth="150" VerticalAlignment="Center" />
                                                            <Button x:Name="addTabButton" Content="Add Tab" VerticalAlignment="Center"/>
                                                        </StackPanel>
                                                        <StackPanel x:Name="tabSettingStackPanel" Orientation="Vertical">
                                                            <GroupBox Header="Tab Overview" Margin="0,10">
                                                                <StackPanel Orientation="Vertical" Margin="5">
                                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                        <TextBlock Text="Name:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                        <TextBox x:Name="tabDisplayNameTextBox" TextWrapping="Wrap" Margin="5,0" Width="850" VerticalAlignment="Center"/>
                                                                    </StackPanel>
                                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                        <TextBlock Text="App ID:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                        <TextBox x:Name="tabTeamsAppIdTextBox" TextWrapping="Wrap" Margin="5,0" Width="850" VerticalAlignment="Center" IsEnabled="False"/>
                                                                    </StackPanel>
                                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                        <TextBlock Text="Entity ID:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                        <TextBox x:Name="tabEntityIdTextBox" TextWrapping="Wrap" Margin="5,0" Width="850" VerticalAlignment="Center" IsEnabled="False"/>
                                                                    </StackPanel>
                                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                        <TextBlock Text="Content URL:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                        <TextBox x:Name="tabContentUrlTextBox" TextWrapping="Wrap" Margin="5,0" Width="850" VerticalAlignment="Center" IsEnabled="False"/>
                                                                    </StackPanel>
                                                                    <StackPanel Orientation="Horizontal" Margin="0,2" >
                                                                        <TextBlock Text="Remove URL:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                        <TextBox x:Name="tabRemoveUrlTextBox" TextWrapping="Wrap" Margin="5,0" Width="850" VerticalAlignment="Center" IsEnabled="False"/>
                                                                    </StackPanel>
                                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                        <TextBlock Text="Website URL:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                        <TextBox x:Name="tabWebsiteUrlTextBox" TextWrapping="Wrap" Margin="5,0" Width="850" VerticalAlignment="Center" IsEnabled="False"/>
                                                                    </StackPanel>
                                                                </StackPanel>
                                                            </GroupBox>
                                                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                                                <Button x:Name="updateTabButton" Content="Save"/>
                                                                <Button x:Name="deleteTabButton" Content="Delete"/>
                                                            </StackPanel>
                                                        </StackPanel>
                                                    </StackPanel>
                                                </TabItem>
                                                <TabItem x:Name="channelSettingsTabItem" Header="Settings" Width="100" Height="25">
                                                    <StackPanel Orientation="Vertical" Margin="10">
                                                        <GroupBox Header="General">
                                                            <StackPanel Orientation="Vertical" Margin="5">
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Name:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <TextBox x:Name="channelDisplayNameTextBox" Width="700" TextWrapping="Wrap" Margin="5,0"/>
                                                                </StackPanel>
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Description:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <TextBox x:Name="channelDescriptionTextBox" Width="700" TextWrapping="Wrap" Margin="5,0"/>
                                                                </StackPanel>
                                                                <StackPanel Orientation="Horizontal" Margin="0,2">
                                                                    <TextBlock Text="Favourite:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                                    <CheckBox x:Name="channelIsFavouriteByDefault2CheckBox" Margin="5,0" VerticalAlignment="Center"/>
                                                                </StackPanel>
                                                            </StackPanel>
                                                        </GroupBox>
                                                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom" Margin="0,10">
                                                            <Button x:Name="updateChannelButton" Content="Save" />
                                                            <Button x:Name="refreshChannelButton" Content="Refresh" />
                                                        </StackPanel>
                                                    </StackPanel>
                                                </TabItem>
                                                <TabItem x:Name="channelActionsTabItem" Header="Actions" Width="100" Height="25">
                                                    <StackPanel Orientation="Vertical" Margin="10" >
                                                        <StackPanel Orientation="Horizontal" Margin="5">
                                                            <Button x:Name="deleteChannelButton" Content="Delete" />
                                                            <TextBlock Text="Delete this Channel" VerticalAlignment="Center" />
                                                        </StackPanel>
                                                    </StackPanel>
                                                </TabItem>
                                            </TabControl>
                                        </StackPanel>
                                    </StackPanel>
                                </TabItem>
                                <TabItem x:Name="teamSettingsTabItem" Header="Settings" Width="100" Height="30" >
                                    <StackPanel Orientation="Vertical" Margin="10">
                                        <StackPanel x:Name="teamSettingsStackPanel" Orientation="Vertical">
                                            <GroupBox Header="General">
                                                <StackPanel Orientation="Vertical" Margin="5">
                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                        <TextBlock Text="Name:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                        <TextBox x:Name="teamDisplayNameTextBox" Width="700" TextWrapping="Wrap" Margin="5,0"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                        <TextBlock Text="Description:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                                        <TextBox x:Name="teamDescriptionTextBox" Width="700" TextWrapping="Wrap" Margin="5,0"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                                        <TextBlock Text="Privacy:" Width="120" VerticalAlignment="Center" TextAlignment="Right"/>
                                                        <ComboBox x:Name="teamPrivacyComboBox" MinWidth="100" VerticalAlignment="Center" Margin="5,0"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </GroupBox>
                                            <GroupBox Header="Member Settings">
                                                <StackPanel Orientation="Horizontal" Margin="30,10">
                                                    <StackPanel Orientation="Vertical" Width="300">
                                                        <CheckBox x:Name="allowCreateUpdateChannelsCheckBox" Content="Create/Update Channels"/>
                                                        <CheckBox x:Name="allowDeleteChannelsCheckBox" Content="Delete Channels"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" Margin="10,0,0,0" Width="300">
                                                        <CheckBox x:Name="allowCreateUpdateRemoveTabsCheckBox" Content="Create/Update/Remove Tabs"/>
                                                        <CheckBox x:Name="allowAddRemoveAppsCheckBox" Content="Add/Remove Apps"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" Margin="10,0,0,0" Width="300">
                                                        <CheckBox x:Name="allowCreateUpdateRemoveConnectorsCheckBox" Content="Create/Update/Remove Connectors"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </GroupBox>
                                            <GroupBox Header="Guest Settings">
                                                <StackPanel Orientation="Horizontal" Margin="30,10">
                                                    <StackPanel Orientation="Vertical" Width="300">
                                                        <CheckBox x:Name="allowGuestCreateUpdateChannelsCheckBox" Content="Create/Update Channels"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" Margin="10,0,0,0" Width="300">
                                                        <CheckBox x:Name="allowGuestDeleteChannelsCheckBox" Content="Delete Channels"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </GroupBox>
                                            <GroupBox Header="Messaging Settings">
                                                <StackPanel Orientation="Horizontal" Margin="30,10">
                                                    <StackPanel Orientation="Vertical" Width="300">
                                                        <CheckBox x:Name="allowUserEditMessagesCheckBox" Content="Members Edit Messages"/>
                                                        <CheckBox x:Name="allowUserDeleteMessagesCheckBox" Content="Members Delete Messages"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" Margin="10,0,0,0" Width="300">
                                                        <CheckBox x:Name="allowTeamMentionsCheckBox" Content="Team Mentions"/>
                                                        <CheckBox x:Name="allowChannelMentionsCheckBox" Content="Channel Mentions"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" Margin="10,0,0,0" Width="300">
                                                        <CheckBox x:Name="allowOwnerDeleteMessagesCheckBox" Content="Owners Delete Messages"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </GroupBox>
                                            <GroupBox Header="Fun Settings">
                                                <StackPanel Orientation="Horizontal" Margin="30,10">
                                                    <StackPanel Orientation="Vertical" Width="300">
                                                        <CheckBox x:Name="allowGiphyCheckBox" Content="Allow Giphy"/>
                                                        <StackPanel Orientation="Horizontal">
                                                            <Label Content="Giphy Rating:"/>
                                                            <ComboBox x:Name="giphyContentRatingComboBox" MinWidth="100" Margin="5,0,0,0"/>
                                                        </StackPanel>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" Margin="10,0,0,0" Width="300">
                                                        <CheckBox x:Name="allowStickersAndMemesCheckBox" Content="Allow Stickers and Memes"/>
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" Margin="10,0,0,0" Width="300">
                                                        <CheckBox x:Name="allowCustomMemesCheckBox" Content="Allow Custom Memes"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </GroupBox>
                                            <GroupBox Header="Discovery Settings">
                                                <StackPanel Orientation="Horizontal" Margin="30,10">
                                                    <StackPanel Orientation="Vertical" Width="300">
                                                        <CheckBox x:Name="showInTeamsSearchAndSuggestionsCheckBox" Content="Show In Teams Search and Suggestions"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </GroupBox>
                                        </StackPanel>
                                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom" Margin="0,10">
                                            <Button x:Name="updateTeamButton" Content="Save"/>
                                            <Button x:Name="refreshTeamButton" Content="Refresh"/>
                                        </StackPanel>
                                    </StackPanel>
                                </TabItem>
                                <TabItem x:Name="teamActionsTabItem" Header="Actions" Width="100" Height="30" >
                                    <StackPanel Orientation="Vertical" Margin="10">
                                        <StackPanel Orientation="Horizontal" Margin="5">
                                            <Button x:Name="deleteTeamButton" Content="Delete"/>
                                            <TextBlock Text="Delete this Team (and group)" VerticalAlignment="Center" />
                                        </StackPanel>
                                        <StackPanel Orientation="Horizontal" Margin="5">
                                            <Button x:Name="archiveTeamButton" Content="Archive"/>
                                            <TextBlock Text="Archive this Team - Team will become read-only" VerticalAlignment="Center" />
                                        </StackPanel>
                                        <StackPanel Orientation="Horizontal" Margin="5">
                                            <Button x:Name="unArchiveTeamButton" Content="Un-Archive"/>
                                            <TextBlock Text="Un-archive this Team - Team will be restored to full functionality" VerticalAlignment="Center" />
                                        </StackPanel>
                                        <StackPanel Orientation="Horizontal" Margin="5">
                                            <Button x:Name="cloneTeamButton" Content="Clone"/>
                                            <TextBlock Text="Clone this Team" VerticalAlignment="Center" />
                                        </StackPanel>
                                    </StackPanel>
                                </TabItem>
                            </TabControl>
                        </StackPanel>
                    </StackPanel>
                </StackPanel>
            </TabItem>
        </TabControl>
    </StackPanel>
</Window>
"@

    # Feed XAML in to XMLNodeReader
    $XMLReader = (New-Object System.Xml.XmlNodeReader $xaml)

    # Create a Window Object
    $WindowObject = [Windows.Markup.XamlReader]::Load($XMLReader)

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {

        try {
            $script:mainWindow.Add($_.Name, $WindowObject.FindName($_.Name))
        }
        catch {

        }

    }

}

function GetAuthToken {

    param (

    )

    # Reset Auth Token
    $script:issuedToken = $null

    # Reset Logged In User
    $script:me = $null

    # Shared (multitenant) Azure AD App checked
    if ($script:mainWindow.useSharedAzureADApplicationRadioButton.IsChecked -eq $true) {
        
        # Set tenant ID to 'common' as multitenant Azure AD App
        $script:tenant = [string]"common"
        $script:clientId = [string]"6d84adaa-2a01-4f45-964b-180cbdbfd20d"

        # Get Issued Token for User
        $script:issuedToken = GetAuthTokenUser

        # If there is an issued token
        if ($script:issuedToken.access_token) {

            # Set the token timer
            $script:tokenTimer = Get-Date
            
            # Get Logged in User
            $script:me = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/v1.0/me"
            
        }
    
    }
    # Custom Azure AD App checked
    elseif ($script:mainWindow.useCustomAzureADApplicationRadioButton.IsChecked -eq $true) {
        
        # User permissions checked
        if ($script:mainWindow.userPermissionsRadioButton.IsChecked -eq $true) {

            # Check all required fields are valid
            $validClientId = ValidateRegex $script:mainWindow.clientIdTextBox.Text "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
            $validTenantId = ValidateRegex $script:mainWindow.tenantIdTextBox.Text "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
            $validRedirectUri = ValidateRegex $script:mainWindow.redirectUriTextBox.Text "(.+)"

            if ($validClientId -eq $true -and $validTenantId -eq $true -and $validRedirectUri -eq $true) {

                # Set tenant and client ID based off valid IDs
                $script:tenant = [string]$script:mainWindow.tenantIdTextBox.Text
                $script:clientId = [string]$script:mainWindow.clientIdTextBox.Text

                # Get Issued Token for User
                $script:issuedToken = GetAuthTokenUser

                # If there is an issued token
                if ($script:issuedToken.access_token) {

                    # Set the token timer
                    $script:tokenTimer = Get-Date

                    # Get Logged in User
                    $script:me = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/v1.0/me"

                }

                

            }
            else {

                ErrorPrompt -messageBody "Not all fields populated to request token. Ensure Client ID, Tenant ID and Redirect URI are populated correcty." -messageTitle "Not all fields populated"

            }

        }

        # Application permissions selected
        if ($script:mainWindow.applicationPermissionsRadioButton.IsChecked -eq $true) {

            # Check all required fields are valid
            $validClientId = ValidateRegex $script:mainWindow.clientIdTextBox.Text "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
            $validTenantId = ValidateRegex $script:mainWindow.tenantIdTextBox.Text "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
            $validClientSecret = ValidateRegex $script:mainWindow.clientSecretPasswordBox.Password "(.+)"

            if ($validClientId -eq $true -and $validTenantId -eq $true -and $validClientSecret -eq $true) {

                # Get Issued Token for Application
                $script:issuedToken = GetAuthTokenApplication

                # If there is an issued token
                if ($script:issuedToken.access_token) {

                    # Set the token timer
                    $script:tokenTimer = Get-Date

                }

            }
            else {

                ErrorPrompt -messageBody "Not all fields populated to request token. Ensure Client ID, Tenant ID and Client Secret are populated correcty." -messageTitle "Not all fields populated"

            }
        
        }

    }


}

function GetAuthCodeUser {
    param (

    )

    # Random State
    $state = Get-Random

    # Encode scope to fit inside query string
    $script:scope = "Group.ReadWrite.All Directory.Read.All Notes.ReadWrite.All"
    $encodedScope = [System.Web.HttpUtility]::UrlEncode($script:scope)

    # Redirect URI (encode it to fit inside query string)
    $redirectUri = [System.Web.HttpUtility]::UrlEncode($script:mainWindow.redirectUriTextBox.Text)

    # Construct URI
    $uri = [uri]"https://login.microsoftonline.com/$script:tenant/oauth2/v2.0/authorize?client_id=$script:clientId&response_type=code&redirect_uri=$redirectUri&response_mode=query&scope=$encodedScope&state=$state&prompt=login"

    # Create Window for User Sign-In
    $windowProperty = @{
        Width  = 500
        Height = 700
    }
    $signInWindow = New-Object System.Windows.Window -Property $windowProperty
    
    # Create WebBrowser for Window
    $browserProperty = @{
        Width  = 480
        Height = 680
    }
    $signInBrowser = New-Object System.Windows.Controls.WebBrowser -Property $browserProperty

    # Navigate Browser to sign-in page
    $signInBrowser.navigate($uri)
    
    # Create a condition to check after each page load
    $pageLoaded = {

        # Once a URL contains "code=*", close the Window
        if ($signInBrowser.Source -match "code=[^&]*") {

            # With the form closed and complete with the code, parse the query string

            $urlQueryString = [System.Uri]($signInBrowser.Source).Query
            $script:urlQueryValues = [System.Web.HttpUtility]::ParseQueryString($urlQueryString)

            $signInWindow.Close()

        }
    }

    # Add condition to document completed
    $signInBrowser.Add_LoadCompleted($pageLoaded)

    # Show Window
    $signInWindow.AddChild($signInBrowser)
    $signInWindow.ShowDialog()

    # Extract code and state from query string
    $authCode = $script:urlQueryValues.GetValues(($script:urlQueryValues.keys | Where-Object { $_ -eq "code" }))
    $returnedState = $script:urlQueryValues.GetValues(($script:urlQueryValues.keys | Where-Object { $_ -eq "state" }))

    # If auth code has been extracted and return state matches original state
    if ($authCode -and $state -match $returnedState) {

        return $authCode

    }
    else {

        ErrorPrompt -messageBody "Unable to obtain OAuth Code or State mismatch!" -messageTitle "Unable to obtain OAuth code"

    }
    
}

function GetAuthTokenUser {
    param (

    )
    
    # Get Auth Code needed for Token
    $authCode = GetAuthCodeUser

    # Construct URI
    $uri = [uri]"https://login.microsoftonline.com/$script:tenant/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id    = $script:clientId
        scope        = "$script:scope offline_access" # Add offline_access to scope to ensure refresh_token is issued
        code         = $authCode[1]
        redirect_uri = $script:mainWindow.redirectUriTextBox.Text
        grant_type   = "authorization_code"
    }

    # Get OAuth 2.0 Token
    $tokenRequest = try {

        Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop

    }
    catch [System.Net.WebException] {

        ErrorPrompt -messageBody "Exception was caught: $($_.Exception.Message)" -messageTitle "Unable to obtain OAuth token"
    
    }

    # If token request was a success
    if ($tokenRequest.StatusCode -eq 200) {

        return $tokenRequest.Content | ConvertFrom-Json
        
    }

}

function GetAuthTokenUserRefresh {

    param (

    )

    # Construct URI
    $uri = [uri]"https://login.microsoftonline.com/$script:tenant/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $script:clientId
        scope         = "$script:scope offline_access" # Add offline_access to scope to ensure refresh_token is issued
        redirect_uri  = $script:mainWindow.redirectUriTextBox.Text
        grant_type    = "refresh_token"
        refresh_token = $script:issuedToken.refresh_token
    }

    # Get OAuth 2.0 Token
    $tokenRequest = try {

        Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop

    }
    catch [System.Net.WebException] {

        ErrorPrompt -messageBody "Exception was caught: $($_.Exception.Message)" -messageTitle "Unable to obtain OAuth token"
    
    }

    # If token request was a success
    if ($tokenRequest.StatusCode -eq 200) {

        # Set the token timer
        $script:tokenTimer = Get-Date

        return $tokenRequest.Content | ConvertFrom-Json
        
    }

}

function GetAuthTokenApplication {

    param (

    )

    $clientId = [string]$script:mainWindow.clientIdTextBox.Text
    $tenantId = [string]$script:mainWindow.tenantIdTextBox.Text
    $clientSecret = [string]$script:mainWindow.clientSecretPasswordBox.Password

    # Construct URI
    $uri = [uri]"https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    # Get OAuth 2.0 Token
    $tokenRequest = try {

        Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop

    }
    catch [System.Net.WebException] {

        ErrorPrompt -messageBody "Exception was caught: $($_.Exception.Message)" -messageTitle "Unable to obtain OAuth token"
            
    }

    # If token request was a success
    if ($tokenRequest.StatusCode -eq 200) {

        return $tokenRequest.Content | ConvertFrom-Json
        
    }
    else {

        ErrorPrompt -messageBody "Unexpected response code" -messageTitle "Unable to obtain OAuth token"

    }

}

function InvokeGraphAPICall {
    param (
        
        [Parameter(mandatory = $true)][string]$method,
        [Parameter(mandatory = $true)][uri]$uri,
        [Parameter(mandatory = $false)]$body,
        [Parameter(mandatory = $false)][string]$contentType,
        [Parameter(mandatory = $false)][switch]$silent

    )

    # Calculate current token age
    $tokenAge = New-TimeSpan $script:tokenTimer (Get-Date)

    # Check token has not expired
    if ($tokenAge.TotalSeconds -gt 3500) {

        Write-Warning "Token May Have Expired! Refreshing..."

        # If last token issued included a refresh token (user)
        if ($script:issuedToken.refresh_token) {

            # Get new token using refresh token
            $script:issuedToken = GetAuthTokenUserRefresh

        }
        else {

            # Otherwise authenticate without
            GetAuthToken

        }

    }

    # If no content type, use application/json
    if (-not $contentType) {

        $contentType = "application/json"

    }
        
    # Construct headers
    $Headers = @{"Authorization" = "Bearer $($script:issuedToken.access_token)" }

    $apiCall = try {

        # Assume success
        $script:lastAPICallSuccess = $true

        # If there is a body (and it's not a GET), use it
        if ($body -and $method -ne "GET") {

            $bodyJson = $body | ConvertTo-Json -Depth 5 | ForEach-Object { [regex]::Unescape($_) }

            Invoke-WebRequest -Method $method -Uri $uri -ContentType $contentType -Headers $Headers -Body $bodyJson -ErrorAction Stop

        }
        else {

            Invoke-WebRequest -Method $method -Uri $uri -ContentType $contentType -Headers $Headers -ErrorAction Stop

        }

    }
    catch [System.Net.WebException] {

        if (-not $silent) {

            # Error handling taken from http://wahlnetwork.com/2015/02/19/using-try-catch-powershells-invoke-webrequest/ 
            $responseResult = $_.Exception.Response.GetResponseStream()
            $responseReader = New-Object System.IO.StreamReader($responseResult)
            $responseBody = $responseReader.ReadToEnd() | ConvertFrom-Json            
            
            ErrorPrompt -messageBody "Exception was caught: $($_.Exception.Message) `rURI: $uri `rResponse Code: $($responseBody.error.code) `rResponse Message: $($responseBody.error.message)" -messageTitle "Error on Graph API call"

        }

        $script:lastAPICallSuccess = $false

    }

    # Store any response incase it is needed
    $script:lastAPICallReponse = $apiCall

    # Return content (assume it's in JSON format)
    if ($apiCall.Content) {

        return $apiCall.Content | ConvertFrom-Json  

    }

}
function ValidateRegex {
    param (
        
        [Parameter(mandatory = $true)][string]$value,
        [Parameter(mandatory = $true)][string]$regex

    )

    if ($value -match $regex) {

        return $true

    }
    else {

        return $false

    }

}

function GetTeamInformation {
    param (


    )

    if ($script:mainWindow.teamsListBox.SelectedValue) {

        $script:currentTeamId = $script:mainWindow.teamsListBox.SelectedValue

    }

    # Reset current channel and tab ID and disable channel tab control (as Team may have changed)
    $script:currentChannelId = $null
    $script:currentTabId = $null
    $script:mainWindow.channelTabControl.IsEnabled = $false

    if (-not [string]::IsNullOrEmpty($script:currentTeamId)) {

        # Team #
        ########
        $team = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId"
        $script:mainWindow.teamArchivedCheckBox.IsChecked = $team.isArchived
        # Member Settings
        $script:mainWindow.allowCreateUpdateChannelsCheckBox.IsChecked = $team.memberSettings.allowCreateUpdateChannels
        $script:mainWindow.allowDeleteChannelsCheckBox.IsChecked = $team.memberSettings.allowDeleteChannels
        $script:mainWindow.allowAddRemoveAppsCheckBox.IsChecked = $team.memberSettings.allowAddRemoveApps
        $script:mainWindow.allowCreateUpdateRemoveTabsCheckBox.IsChecked = $team.memberSettings.allowCreateUpdateRemoveTabs
        $script:mainWindow.allowCreateUpdateRemoveConnectorsCheckBox.IsChecked = $team.memberSettings.allowCreateUpdateRemoveConnectors
        # Guest Settings
        $script:mainWindow.allowGuestCreateUpdateChannelsCheckBox.IsChecked = $team.guestSettings.allowCreateUpdateChannels
        $script:mainWindow.allowGuestDeleteChannelsCheckBox.IsChecked = $team.guestSettings.allowDeleteChannels
        # Messaging Settings
        $script:mainWindow.allowUserEditMessagesCheckBox.IsChecked = $team.messagingSettings.allowUserEditMessages
        $script:mainWindow.allowUserDeleteMessagesCheckBox.IsChecked = $team.messagingSettings.allowUserDeleteMessages
        $script:mainWindow.allowOwnerDeleteMessagesCheckBox.IsChecked = $team.messagingSettings.allowOwnerDeleteMessages
        $script:mainWindow.allowTeamMentionsCheckBox.IsChecked = $team.messagingSettings.allowTeamMentions
        $script:mainWindow.allowChannelMentionsCheckBox.IsChecked = $team.messagingSettings.allowChannelMentions
        # Fun Settings
        $script:mainWindow.allowGiphyCheckBox.IsChecked = $team.funSettings.allowGiphy
        $script:mainWindow.giphyContentRatingComboBox.SelectedItem = (Get-Culture).TextInfo.ToTitleCase($team.funSettings.giphyContentRating)
        $script:mainWindow.allowStickersAndMemesCheckBox.IsChecked = $team.funSettings.allowStickersAndMemes
        $script:mainWindow.allowCustomMemesCheckBox.IsChecked = $team.funSettings.allowCustomMemes
        # Discover
        $script:mainWindow.showInTeamsSearchAndSuggestionsCheckBox.IsChecked = $team.discoverySettings.showInTeamsSearchAndSuggestions

        # Group #
        #########
        $group = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId"
        $script:mainWindow.teamDisplayNameTextBlock.Text = $group.displayName
        $script:mainWindow.teamDescriptionTextBlock.Text = $group.description
        $script:mainWindow.teamVisibilityTextBlock.Text = $group.visibility
        $script:mainWindow.teamPrivacyComboBox.SelectedItem = (Get-Culture).TextInfo.ToTitleCase($group.visibility)
        $script:mainWindow.teamCreatedDateTextBlock.Text = (Get-Date -Date $group.createdDateTime -Format 'g')
        # Expiry date
        if ($group.expirationDateTime) {

            $script:mainWindow.teamExpirationDateTextBlock.Text = (Get-Date -Date $group.expirationDateTime -Format 'g')

        }
        else {

            $script:mainWindow.teamExpirationDateTextBlock.Text = "None"

        }
        $script:mainWindow.teamMailTextBlock.Text = $group.mail
        $script:mainWindow.teamDisplayNameTextBox.Text = $group.displayName
        $script:mainWindow.teamDescriptionTextBox.Text = $group.description

        # Members #
        ###########
        $members = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/members"
        $totalMembers = (($members).value).Count
        $script:mainWindow.totalTeamMembersTextBlock.Text = $totalMembers
        $script:mainWindow.teamMembersTextBlock.Text = "Members ($totalMembers):"
        $membersDataGrid = @()
        $members.Value | ForEach-Object {

            $member = @{

                displayName       = $_.displayName
                userPrincipalName = $_.userPrincipalName
                jobTitle          = $_.jobTitle
                officeLocation    = $_.officeLocation
                id                = $_.id

            }

            $membersDataGrid += New-Object PSObject -Property $member

        }

        $script:mainWindow.teamMembersDataGrid.ItemsSource = $membersDataGrid

        # Owners #
        ##########
        $owners = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/owners"
        $totalOwners = (($owners).value).Count
        $script:mainWindow.totalTeamOwnersTextBlock.Text = $totalOwners
        $script:mainWindow.teamOwnersTextBlock.Text = "Owners ($totalOwners):"
        $ownersDataGrid = @()
        $owners.Value | ForEach-Object {

            $owner = @{

                displayName       = $_.displayName
                userPrincipalName = $_.userPrincipalName
                jobTitle          = $_.jobTitle
                officeLocation    = $_.officeLocation
                id                = $_.id

            }

            $ownersDataGrid += New-Object PSObject -Property $owner

        }

        $script:mainWindow.teamOwnersDataGrid.ItemsSource = $ownersDataGrid
        # Disable delete owner button if one owner
        if ($totalOwners -le 1) {

            $script:mainWindow.removeTeamOwnerButton.IsEnabled = $false

        }
        else {

            $script:mainWindow.removeTeamOwnerButton.IsEnabled = $true

        }

        # Channels #
        ############
        $channels = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels"
        $totalChannels = (($channels).value).Count
        $script:mainWindow.totalTeamChannelsTextBlock.Text = $totalChannels
        $script:mainWindow.teamChannelsTextBlock.Text = "Channels ($totalChannels):"
        $script:mainWindow.channelsComboBox.DisplayMemberPath = "displayName"
        $script:mainWindow.channelsComboBox.SelectedValuePath = "id"
        $script:mainWindow.channelsComboBox.ItemsSource = $channels.value

        # Enable Tabs
        $script:mainWindow.teamTabControl.IsEnabled = $true

        # Set UI based on archive status
        # Archived
        if ($team.isArchived -eq $true) {
            
            $script:mainWindow.teamSettingsStackPanel.IsEnabled = $false
            $script:mainWindow.channelsStackPanel.IsEnabled = $false
            $script:mainWindow.archiveTeamButton.IsEnabled = $false
            $script:mainWindow.updateTeamButton.IsEnabled = $false
            $script:mainWindow.unarchiveTeamButton.IsEnabled = $true

            
        }
        # Not archived
        elseif ($team.isArchived -eq $false) {
            
            $script:mainWindow.teamSettingsStackPanel.IsEnabled = $true
            $script:mainWindow.channelsStackPanel.IsEnabled = $true
            $script:mainWindow.archiveTeamButton.IsEnabled = $true
            $script:mainWindow.updateTeamButton.IsEnabled = $true
            $script:mainWindow.unarchiveTeamButton.IsEnabled = $false

        }
    
    }
}

function GetChannelInformation {
    param (


    )

    if ($script:mainWindow.channelsComboBox.SelectedValue) {

        $script:currentChannelId = $script:mainWindow.channelsComboBox.SelectedValue

    }

    if (-not [string]::IsNullOrEmpty($script:currentChannelId)) {

        # Reset Tab Id and disable Tab as Channel may have changed
        $script:currentTabId = $null
        $script:mainWindow.tabSettingStackPanel.IsEnabled = $false

        # Channel #
        ###########
        $channel = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId"
        $script:mainWindow.channelDisplayNameTextBlock.Text = $channel.displayName
        $script:mainWindow.channelDescriptionTextBlock.Text = $channel.description
        $script:mainWindow.channelDisplayNameTextBox.Text = $channel.displayName
        $script:mainWindow.channelDescriptionTextBox.Text = $channel.description
        $script:mainWindow.channelMailTextBlock.Text = $channel.email
        if (-not [string]::IsNullOrEmpty($channel.isFavoriteByDefault)) {

            $script:mainWindow.channelIsFavouriteByDefault1CheckBox.IsChecked = $channel.isFavoriteByDefault
            $script:mainWindow.channelIsFavouriteByDefault2CheckBox.IsChecked = $channel.isFavoriteByDefault

        }
        else {

            $script:mainWindow.channelIsFavouriteByDefault1CheckBox.IsChecked = $false
            $script:mainWindow.channelIsFavouriteByDefault2CheckBox.IsChecked = $false

        }

        # Tabs #
        ########
        $tabs = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs"
        $totalTabs = (($tabs).value).Count
        $script:mainWindow.totalChannelTabsTextBlock.Text = $totalTabs
        $script:mainWindow.channelTabsTextBlock.Text = "Tabs ($totalTabs):"
        $script:mainWindow.channelTabsComboBox.DisplayMemberPath = "displayName"
        $script:mainWindow.channelTabsComboBox.SelectedValuePath = "id"
        $script:mainWindow.channelTabsComboBox.ItemsSource = $tabs.value

        # Enable Tabs
        $script:mainWindow.channelTabControl.IsEnabled = $true
    
    }

}

function GetTabInformation {
    param (
        
    )

    if ($script:mainWindow.channelTabsComboBox.SelectedValue) {

        $script:currentTabId = $script:mainWindow.channelTabsComboBox.SelectedValue

    }

    if (-not [string]::IsNullOrEmpty($script:currentTabId)) {

        # Tab #
        #######
        $tab = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs/$script:currentTabId"
        $script:mainWindow.tabTeamsAppIdTextBox.Text = $tab.teamsAppId
        $script:mainWindow.tabDisplayNameTextBox.Text = $tab.displayName
        $script:mainWindow.tabEntityIdTextBox.Text = $tab.configuration.entityId
        $script:mainWindow.tabContentUrlTextBox.Text = $tab.configuration.contentUrl
        $script:mainWindow.tabRemoveUrlTextBox.Text = $tab.configuration.removeUrl
        $script:mainWindow.tabWebsiteUrlTextBox.Text = $tab.configuration.websiteUrl

        # Enable UI
        $script:mainWindow.tabSettingStackPanel.IsEnabled = $true
    
    }
    
}
function ListTeams {
    param (
        
    )
    
    # List Teams
    $script:teams = InvokeGraphAPICall -method "GET" -uri "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"

    # Update Total
    $totalTeams = (($script:teams).value).Count
    $script:mainWindow.totalTeamsTextBlock.Text = "Teams ($totalTeams):"

    # Populate List
    $script:mainWindow.teamsListBox.Visibility = "Visible"
    $script:mainWindow.teamsListBox.DisplayMemberPath = "displayName"
    $script:mainWindow.teamsListBox.SelectedValuePath = "id"
    $script:mainWindow.teamsListBox.ItemsSource = $script:teams.value | Sort-Object -Property displayName

}

function FilterTeams {

    # Get text from textbox
    $teamsFilterText = $script:mainWindow.teamsFilterTextBox.text
    
    $script:mainWindow.teamsListBox.DisplayMemberPath = "displayName"
    $script:mainWindow.teamsListBox.SelectedValuePath = "id"

    # Filter Teams
    $filteredTeams = $script:teams.value | Where-Object { $_.displayName -like "*$teamsFilterText*" } | Sort-Object -Property displayName
    
    # Teams Total
    $totalTeams = ($filteredTeams).Count
    
    # 1 or more found
    if ($totalTeams -gt 1) {

        $script:mainWindow.teamsListBox.Visibility = "Visible"
        $script:mainWindow.teamsListBox.ItemsSource = $filteredTeams
        $script:mainWindow.totalTeamsTextBlock.Text = "Teams ($totalTeams):"

        
    } # 0 found
    elseif ($totalTeams -eq 0) {

        # Hide ListBox
        $script:mainWindow.teamsListBox.Visibility = "Hidden"
        $script:mainWindow.teamsListBox.ItemsSource = $null

        $script:mainWindow.totalTeamsTextBlock.Text = "Teams (0):"

    } # 1 found (1 doesn't return on the count, annoyingly!)
    else {

        # Hide ListBox
        $script:mainWindow.teamsListBox.Visibility = "Hidden"
        $script:mainWindow.teamsListBox.ItemsSource = $null

        $script:mainWindow.totalTeamsTextBlock.Text = "Teams (1):"

        # Change selected Team name and description in UI
        $script:mainWindow.teamDisplayNameTextBlock.Text = $filteredTeams.displayName
        $script:mainWindow.teamDescriptionTextBlock.Text = $filteredTeams.description

        # As there is 1 Team, load it
        $script:currentTeamId = $filteredTeams.id
        GetTeamInformation

    }

}

function AddTeam {

    param (

    )

    # Load Window
    # Declare Objects
    $window = @{ }

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="AddTeamWindow"
    Title="Add Team" MinWidth="640" SizeToContent="WidthAndHeight" FontSize="13.5"
    >
    <Window.Resources>
        <Style TargetType="{x:Type DataGridRow}">
            <Setter Property="Foreground" Value="#252424"/>
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="Height" Value="30"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E2E2F6"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#6264A7"/>
                    <Setter Property="Foreground" Value="White" />
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="{x:Type DataGridCell}">
            <Setter Property="Foreground" Value="#252424"/>
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type DataGridCell}">
                        <Border Padding="10,0">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#6264A7"/>
                    <Setter Property="Foreground" Value="White" />
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="{x:Type DataGridColumnHeader}">
            <Setter Property="Foreground" Value="#252424"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type DataGridColumnHeader}">
                        <Border BorderBrush="#DFDEDE" BorderThickness="1,0,1,1">
                            <ContentPresenter Margin="10" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type DataGrid}">
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="Background" Value="White" />
            <Setter Property="SelectionUnit" Value="FullRow" />
            <Setter Property="GridLinesVisibility" Value="None" />
            <Setter Property="HeadersVisibility" Value="Column" />
        </Style>
        <!-- Inspired from https://stackoverflow.com/questions/22673483/how-do-i-create-a-flat-combo-box-using-wpf -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="{x:Type ToggleButton}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="32" />
                </Grid.ColumnDefinitions>
                <Border
                  x:Name="Border" 
                  Grid.ColumnSpan="2"
                  Height="30"
                  CornerRadius="3"
                  BorderBrush="#DFDEDE"
                  Background="White"
                  BorderThickness="1" />
                <Border 
                  Grid.Column="0"
                  CornerRadius="3,0,0,3" 
                  Margin="1"
                 
                  />
                <Path 
                  x:Name="Arrow"
                  Grid.Column="1"     
                  Fill="#6264A7"
                  HorizontalAlignment="Center"
                  VerticalAlignment="Center"
                  Data="M7.41,8.58L12,13.17L16.59,8.58L18,10L12,16L6,10L7.41,8.58Z"
                    Width="24"
                    Height="24"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter TargetName="Border" Property="Background" Value="#F3F2F1" />
                    <Setter TargetName="Border" Property="BorderBrush" Value="#DFDEDE" />
                    <Setter Property="Foreground" Value="#252424"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="#DFDEDE" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="{x:Type TextBox}">
            <Border x:Name="PART_ContentHost" Focusable="False" Background="{TemplateBinding Background}" />
        </ControlTemplate>
        <Style x:Key="{x:Type ComboBox}" TargetType="{x:Type ComboBox}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="true"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="MinHeight" Value="20"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBox}">
                        <Grid>
                            <ToggleButton 
                            Name="ToggleButton" 
                            Template="{StaticResource ComboBoxToggleButton}" 
                            Grid.Column="2" 
                            Focusable="false"
                            IsChecked="{Binding Path=IsDropDownOpen,Mode=TwoWay,RelativeSource={RelativeSource TemplatedParent}}"
                            ClickMode="Press">
                            </ToggleButton>
                            <ContentPresenter
                            Name="ContentSite"
                            IsHitTestVisible="False" 
                            Content="{TemplateBinding SelectionBoxItem}"
                            ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                            ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                            Margin="8,3,28,3"
                            VerticalAlignment="Center"
                            HorizontalAlignment="Left" />
                            <TextBox x:Name="PART_EditableTextBox"
                            Style="{x:Null}" 
                            Template="{StaticResource ComboBoxTextBox}" 
                            HorizontalAlignment="Left" 
                            VerticalAlignment="Center"
                            Focusable="True" 
                            Background="Transparent"
                            Visibility="Hidden"
                            />
                            <Popup 
                            Name="Popup"
                            Placement="Bottom"
                            IsOpen="{TemplateBinding IsDropDownOpen}"
                            AllowsTransparency="True" 
                            Focusable="False"
                            PopupAnimation="Slide">
                                <Grid 
                                  Name="DropDown"
                                  SnapsToDevicePixels="True"                
                                  MinWidth="{TemplateBinding ActualWidth}"
                                  MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border 
                                    x:Name="DropDownBorder"
                                    Background="White"
                                    BorderThickness="1"
                                    BorderBrush="#DFDEDE"
                                    CornerRadius="3"    
                                        />
                                    <ScrollViewer Margin="0,5" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="95"/>
                            </Trigger>
                            <Trigger Property="IsGrouping" Value="true">
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </Trigger>
                            <Trigger Property="IsEditable" Value="true">
                                <Setter Property="IsTabStop" Value="false"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
            </Style.Triggers>
        </Style>
        <Style x:Key="{x:Type ComboBoxItem}" TargetType="{x:Type ComboBoxItem}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBoxItem}">
                        <Border 
                      Name="Border"
                      Padding="8,5"
                      SnapsToDevicePixels="true">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="Border" Property="Background" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Foreground" Value="#252424"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!---->
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#6264A7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="Height" Value="26" />
            <Setter Property="MinWidth" Value="100" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="SnapsToDevicePixels" Value="True" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border CornerRadius="3" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter Content="{TemplateBinding Content}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#464775" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#33344A" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#BDBDBD" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="Foreground" Value="#252424" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="2,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                        >
                            <ScrollViewer x:Name="PART_ContentHost" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#6264A7" />
                                <Setter Property="BorderThickness" Value="1" />
                            </Trigger>
                            <Trigger Property="IsReadOnly" Value="True">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type RadioButton}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type RadioButton}">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Grid Height="20" Width="20">
                                    <Border Name="Outer" Background="Transparent" BorderBrush="#DFDEDE" BorderThickness="2" CornerRadius="3" />
                                    <Border Name="Mark" CornerRadius="2" Margin="5" Visibility="Hidden" />
                                </Grid>
                            </BulletDecorator.Bullet>
                            <TextBlock Margin="5,0" Foreground="#252424" VerticalAlignment="Center">
                            <ContentPresenter />
                            </TextBlock>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Mark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="Outer" Property="BorderBrush" Value="#E2E2F6" />
                                <Setter TargetName="Mark" Property="Background" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Mark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="Outer" Property="BorderBrush" Value="#6264A7" />
                                <Setter TargetName="Mark" Property="Background" Value="#6264A7" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Outer" Property="BorderBrush" Value="#DFDEDE" />
                                <Setter TargetName="Mark" Property="Background" Value="#DFDEDE" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type CheckBox}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type CheckBox}">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Grid Height="24" Width="24">
                                    <Border Name="Outer" Background="Transparent" BorderBrush="#DFDEDE" BorderThickness="2" CornerRadius="3" Margin="2"/>
                                    <Border Name="CheckMark" Visibility="Hidden" >
                                        <Path
                                        x:Name="CheckMarkPath"
                                        Width="24" Height="24"
                                        SnapsToDevicePixels="False"
                                        StrokeThickness="2"
                                        Data="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" />
                                    </Border>
                                </Grid>
                            </BulletDecorator.Bullet>
                            <TextBlock Margin="5,0" Foreground="#252424" VerticalAlignment="Center">
                            <ContentPresenter />
                            </TextBlock>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#6264A7" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#DFDEDE" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto" />
                                <RowDefinition Height="*" />
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0">
                                <Label Foreground="#252424"
                                       FontSize="16">
                                    <ContentPresenter ContentSource="Header"/>
                                </Label>
                            </Border>
                            <Border Grid.Row="1"
                            BorderThickness="1"
                            BorderBrush="#DFDEDE"
                            CornerRadius="3"
                            >
                                <ContentPresenter />
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <StackPanel Orientation="Vertical">
        <StackPanel Orientation="Vertical" Margin="10,0">
            <GroupBox Header="Standard Team" Margin="0,5">
                <StackPanel Orientation="Vertical" Margin="10">
                    <RadioButton x:Name="standardTeamRadioButton" Content="Create a Standard Team" GroupName="teamType" IsChecked="True" Margin="0,0,0,10"/>
                    <Expander Header="General" IsExpanded="True">
                        <StackPanel Orientation="Vertical" Margin="5">
                            <StackPanel Orientation="Horizontal" Margin="0,2">
                                <TextBlock Text="Name:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                <TextBox x:Name="teamNameTextBox" Width="300" TextWrapping="Wrap" Margin="5,0"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,2">
                                <TextBlock Text="Description:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                                <TextBox x:Name="teamDescriptionTextBox" Width="300" TextWrapping="Wrap" Margin="5,0"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,2">
                                <TextBlock Text="Privacy:" Width="120" VerticalAlignment="Center" TextAlignment="Right"/>
                                <ComboBox x:Name="teamPrivacyComboBox" MinWidth="100" VerticalAlignment="Center" Margin="5,0"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" Margin="0,2">
                                
                                    <TextBlock Text="Owner:" Width="120" HorizontalAlignment="Left" TextAlignment="Right" Margin="0,5"/>
                                    <StackPanel Orientation="Vertical">
                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,2">
                                        <TextBlock Text="User (UPN):" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
                                        <TextBox x:Name="userTextBox" TextWrapping="Wrap" Margin="5,0" Width="350"/>
                                    </StackPanel>
                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                        <TextBlock Text="Name:" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="userDisplayNameTextBlock" TextWrapping="Wrap" Margin="5,0" Width="200"/>
                                    </StackPanel>
                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                        <TextBlock Text="Title:" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="userJobTitleTextBlock" TextWrapping="Wrap" Margin="5,0" Width="200"/>
                                    </StackPanel>
                                    <StackPanel Orientation="Horizontal" Margin="0,2">
                                        <TextBlock Text="Location:" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="userLocationTextBlock" TextWrapping="Wrap" Margin="5,0" Width="200"/>
                                    </StackPanel>
                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                        <Button x:Name="addTeamOwnerButton" Content="Add Owner" Margin="5,0" IsEnabled="False"/>
                                    </StackPanel>
                                    <DataGrid x:Name="teamOwnersDataGrid" AutoGenerateColumns="False" Margin="0,5" MinHeight="50" SelectionMode="Single" AlternatingRowBackground="#FFE5E5F1" AlternationCount="2" GridLinesVisibility="None">
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="Name" Binding="{Binding displayName}" IsReadOnly="True"/>
                                            <DataGridTextColumn Header="Title" Binding="{Binding jobTitle}" IsReadOnly="True" />
                                            <DataGridTextColumn Header="Location" Binding="{Binding officeLocation}" IsReadOnly="True" />
                                            <DataGridTextColumn Binding="{Binding id}" IsReadOnly="True" Visibility="Hidden"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                </StackPanel>
                                </StackPanel>
                                
                        </StackPanel>
                    </Expander>
                    <Expander Header="Additional Channels">
                        <StackPanel Orientation="Vertical" Margin="30,10">
                            <StackPanel Orientation="Horizontal">
                                <StackPanel Orientation="Vertical" Margin="5,0,0,0">
                                    <Label Content="Name:"/>
                                    <TextBox x:Name="teamChannelDisplayName1TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDisplayName2TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDisplayName3TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDisplayName4TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDisplayName5TextBox" Width="150" Margin="0,2" Height="22"/>
                                </StackPanel>
                                <StackPanel Orientation="Vertical" Margin="5,0,0,0">
                                    <Label Content="Description:"/>
                                    <TextBox x:Name="teamChannelDescription1TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDescription2TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDescription3TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDescription4TextBox" Width="150" Margin="0,2" Height="22"/>
                                    <TextBox x:Name="teamChannelDescription5TextBox" Width="150" Margin="0,2" Height="22"/>
                                </StackPanel>
                                <StackPanel Orientation="Vertical" Margin="5,0,0,0">
                                    <Label Content="Is Favourite:"/>
                                    <CheckBox x:Name="teamChannelFavourite1CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelFavourite2CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelFavourite3CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelFavourite4CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelFavourite5CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                </StackPanel>
                                <StackPanel Orientation="Vertical" Margin="5,0,0,0">
                                    <Label Content="Wiki:"/>
                                    <CheckBox x:Name="teamChannelWiki1CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelWiki2CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelWiki3CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelWiki4CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelWiki5CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                </StackPanel>
                                <StackPanel Orientation="Vertical" Margin="5,0,0,0">
                                    <Label Content="OneNote:"/>
                                    <CheckBox x:Name="teamChannelOneNote1CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelOneNote2CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelOneNote3CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelOneNote4CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                    <CheckBox x:Name="teamChannelOneNote5CheckBox" HorizontalAlignment="Center" Margin="0,5"/>
                                </StackPanel>
                            </StackPanel>
                        </StackPanel>
                    </Expander>
                    <Expander Header="Member Settings">
                        <StackPanel Orientation="Horizontal" Margin="30,10">
                            <StackPanel Orientation="Vertical">
                                <CheckBox x:Name="allowCreateUpdateChannelsCheckBox" Content="Create/Update Channels" IsChecked="True"/>
                                <CheckBox x:Name="allowDeleteChannelsCheckBox" Content="Delete Channels" IsChecked="True"/>
                                <CheckBox x:Name="allowAddRemoveAppsCheckBox" Content="Add/Remove Apps" IsChecked="True"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                <CheckBox x:Name="allowCreateUpdateRemoveTabsCheckBox" Content="Create/Update/Remove Tabs" IsChecked="True"/>
                                <CheckBox x:Name="allowCreateUpdateRemoveConnectorsCheckBox" Content="Create/Update/Remove Connectors" IsChecked="True"/>
                            </StackPanel>
                        </StackPanel>
                    </Expander>
                    <Expander Header="Guest Settings">
                        <StackPanel Orientation="Horizontal" Margin="30,10">
                            <StackPanel Orientation="Vertical">
                                <CheckBox x:Name="allowGuestCreateUpdateChannelsCheckBox" Content="Create/Update Channels"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                <CheckBox x:Name="allowGuestDeleteChannelsCheckBox" Content="Delete Channels"/>
                            </StackPanel>
                        </StackPanel>
                    </Expander>
                    <Expander Header="Messaging Settings">
                        <StackPanel Orientation="Horizontal" Margin="30,10">
                            <StackPanel Orientation="Vertical">
                                <CheckBox x:Name="allowUserEditMessagesCheckBox" Content="Members Edit Messages" IsChecked="True"/>
                                <CheckBox x:Name="allowUserDeleteMessagesCheckBox" Content="Members Delete Messages" IsChecked="True"/>
                                <CheckBox x:Name="allowOwnerDeleteMessagesCheckBox" Content="Owners Delete Messages" IsChecked="True"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                <CheckBox x:Name="allowTeamMentionsCheckBox" Content="Team Mentions" IsChecked="True"/>
                                <CheckBox x:Name="allowChannelMentionsCheckBox" Content="Channel Mentions" IsChecked="True"/>
                            </StackPanel>
                        </StackPanel>
                    </Expander>
                    <Expander Header="Fun Settings">
                        <StackPanel Orientation="Horizontal" Margin="30,10">
                            <StackPanel Orientation="Vertical">
                                <CheckBox x:Name="allowGiphyCheckBox" Content="Allow Giphy" IsChecked="True"/>
                                <StackPanel Orientation="Horizontal">
                                    <Label Content="Giphy Rating:"/>
                                    <ComboBox x:Name="giphyContentRatingComboBox" MinWidth="100" Margin="5,0,0,0"/>
                                </StackPanel>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                <CheckBox x:Name="allowStickersAndMemesCheckBox" Content="Allow Stickers and Memes" IsChecked="True"/>
                                <CheckBox x:Name="allowCustomMemesCheckBox" Content="Allow Custom Memes" IsChecked="True"/>
                            </StackPanel>
                        </StackPanel>
                    </Expander>
                    <Expander Header="Discovery Settings">
                        <StackPanel Orientation="Horizontal" Margin="30,10">
                            <StackPanel Orientation="Vertical">
                                <CheckBox x:Name="showInTeamsSearchAndSuggestionsCheckBox" Content="Show In Teams Search and Suggestions" IsChecked="True" IsEnabled="False"/>
                            </StackPanel>
                        </StackPanel>
                    </Expander>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="Custom Team" Margin="0,5">
                <StackPanel Orientation="Vertical" Margin="10">
                    <RadioButton x:Name="customTeamRadioButton" Content="Create a Custom Team using JSON" GroupName="teamType" IsChecked="False" Margin="0,0,0,10"/>
                    <TextBlock Text="Example JSON for creating Teams can be found at the Graph API documentation:" FontStyle="Italic" Margin="10,0"/>
                    <TextBlock Text="https://docs.microsoft.com/en-gb/graph/api/team-post?view=graph-rest-beta" Margin="10,0"/>
                    <Expander Header="JSON" Margin="0,10">
                        <TextBox x:Name="customJSONTextBox" MaxLines="30" MinLines="10" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" Margin="5"/>
                    </Expander>
                </StackPanel>
            </GroupBox>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="addTeamButton" Content="Add" Margin="10"/>
            <Button x:Name="cancelButton" Content="Cancel" Margin="10"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    # Feed XAML in to XMLNodeReader
    $XMLReader = (New-Object System.Xml.XmlNodeReader $xaml)

    # Create a Window Object
    $WindowObject = [Windows.Markup.XamlReader]::Load($XMLReader)

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {

        try {
            $window.Add($_.Name, $WindowObject.FindName($_.Name))
        }
        catch {

        }

    }

    # Populate UI
    # Privacy
    $privacyOptions = @("Private", "Public") | Sort-Object
    $window.teamPrivacyComboBox.ItemsSource = $privacyOptions
    $window.teamPrivacyComboBox.SelectedItem = "Public"
    # Giphy
    $giphyOptions = @("Strict", "Moderate") | Sort-Object
    $window.giphyContentRatingComboBox.ItemsSource = $giphyOptions
    $window.giphyContentRatingComboBox.SelectedItem = "Moderate"

    # Events
    # Cancel Clicked
    $window.cancelButton.add_Click( {

            # Close Window
            $window.AddTeamWindow.Close()

        })

    # Team Visibility Changed
    $window.teamPrivacyComboBox.add_SelectionChanged( {

            # Change discovery settings based on visibility
            if ($window.teamPrivacyComboBox.SelectedItem -eq "Public") {
            
                # Enable discovery
                $window.showInTeamsSearchAndSuggestionsCheckBox.IsChecked = $true
                # Disable Checkbox
                $window.showInTeamsSearchAndSuggestionsCheckBox.IsEnabled = $false

            }
            elseif ($window.teamPrivacyComboBox.SelectedItem -eq "Private") {

                # Disable discovery
                $window.showInTeamsSearchAndSuggestionsCheckBox.IsChecked = $false
                # Enable Checkbox
                $window.showInTeamsSearchAndSuggestionsCheckBox.IsEnabled = $true

            }

        })

    # User text changed
    $window.userTextBox.add_TextChanged( {

            $upn = $window.userTextBox.Text

            # Validate UPN
            if ($upn) {

                $validUPN = ValidateRegex $upn "([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)"

                if ($validUPN -eq $true) {

                    $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

                }

                # If user id is found
                if ($user.id) {

                    $window.userDisplayNameTextBlock.Text = $user.displayName
                    $window.userJobTitleTextBlock.Text = $user.jobTitle
                    $window.userLocationTextBlock.Text = $user.officeLocation
                    $window.addTeamOwnerButton.IsEnabled = $true

                }
                else {

                    $window.userDisplayNameTextBlock.Text = $null
                    $window.userJobTitleTextBlock.Text = $null
                    $window.userLocationTextBlock.Text = $null
                    $window.addTeamOwnerButton.IsEnabled = $false

                }

            }

        })

    # Add Clicked
    $window.addTeamButton.add_Click( {

            # Standard Team
            if ($window.standardTeamRadioButton.IsChecked -eq $true) {

                # Check the bare minumum is valid
                if ($window.teamNameTextBox.Text -and $window.teamPrivacyComboBox.SelectedItem) {

                    # Build Channel
                    $channels = @()
                    Foreach ($i in 1..5) {

                        # Current values
                        $currentTeamChannelName = "teamChannelDisplayName$($i)TextBox"
                        $currentTeamChannelDescription = "teamChannelDescription$($i)TextBox"
                        $currentTeamChannelFavourite = "teamChannelFavourite$($i)CheckBox"

                        # If channel has a name, let's use it
                        if ($window.$currentTeamChannelName.Text) {

                            $channel = @{

                                displayName         = $window.$currentTeamChannelName.Text
                                description         = $window.$currentTeamChannelDescription.Text
                                isFavoriteByDefault = $window.$currentTeamChannelFavourite.IsChecked

                            }

                            $channels += $channel

                        }
            
                    }

                    # Owners
                    $owners = @()
                    $window.teamOwnersDataGrid.ItemsSource | ForEach-Object {

                        $owners += "https://graph.microsoft.com/beta/users('$($_.id)')"

                    }

                    # Build base request body
                    $body = @{

                        # General
                        'template@odata.bind' = "https://graph.microsoft.com/beta/teamsTemplates('standard')"
                        displayName           = $window.teamNameTextBox.Text
                        description           = $window.teamDescriptionTextBox.Text
                        'owners@odata.bind'   = $owners
                        visibility            = $window.teamPrivacyComboBox.SelectedItem

                        # Member Settings
                        memberSettings        = @{

                            allowCreateUpdateChannels         = $window.allowCreateUpdateChannelsCheckBox.IsChecked
                            allowDeleteChannels               = $window.allowDeleteChannelsCheckBox.IsChecked
                            allowAddRemoveApps                = $window.allowAddRemoveAppsCheckBox.IsChecked
                            allowCreateUpdateRemoveTabs       = $window.allowCreateUpdateRemoveTabsCheckBox.IsChecked
                            allowCreateUpdateRemoveConnectors = $window.allowCreateUpdateRemoveConnectorsCheckBox.IsChecked

                        }

                        # Guest Settings
                        guestSettings         = @{

                            allowCreateUpdateChannels = $window.allowGuestCreateUpdateChannelsCheckBox.IsChecked
                            allowDeleteChannels       = $window.allowGuestDeleteChannelsCheckBox.IsChecked

                        }

                        # Messaging Settings
                        messagingSettings     = @{

                            allowUserEditMessages    = $window.allowUserEditMessagesCheckBox.IsChecked
                            allowUserDeleteMessages  = $window.allowUserDeleteMessagesCheckBox.IsChecked
                            allowOwnerDeleteMessages = $window.allowOwnerDeleteMessagesCheckBox.IsChecked
                            allowTeamMentions        = $window.allowTeamMentionsCheckBox.IsChecked
                            allowChannelMentions     = $window.allowChannelMentionsCheckBox.IsChecked

                        }

                        # Fun Settings
                        funSettings           = @{

                            allowGiphy            = $window.allowGiphyCheckBox.IsChecked
                            giphyContentRating    = $window.giphyContentRatingComboBox.SelectedItem
                            allowStickersAndMemes = $window.allowStickersAndMemesCheckBox.IsChecked
                            allowCustomMemes      = $window.allowCustomMemesCheckBox.IsChecked

                        }

                        # Discovery Settings
                        discoverySettings     = @{

                            showInTeamsSearchAndSuggestions = $window.showInTeamsSearchAndSuggestionsCheckBox.IsChecked

                        }

                        # Channels
                        channels              = $channels

                    }

                }
                else {

                    ErrorPrompt -messageBody "Not all fields are complete. Ensure the Team has a name and an owner." -messageTitle "Incomplete"

                }

                # Custom Team
            }
            elseif ($window.customTeamRadioButton.IsChecked -eq $true) {
                
                # If text in JSON text box
                if ($window.customJSONTextBox.Text) {

                    $body = $window.customJSONTextBox.Text | ConvertFrom-Json

                }
                else {

                    ErrorPrompt -messageBody "Not all fields are complete. Ensure you have added some JSON for the Team." -messageTitle "Incomplete"

                }

            }

            # If JSON body is available
            if ($body) {

                # Create Team
                InvokeGraphAPICall -method "POST" -uri "https://graph.microsoft.com/beta/teams" -body $body

                # Check if success
                if ($script:lastAPICallSuccess -eq $true) {
        
                    # Get newly created Team Id
                    $script:lastAPICallReponse.Headers.Location -match "\/teams\('([a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12})'\)\/operations\('([a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12})'\)"
                    if ($matches[1]) { $newTeamId = $matches[1] }

                    # If new Team Id
                    if ($newTeamId) {
                        
                        $script:currentTeamId = $newTeamId

                        # Get New Team/Group, may need to keep trying as there can be a small delay in creating a Team
                        while ([string]::IsNullOrEmpty($newTeam)) {

                            $newTeam = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId" -silent

                            Start-Sleep -Seconds 2

                        }

                        # Process Channel Tabs post Team creation e.g. Wiki, OneNote
                        Foreach ($i in 1..5) {

                            $requestedTeamChannelName = "teamChannelDisplayName$($i)TextBox"
                            $requestedTeamChannelWiki = "teamChannelWiki$($i)CheckBox"
                            $requestedTeamChannelOneNote = "teamChannelOneNote$($i)CheckBox"
                            $script:currentChannelId = $null
                            $channel = $null

                            # If channel has a name, let's look for it in created Channels
                            if (-not [string]::IsNullOrEmpty($window.$requestedTeamChannelName.Text)) {

                                while ([string]::IsNullOrEmpty($channel.value.id)) {

                                    $channel = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels?`$filter=displayName eq '$($window.$requestedTeamChannelName.Text)'"
        
                                    Start-Sleep -Seconds 2
        
                                }

                                if ($channel.value.id) {

                                    # Set as current channel Id 
                                    $script:currentChannelId = $channel.value.id

                                    # Wiki Not Checked? - Delete It!
                                    if ($window.$requestedTeamChannelWiki.IsChecked -eq $false) {

                                        RemoveWikiTab

                                    }

                                    # OneNote Checked? - Create It!
                                    if ($window.$requestedTeamChannelOneNote.IsChecked -eq $true) {

                                        AddOneNoteTab -notebookName $window.$requestedTeamChannelName.Text -tabName "OneNote"

                                    }

                                }

                            }
            
                        }
                        

                        OKPrompt -messageBody "Team has been created." -messageTitle "Team Created"

                        # Close Window
                        $window.AddTeamWindow.Close()

                        ListTeams

                        GetTeamInformation

                    }
                    else {
                        
                        ErrorPrompt -messageBody "Team not created." -messageTitle "Team Creation Failed"

                    }
                }

            }

        })

    # Add Owner Clicked
    $window.addTeamOwnerButton.add_Click( {

            # Get ID of UPN in textbox
            $upn = $window.userTextBox.Text
            $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

            # If user is found
            if ($user.id -and $script:lastAPICallSuccess -eq $true) {

                $ownersDataGrid = @()

                <# Currently can only have one owner at time of creation, so cannot use!

                # Existing owners
                $existingOwners = $window.teamOwnersDataGrid.ItemsSource
                $existingOwners | ForEach-Object {

                    # If a valid owner with id and not the same as the 'new owner' id
                    if ($_.id -and $_.id -ne $user.id) {

                        $owner = @{

                            displayName       = $_.displayName
                            userPrincipalName = $_.userPrincipalName
                            jobTitle          = $_.jobTitle
                            officeLocation    = $_.officeLocation
                            id                = $_.id
        
                        }
        
                        $ownersDataGrid += New-Object PSObject -Property $owner

                    }
                }

                #>

                # New owner
                $owner = @{

                    displayName       = $user.displayName
                    userPrincipalName = $user.userPrincipalName
                    jobTitle          = $user.jobTitle
                    officeLocation    = $user.officeLocation
                    id                = $user.id

                }

                $ownersDataGrid += New-Object PSObject -Property $owner

                # Add new and existing owners
                $window.teamOwnersDataGrid.ItemsSource = $ownersDataGrid

                # 

            }
        })

    # Show Window
    $window.AddTeamWindow.ShowDialog()
}

function DeleteTeam {

    param (

    )

    # If there is a current team
    if (-not [string]::IsNullOrEmpty($script:currentTeamId)) {

        $teamName = $script:mainWindow.teamDisplayNameTextBlock.Text

        # Confirm deletion
        $buttonType = [System.Windows.MessageBoxButton]::YesNo
        $messageIcon = [System.Windows.MessageBoxImage]::Question
        $messageBody = "Are you sure you want to delete the Team '$teamName'?"
        $messageTitle = "Confirm Deletion"
 
        $result = [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

        # If yes
        if ($result -eq "Yes") {

            # Delete Team
            InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId"

            # Check if success
            if ($script:lastAPICallSuccess -eq $true) {
        
                OKPrompt -messageBody "Team has been deleted." -messageTitle "Team Deleted"

                # Clear current Team, Channel and Tab as it no longer exists
                $script:currentTeamId = $null
                $script:currentChannelId = $null
                $script:currentTabId = $null

                # Disable UI
                $script:mainWindow.teamTabControl.IsEnabled = $false

                # Reload Team List
                ListTeams            

            }
            
        }

    }

}

function UpdateTeam {

    param (


    )

    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and $script:mainWindow.teamDisplayNameTextBox.Text -and $script:mainWindow.teamDescriptionTextBox.Text -and $script:mainWindow.teamPrivacyComboBox.SelectedItem) {

        # Build request body for Group
        $groupBody = @{

            # General
            displayName = $script:mainWindow.teamDisplayNameTextBox.Text
            description = $script:mainWindow.teamDescriptionTextBox.Text
            visibility  = $script:mainWindow.teamPrivacyComboBox.SelectedItem

        }

        # Build request body for Team
        $teamBody = @{

            # Member Settings
            memberSettings    = @{

                allowCreateUpdateChannels         = $script:mainWindow.allowCreateUpdateChannelsCheckBox.IsChecked
                allowDeleteChannels               = $script:mainWindow.allowDeleteChannelsCheckBox.IsChecked
                allowAddRemoveApps                = $script:mainWindow.allowAddRemoveAppsCheckBox.IsChecked
                allowCreateUpdateRemoveTabs       = $script:mainWindow.allowCreateUpdateRemoveTabsCheckBox.IsChecked
                allowCreateUpdateRemoveConnectors = $script:mainWindow.allowCreateUpdateRemoveConnectorsCheckBox.IsChecked

            }

            # Guest Settings
            guestSettings     = @{

                allowCreateUpdateChannels = $script:mainWindow.allowGuestCreateUpdateChannelsCheckBox.IsChecked
                allowDeleteChannels       = $script:mainWindow.allowGuestDeleteChannelsCheckBox.IsChecked

            }

            # Messaging Settings
            messagingSettings = @{

                allowUserEditMessages    = $script:mainWindow.allowUserEditMessagesCheckBox.IsChecked
                allowUserDeleteMessages  = $script:mainWindow.allowUserDeleteMessagesCheckBox.IsChecked
                allowOwnerDeleteMessages = $script:mainWindow.allowOwnerDeleteMessagesCheckBox.IsChecked
                allowTeamMentions        = $script:mainWindow.allowTeamMentionsCheckBox.IsChecked
                allowChannelMentions     = $script:mainWindow.allowChannelMentionsCheckBox.IsChecked

            }

            # Fun Settings
            funSettings       = @{

                allowGiphy            = $script:mainWindow.allowGiphyCheckBox.IsChecked
                giphyContentRating    = $script:mainWindow.giphyContentRatingComboBox.SelectedItem
                allowStickersAndMemes = $script:mainWindow.allowStickersAndMemesCheckBox.IsChecked
                allowCustomMemes      = $script:mainWindow.allowCustomMemesCheckBox.IsChecked

            }

        }

        # If Team is set to Private, update discovery settings too
        if ($script:mainWindow.teamPrivacyComboBox.SelectedItem -eq "Private") {

            $discoverySettings = @{

                discoverySettings = @{

                    showInTeamsSearchAndSuggestions = $script:mainWindow.showInTeamsSearchAndSuggestionsCheckBox.IsChecked

                }                

            }

            $teamBody += $discoverySettings
            
        }

    }
    else {

        ErrorPrompt -messageBody "Not all fields are complete. Ensure the Team has a name and a description." -messageTitle "Incomplete"

    }

    # If JSON body is available
    if ($groupBody -and $teamBody) {

        # Update Group
        InvokeGraphAPICall -method "PATCH" -uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId" -body $groupBody

        # Pause for a second or two
        Start-Sleep -Seconds 2

        # Update Team
        InvokeGraphAPICall -method "PATCH" -uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId" -body $teamBody

        # Check if success
        if ($script:lastAPICallSuccess -eq $true) {
    
            OKPrompt -messageBody "Team has been updated." -messageTitle "Team Updated"

        }

        # Reload Team
        GetTeamInformation

        # Update List (as name may have changed)
        ListTeams

    }

}

function ArchiveTeam {

    param (

        [Parameter(mandatory = $true)][ValidateSet('archive', 'unarchive')][string]$task
        
    )

    if (-not [string]::IsNullOrEmpty($script:currentTeamId)) {

        InvokeGraphAPICall -method "POST" -uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/$task"

    }

    # Check if success
    if ($script:lastAPICallSuccess -eq $true) {

        OKPrompt -messageBody "Team has been $task`d." -messageTitle "Team $task`d"

    }

    # Reload Team
    GetTeamInformation
    
}

function CloneTeam {

    param (

    )

    # Load Window
    # Declare Objects
    $window = @{ }

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="CloneTeamWindow"
    Title="Clone Team" SizeToContent="WidthAndHeight" FontSize="13.5"
    >
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#6264A7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="Height" Value="26" />
            <Setter Property="MinWidth" Value="100" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="SnapsToDevicePixels" Value="True" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border CornerRadius="3" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter Content="{TemplateBinding Content}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#464775" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#33344A" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#BDBDBD" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="Foreground" Value="#252424" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="2,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                        >
                            <ScrollViewer x:Name="PART_ContentHost" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#6264A7" />
                                <Setter Property="BorderThickness" Value="1" />
                            </Trigger>
                            <Trigger Property="IsReadOnly" Value="True">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Inspired from https://stackoverflow.com/questions/22673483/how-do-i-create-a-flat-combo-box-using-wpf -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="{x:Type ToggleButton}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="32" />
                </Grid.ColumnDefinitions>
                <Border
                  x:Name="Border" 
                  Grid.ColumnSpan="2"
                  Height="30"
                  CornerRadius="3"
                  BorderBrush="#DFDEDE"
                  Background="White"
                  BorderThickness="1" />
                <Border 
                  Grid.Column="0"
                  CornerRadius="3,0,0,3" 
                  Margin="1"
                 
                  />
                <Path 
                  x:Name="Arrow"
                  Grid.Column="1"     
                  Fill="#6264A7"
                  HorizontalAlignment="Center"
                  VerticalAlignment="Center"
                  Data="M7.41,8.58L12,13.17L16.59,8.58L18,10L12,16L6,10L7.41,8.58Z"
                    Width="24"
                    Height="24"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter TargetName="Border" Property="Background" Value="#F3F2F1" />
                    <Setter TargetName="Border" Property="BorderBrush" Value="#DFDEDE" />
                    <Setter Property="Foreground" Value="#252424"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="#DFDEDE" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="{x:Type TextBox}">
            <Border x:Name="PART_ContentHost" Focusable="False" Background="{TemplateBinding Background}" />
        </ControlTemplate>
        <Style x:Key="{x:Type ComboBox}" TargetType="{x:Type ComboBox}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="true"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="MinHeight" Value="20"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBox}">
                        <Grid>
                            <ToggleButton 
                            Name="ToggleButton" 
                            Template="{StaticResource ComboBoxToggleButton}" 
                            Grid.Column="2" 
                            Focusable="false"
                            IsChecked="{Binding Path=IsDropDownOpen,Mode=TwoWay,RelativeSource={RelativeSource TemplatedParent}}"
                            ClickMode="Press">
                            </ToggleButton>
                            <ContentPresenter
                            Name="ContentSite"
                            IsHitTestVisible="False" 
                            Content="{TemplateBinding SelectionBoxItem}"
                            ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                            ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                            Margin="8,3,28,3"
                            VerticalAlignment="Center"
                            HorizontalAlignment="Left" />
                            <TextBox x:Name="PART_EditableTextBox"
                            Style="{x:Null}" 
                            Template="{StaticResource ComboBoxTextBox}" 
                            HorizontalAlignment="Left" 
                            VerticalAlignment="Center"
                            Focusable="True" 
                            Background="Transparent"
                            Visibility="Hidden"
                            />
                            <Popup 
                            Name="Popup"
                            Placement="Bottom"
                            IsOpen="{TemplateBinding IsDropDownOpen}"
                            AllowsTransparency="True" 
                            Focusable="False"
                            PopupAnimation="Slide">
                                <Grid 
                                  Name="DropDown"
                                  SnapsToDevicePixels="True"                
                                  MinWidth="{TemplateBinding ActualWidth}"
                                  MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border 
                                    x:Name="DropDownBorder"
                                    Background="White"
                                    BorderThickness="1"
                                    BorderBrush="#DFDEDE"
                                    CornerRadius="3"    
                                        />
                                    <ScrollViewer Margin="0,5" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="95"/>
                            </Trigger>
                            <Trigger Property="IsGrouping" Value="true">
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </Trigger>
                            <Trigger Property="IsEditable" Value="true">
                                <Setter Property="IsTabStop" Value="false"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
            </Style.Triggers>
        </Style>
        <Style x:Key="{x:Type ComboBoxItem}" TargetType="{x:Type ComboBoxItem}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBoxItem}">
                        <Border 
                      Name="Border"
                      Padding="8,5"
                      SnapsToDevicePixels="true">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="Border" Property="Background" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Foreground" Value="#252424"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!---->
                <Style TargetType="{x:Type CheckBox}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type CheckBox}">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Grid Height="24" Width="24">
                                    <Border Name="Outer" Background="Transparent" BorderBrush="#DFDEDE" BorderThickness="2" CornerRadius="3" Margin="2"/>
                                    <Border Name="CheckMark" Visibility="Hidden" >
                                        <Path
                                        x:Name="CheckMarkPath"
                                        Width="24" Height="24"
                                        SnapsToDevicePixels="False"
                                        StrokeThickness="2"
                                        Data="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" />
                                    </Border>
                                </Grid>
                            </BulletDecorator.Bullet>
                            <TextBlock Margin="5,0" Foreground="#252424" VerticalAlignment="Center">
                            <ContentPresenter />
                            </TextBlock>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#6264A7" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#DFDEDE" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <StackPanel Orientation="Vertical" Margin="10,0">
        <StackPanel Orientation="Vertical" Margin="5">
            
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Name:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <TextBox x:Name="teamNameTextBox" TextWrapping="Wrap" Margin="5,0" Width="250"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Description:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <TextBox x:Name="teamDescriptionTextBox" TextWrapping="Wrap" Margin="5,0" Width="250"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Mail Nickname:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <TextBox x:Name="mailNicknameTextBox" TextWrapping="Wrap" Margin="5,0" Width="250"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Privacy:" Width="120" VerticalAlignment="Center" TextAlignment="Right"/>
                <ComboBox x:Name="teamPrivacyComboBox" MinWidth="100" VerticalAlignment="Center" Margin="5,0"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Clone Apps:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <CheckBox x:Name="cloneAppsCheckBox" VerticalAlignment="Center" Margin="5,0" IsChecked="True"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Clone Tabs:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <CheckBox x:Name="cloneTabsCheckBox" VerticalAlignment="Center" Margin="5,0" IsChecked="True"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Clone Settings:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <CheckBox x:Name="cloneSettingsCheckBox" VerticalAlignment="Center" Margin="5,0" IsChecked="True"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Clone Channels:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <CheckBox x:Name="cloneChannelsCheckBox" VerticalAlignment="Center" Margin="5,0" IsChecked="True"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Clone Members:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <CheckBox x:Name="cloneMembersCheckBox" VerticalAlignment="Center" Margin="5,0" IsChecked="True"/>
            </StackPanel>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="cloneTeamButton" Content="Clone" Margin="10" Width="100"/>
            <Button x:Name="cancelButton" Content="Cancel" Margin="10" Width="100"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    # Feed XAML in to XMLNodeReader
    $XMLReader = (New-Object System.Xml.XmlNodeReader $xaml)

    # Create a Window Object
    $WindowObject = [Windows.Markup.XamlReader]::Load($XMLReader)

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {

        try {
            $window.Add($_.Name, $WindowObject.FindName($_.Name))
        }
        catch {

        }

    }

    # Get current Group info
    $group = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId"

    # Populate UI
    # Privacy
    $privacyOptions = @("Private", "Public") | Sort-Object
    $window.teamPrivacyComboBox.ItemsSource = $privacyOptions
    $window.teamPrivacyComboBox.SelectedItem = $group.visibility

    # Events
    # Cancel Clicked
    $window.cancelButton.add_Click( {

            # Close Window
            $window.CloneTeamWindow.Close()

        })

    # Clone Clicked
    $window.cloneTeamButton.add_Click( {

            # Check the bare minumum is valid
            if ($window.teamNameTextBox.Text -and $window.teamPrivacyComboBox.SelectedItem -and $window.mailNicknameTextBox.Text) {

                # Collect options selected
                $options = @()

                if ($window.cloneAppsCheckBox.IsChecked -eq $true) {
                    
                    $options += "apps"

                }

                if ($window.cloneTabsCheckBox.IsChecked -eq $true) {
                    
                    $options += "tabs"

                }

                if ($window.cloneSettingsCheckBox.IsChecked -eq $true) {
                    
                    $options += "settings"

                }

                if ($window.cloneChannelsCheckBox.IsChecked -eq $true) {
                    
                    $options += "channels"

                }

                if ($window.cloneMembersCheckBox.IsChecked -eq $true) {
                    
                    $options += "members"

                }

                # Build Team
                $body = @{

                    displayName  = $window.teamNameTextBox.Text
                    description  = $window.teamDescriptionTextBox.Text
                    mailNickname = $window.mailNicknameTextBox.Text # Known issue, graph ignores this value!
                    visibility   = $window.teamPrivacyComboBox.SelectedItem
                    partsToClone = $options -join ","

                }

            }
            else {

                ErrorPrompt -messageBody "Not all fields are complete. Ensure the Team has a name and a mail nickname." -messageTitle "Incomplete"

            }

            # If JSON body is available
            if ($body) {

                # If using user permissions, if not already a member, temporarily add logged in user to the group (as membership is required to clone a Team)
                if ($script:me.id) {

                    # Check if already a member or not
                    $alreadyMember = checkIfMember
                    $alreadyOwner = checkIfOwner
            
                }

                # Get members and owners of Team to be cloned as members and owners sometimes don't clone properly and we can add back in after cloning
                $members = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/members"
                $owners = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/owners"

                # Cloned Team ID
                $clonedTeamId = $script:currentTeamId

                # Clone Team
                InvokeGraphAPICall -method "POST" -uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/clone" -body $body

                # Get newly created Team Id
                $script:lastAPICallReponse.Headers.Location -match "\/teams\('([a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12})'\)\/operations\('([a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12})'\)"
                if ($matches[1]) { $newTeamId = $matches[1] }

                # If new Team Id
                if ($newTeamId) {

                    $script:currentTeamId = $newTeamId

                    # Get New Team/Group, may need to keep trying as there can be a small delay in creating a Team
                    while ([string]::IsNullOrEmpty($newTeam)) {

                        $newTeam = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId" -silent

                        Start-Sleep -Seconds 2

                    }

                    # Add owners and members back in to Team (just in case) if required
                    if ($window.cloneMembersCheckBox.IsChecked -eq $true) {
                        
                    
                        $members.value | ForEach-Object {

                            $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/users/$($_.id)" }

                            # Add User
                            InvokeGraphAPICall -method "POST" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/members/`$ref" -body $body -silent

                        }

                        $owners.value | ForEach-Object {

                            $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/users/$($_.id)" }

                            # Add User
                            InvokeGraphAPICall -method "POST" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/owners/`$ref" -body $body -silent

                        }

                    }
                    # If not a member before, remove
                    if ($alreadyMember -eq $false -and $script:me.id) {

                        # Remove User from cloned Team and new Team
                        InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$clonedTeamId/members/$($script:me.id)/`$ref" -silent
                        InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/members/$($script:me.id)/`$ref" -silent


                    }

                    # If not a owner before, remove
                    if ($alreadyOwner -eq $false -and $script:me.id) {

                        # Remove User from cloned Team and new Team
                        InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$clonedTeamId/owners/$($script:me.id)/`$ref" -silent
                        InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/owners/$($script:me.id)/`$ref" -silent

                    }

                    OKPrompt -messageBody "Team has been cloned." -messageTitle "Team Cloned"

                }

                # Close Window
                $window.CloneTeamWindow.Close()
        
                # Reload Teams List
                ListTeams

                GetTeamInformation

            }

        })

    # Show Window
    $window.CloneTeamWindow.ShowDialog()

}

function AddChannel {

    param (

    )

    # Load Window
    # Declare Objects
    $window = @{ }

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="AddChannelWindow"
    Title="Add Channel" SizeToContent="WidthAndHeight" FontSize="13.5"
    >
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#6264A7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="Height" Value="26" />
            <Setter Property="MinWidth" Value="100" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="SnapsToDevicePixels" Value="True" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border CornerRadius="3" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter Content="{TemplateBinding Content}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#464775" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#33344A" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#BDBDBD" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="Foreground" Value="#252424" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="2,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                        >
                            <ScrollViewer x:Name="PART_ContentHost" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#6264A7" />
                                <Setter Property="BorderThickness" Value="1" />
                            </Trigger>
                            <Trigger Property="IsReadOnly" Value="True">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type CheckBox}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type CheckBox}">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Grid Height="24" Width="24">
                                    <Border Name="Outer" Background="Transparent" BorderBrush="#DFDEDE" BorderThickness="2" CornerRadius="3" Margin="2"/>
                                    <Border Name="CheckMark" Visibility="Hidden" >
                                        <Path
                                        x:Name="CheckMarkPath"
                                        Width="24" Height="24"
                                        SnapsToDevicePixels="False"
                                        StrokeThickness="2"
                                        Data="M21,7L9,19L3.5,13.5L4.91,12.09L9,16.17L19.59,5.59L21,7Z" />
                                    </Border>
                                </Grid>
                            </BulletDecorator.Bullet>
                            <TextBlock Margin="5,0" Foreground="#252424" VerticalAlignment="Center">
                            <ContentPresenter />
                            </TextBlock>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible" />
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#6264A7" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="CheckMarkPath" Property="Stroke" Value="#DFDEDE" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <StackPanel Orientation="Vertical" Margin="10,0">
        <StackPanel Orientation="Vertical" Margin="5">
                    <StackPanel Orientation="Horizontal" Margin="0,2">
                        <TextBlock Text="Name:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                        <TextBox x:Name="channelNameTextBox" TextWrapping="Wrap" Margin="5,0" Width="250"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,2">
                        <TextBlock Text="Description:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                        <TextBox x:Name="channelDescriptionTextBox" TextWrapping="Wrap" Margin="5,0" Width="250"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,2">
                        <TextBlock Text="Is Favourite:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                        <CheckBox x:Name="channelFavouriteCheckBox" VerticalAlignment="Center" Margin="5,0"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,2">
                        <TextBlock Text="Wiki:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                        <CheckBox x:Name="channelWikiCheckBox" VerticalAlignment="Center" Margin="5,0"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,2">
                        <TextBlock Text="OneNote:" Width="120" TextAlignment="Right" VerticalAlignment="Center"/>
                <CheckBox x:Name="channelOneNoteCheckBox" VerticalAlignment="Center" Margin="5,0"/>
            </StackPanel>
        </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="addChannelButton" Content="Add" Margin="10" Width="100"/>
            <Button x:Name="cancelButton" Content="Cancel" Margin="10" Width="100"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    # Feed XAML in to XMLNodeReader
    $XMLReader = (New-Object System.Xml.XmlNodeReader $xaml)

    # Create a Window Object
    $WindowObject = [Windows.Markup.XamlReader]::Load($XMLReader)

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {

        try {
            $window.Add($_.Name, $WindowObject.FindName($_.Name))
        }
        catch {

        }
    }

    # Events
    # Cancel Clicked
    $window.cancelButton.add_Click( {

            # Close Window
            $window.AddChannelWindow.Close()

        })

    # Add Clicked
    $window.addChannelButton.add_Click( {

            # Check the bare minumum is valid
            if ($window.channelNameTextBox.Text) {

                # Build Channel
                $body = @{

                    displayName         = $window.channelNameTextBox.Text
                    description         = $window.channelDescriptionTextBox.Text
                    isFavoriteByDefault = $window.channelFavouriteCheckBox.IsChecked

                }

            }
            else {

                ErrorPrompt -messageBody "Not all fields are complete. Ensure the Channel has a name." -messageTitle "Incomplete"

            }

            # If JSON body is available
            if ($body) {

                # Create Channel
                $channel = InvokeGraphAPICall -method "POST" -uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels" -body $body

                # Check if success
                if ($script:lastAPICallSuccess -eq $true -and $channel.id) {

                    # Set current channel Id to newly created channel
                    $script:currentChannelId = $channel.id

                    # Wiki Tab Unchecked - Remove!
                    if ($window.channelWikiCheckBox.IsChecked -eq $false) {

                        RemoveWikiTab

                    }

                    # Create OneNote notebook and Tab once Channel has been created and returned
                    if ($window.channelOneNoteCheckBox.IsChecked -eq $true) {

                        AddOneNoteTab -notebookName $window.channelNameTextBox.Text -tabName "OneNote"

                    }

                    OKPrompt -messageBody "Channel has been created." -messageTitle "Channel Created"

                    # Close Window
                    $window.AddChannelWindow.Close()
        
                    # Refresh Team
                    GetTeamInformation

                }                

            }

        })

    # Show Window
    $window.AddChannelWindow.ShowDialog()
}

function DeleteChannel {

    param (

    )

    # If there is a current team and channel
    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and -not [string]::IsNullOrEmpty($script:currentChannelId)) {

        $channelName = $script:mainWindow.channelsComboBox.SelectedItem.DisplayName

        # Confirm deletion
        $buttonType = [System.Windows.MessageBoxButton]::YesNo
        $messageIcon = [System.Windows.MessageBoxImage]::Question
        $messageBody = "Are you sure you want to delete the Channel '$channelName'?"
        $messageTitle = "Confirm Deletion"
 
        $result = [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

        # If yes
        if ($result -eq "Yes") {

            # Delete Channel
            InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId"

            # Check if success
            if ($script:lastAPICallSuccess -eq $true) {
        
                OKPrompt -messageBody "Channel has been deleted." -messageTitle "Channel Deleted"

                # Clear current Channel as it no longer exists
                $script:currentChannelId = $null

                # Reload Team
                GetTeamInformation

            }        

        }

    }

}

function UpdateChannel {

    param (

    )

    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and -not [string]::IsNullOrEmpty($script:currentChannelId) -and $script:mainWindow.channelDisplayNameTextBox.Text) {

        # Build request body for Channel
        $body = @{

            description         = $script:mainWindow.channelDescriptionTextBox.Text
            isFavoriteByDefault = $script:mainWindow.channelIsFavouriteByDefault2CheckBox.IsChecked

        }

        # If Channel Name has changed, include it
        $channel = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId"

        if ($channel.displayName -ne $script:mainWindow.ChannelDisplayNameTextBox.Text) {

            $body.Add("displayName", $script:mainWindow.channelDisplayNameTextBox.Text)

        }

    }
    else {

        ErrorPrompt -messageBody "Not all fields are complete. Ensure the Channel has a name." -messageTitle "Incomplete"

    }

    # If JSON body is available
    if ($body) {

        # Update Channel
        InvokeGraphAPICall -method "PATCH" -uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/channels/$script:currentChannelId" -body $body

        # Check if success
        if ($script:lastAPICallSuccess -eq $true) {
    
            OKPrompt -messageBody "Channel has been updated." -messageTitle "Channel Updated"

        }

        # Reload Team
        GetTeamInformation

    }

}

function AddTab {

    param (

    )

    # Load Window
    # Declare Objects
    $window = @{ }

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="AddTabWindow"
    Title="Add Tab" SizeToContent="WidthAndHeight" FontSize="13.5"
    >
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#6264A7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="Height" Value="26" />
            <Setter Property="MinWidth" Value="100" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="SnapsToDevicePixels" Value="True" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border CornerRadius="3" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter Content="{TemplateBinding Content}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#464775" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#33344A" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#BDBDBD" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="Foreground" Value="#252424" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="2,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                        >
                            <ScrollViewer x:Name="PART_ContentHost" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#6264A7" />
                                <Setter Property="BorderThickness" Value="1" />
                            </Trigger>
                            <Trigger Property="IsReadOnly" Value="True">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Inspired from https://stackoverflow.com/questions/22673483/how-do-i-create-a-flat-combo-box-using-wpf -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="{x:Type ToggleButton}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="32" />
                </Grid.ColumnDefinitions>
                <Border
                  x:Name="Border" 
                  Grid.ColumnSpan="2"
                  Height="30"
                  CornerRadius="3"
                  BorderBrush="#DFDEDE"
                  Background="White"
                  BorderThickness="1" />
                <Border 
                  Grid.Column="0"
                  CornerRadius="3,0,0,3" 
                  Margin="1"
                 
                  />
                <Path 
                  x:Name="Arrow"
                  Grid.Column="1"     
                  Fill="#6264A7"
                  HorizontalAlignment="Center"
                  VerticalAlignment="Center"
                  Data="M7.41,8.58L12,13.17L16.59,8.58L18,10L12,16L6,10L7.41,8.58Z"
                    Width="24"
                    Height="24"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter TargetName="Border" Property="Background" Value="#F3F2F1" />
                    <Setter TargetName="Border" Property="BorderBrush" Value="#DFDEDE" />
                    <Setter Property="Foreground" Value="#252424"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="#DFDEDE" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="{x:Type TextBox}">
            <Border x:Name="PART_ContentHost" Focusable="False" Background="{TemplateBinding Background}" />
        </ControlTemplate>
        <Style x:Key="{x:Type ComboBox}" TargetType="{x:Type ComboBox}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="true"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="MinHeight" Value="20"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBox}">
                        <Grid>
                            <ToggleButton 
                            Name="ToggleButton" 
                            Template="{StaticResource ComboBoxToggleButton}" 
                            Grid.Column="2" 
                            Focusable="false"
                            IsChecked="{Binding Path=IsDropDownOpen,Mode=TwoWay,RelativeSource={RelativeSource TemplatedParent}}"
                            ClickMode="Press">
                            </ToggleButton>
                            <ContentPresenter
                            Name="ContentSite"
                            IsHitTestVisible="False" 
                            Content="{TemplateBinding SelectionBoxItem}"
                            ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                            ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                            Margin="8,3,28,3"
                            VerticalAlignment="Center"
                            HorizontalAlignment="Left" />
                            <TextBox x:Name="PART_EditableTextBox"
                            Style="{x:Null}" 
                            Template="{StaticResource ComboBoxTextBox}" 
                            HorizontalAlignment="Left" 
                            VerticalAlignment="Center"
                            Focusable="True" 
                            Background="Transparent"
                            Visibility="Hidden"
                            />
                            <Popup 
                            Name="Popup"
                            Placement="Bottom"
                            IsOpen="{TemplateBinding IsDropDownOpen}"
                            AllowsTransparency="True" 
                            Focusable="False"
                            PopupAnimation="Slide">
                                <Grid 
                                  Name="DropDown"
                                  SnapsToDevicePixels="True"                
                                  MinWidth="{TemplateBinding ActualWidth}"
                                  MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border 
                                    x:Name="DropDownBorder"
                                    Background="White"
                                    BorderThickness="1"
                                    BorderBrush="#DFDEDE"
                                    CornerRadius="3"    
                                        />
                                    <ScrollViewer Margin="0,5" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="95"/>
                            </Trigger>
                            <Trigger Property="IsGrouping" Value="true">
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </Trigger>
                            <Trigger Property="IsEditable" Value="true">
                                <Setter Property="IsTabStop" Value="false"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
            </Style.Triggers>
        </Style>
        <Style x:Key="{x:Type ComboBoxItem}" TargetType="{x:Type ComboBoxItem}">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ComboBoxItem}">
                        <Border 
                      Name="Border"
                      Padding="8,5"
                      SnapsToDevicePixels="true">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="true">
                                <Setter TargetName="Border" Property="Background" Value="#E2E2F6" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="false">
                                <Setter Property="Foreground" Value="#252424"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!---->
    </Window.Resources>
    <StackPanel Orientation="Vertical" Margin="10,0">
        <StackPanel Orientation="Vertical" Margin="5">
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Name:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                <TextBox x:Name="tabNameTextBox" TextWrapping="Wrap" Margin="5,0" Width="260"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Type:" Width="100" VerticalAlignment="Center" TextAlignment="Right"/>
                <ComboBox x:Name="tabTypeComboBox" MinWidth="100" VerticalAlignment="Center" Margin="5,0"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2">
                <TextBlock Text="Website Tab URL:" Width="100" TextAlignment="Right" VerticalAlignment="Center"/>
                <TextBox x:Name="tabURLTextBox" TextWrapping="Wrap" Margin="5,0" Width="260" IsEnabled="False"/>
            </StackPanel>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="addTabButton" Content="Add" Margin="10" Width="100"/>
            <Button x:Name="cancelButton" Content="Cancel" Margin="10" Width="100"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    # Feed XAML in to XMLNodeReader
    $XMLReader = (New-Object System.Xml.XmlNodeReader $xaml)

    # Create a Window Object
    $WindowObject = [Windows.Markup.XamlReader]::Load($XMLReader)

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {

        try {
            $window.Add($_.Name, $WindowObject.FindName($_.Name))
        }
        catch {

        }

    }

    # Populate UI
    # Tab Type
    $tabTypes = @("Wiki", "OneNote", "Website") | Sort-Object
    $window.tabTypeComboBox.ItemsSource = $tabTypes
    $window.tabTypeComboBox.SelectedItem = "Wiki"

    # Events
    # Cancel Clicked
    $window.cancelButton.add_Click( {

            # Close Window
            $window.AddTabWindow.Close()

        })

    # Add Clicked
    $window.addTabButton.add_Click( {

            $tab = $null

            # Check the bare minumum is valid
            if ($window.tabNameTextBox.Text -and $window.tabTypeComboBox.SelectedItem) {

                # Wiki type?
                if ($window.tabTypeComboBox.SelectedItem -eq "Wiki") {
                    
                    $tab = AddWikiTab -wikiName $window.tabNameTextBox.Text
                
                    # OneNote type?
                }
                elseif ($window.tabTypeComboBox.SelectedItem -eq "OneNote") {
                    
                    $tab = AddOneNoteTab -notebookName $window.tabNameTextBox.Text -tabName $window.tabNameTextBox.Text

                    # Website type?
                }
                elseif ($window.tabTypeComboBox.SelectedItem -eq "Website") {
                   
                    # Validate URL
                    $validURL = ValidateRegex $window.tabURLTextBox.Text "http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+"

                    if ($validURL -eq $true) {

                        # Add Website Tab
                        $tab = AddWebsiteTab -WebsiteName $window.tabNameTextBox.Text -WebsiteURL $window.tabURLTextBox.Text

                    }
                    else {

                        ErrorPrompt -messageBody "Not a valid URL, please correct." -messageTitle "Invalid URL"

                    }
                    
                }
                else {

                    ErrorPrompt -messageBody "No Tab Type has been selected." -messageTitle "No Tab Type selected"

                }

                if ($tab) {
            
                    OKPrompt -messageBody "Tab has been created." -messageTitle "Tab Created"   

                    # Close Window
                    $window.AddTabWindow.Close()
            
                    # Refresh Channel
                    GetChannelInformation

                }


            }
            else {

                ErrorPrompt -messageBody "Not all fields are complete. Ensure the Tab has a name and a type selected." -messageTitle "Incomplete"

            }

        })

    # Tab Type Selection Changed
    $window.tabTypeComboBox.add_SelectionChanged( {

            # Website type?
            if ($window.tabTypeComboBox.SelectedItem -eq "Website") {

                # Enable Website textbox
                $window.tabURLTextBox.IsEnabled = $true

            }
            else {

                # Disable Website textbox
                $window.tabURLTextBox.IsEnabled = $false
`
        
            }

        })

    # Show Window
    $window.AddTabWindow.ShowDialog()

}

function DeleteTab {

    param (

    )

    # If there is a current team, channel and tab
    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and -not [string]::IsNullOrEmpty($script:currentChannelId) -and -not [string]::IsNullOrEmpty($script:currentTabId)) {

        $tabName = $script:mainWindow.tabDisplayNameTextBox.Text

        # Confirm deletion
        $buttonType = [System.Windows.MessageBoxButton]::YesNo
        $messageIcon = [System.Windows.MessageBoxImage]::Question
        $messageBody = "Are you sure you want to delete the Tab '$tabName'?"
        $messageTitle = "Confirm Deletion"
 
        $result = [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

        # If yes
        if ($result -eq "Yes") {

            # Delete Tab
            InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs/$script:currentTabId"

            # Check if success
            if ($script:lastAPICallSuccess -eq $true) {
        
                OKPrompt -messageBody "Tab has been deleted." -messageTitle "Tab Deleted"

                # Clear current Tab as it no longer exists
                $script:currentTabId = $null

                # Refresh Channel
                GetChannelInformation

            }        

        }

    }

}

function UpdateTab {

    param (

    )

    # If there is a current team, channel and tab
    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and -not [string]::IsNullOrEmpty($script:currentChannelId) -and -not [string]::IsNullOrEmpty($script:currentTabId) -and $script:mainWindow.tabDisplayNameTextBox.Text) {

        # Build request body for Tab
        $body = @{

            displayName = $script:mainWindow.tabDisplayNameTextBox.Text

        }

    }
    else {

        ErrorPrompt -messageBody "Not all fields are complete. Ensure the Tab has a name." -messageTitle "Incomplete"

    }

    # If JSON body is available
    if ($body) {

        # Update Tab
        InvokeGraphAPICall -method "PATCH" -uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs/$script:currentTabId" -body $body

        # Check if success
        if ($script:lastAPICallSuccess -eq $true) {
    
            OKPrompt -messageBody "Tab has been updated." -messageTitle "Tab Updated"

        }

        # Refresh Channel
        GetChannelInformation

    }

}
function AddUserToTeam {
    param (

        [Parameter(mandatory = $true)][ValidateSet('member', 'owner')][string]$type

    )
    
    # Load Window
    # Declare Objects
    $window = @{ }

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="AddUserWindow"
    Title="Add User" SizeToContent="WidthAndHeight" FontSize="13.5"
    >
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#6264A7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="Height" Value="26" />
            <Setter Property="MinWidth" Value="100" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="SnapsToDevicePixels" Value="True" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border CornerRadius="3" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter Content="{TemplateBinding Content}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#464775" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#33344A" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#BDBDBD" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="Foreground" Value="#252424" />
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Foreground" Value="#252424" />
            <Setter Property="BorderBrush" Value="#DFDEDE" />
            <Setter Property="MinHeight" Value="30" />
            <Setter Property="Margin" Value="5,0" />
            <Setter Property="Padding" Value="2,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border CornerRadius="3" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                        >
                            <ScrollViewer x:Name="PART_ContentHost" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#6264A7" />
                                <Setter Property="BorderThickness" Value="1" />
                            </Trigger>
                            <Trigger Property="IsReadOnly" Value="True">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#F3F2F1" />
                                <Setter Property="Foreground" Value="#717070" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
<StackPanel Orientation="Vertical" Margin="10">
        <StackPanel Orientation="Horizontal" Margin="0,2">
            <TextBlock Text="User (UPN):" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
            <TextBox x:Name="userTextBox" TextWrapping="Wrap" Margin="5,0" Width="400"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,2">
            <TextBlock Text="Name:" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
            <TextBlock x:Name="userDisplayNameTextBlock" TextWrapping="Wrap" Margin="5,0" Width="200"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,2">
            <TextBlock Text="Title:" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
            <TextBlock x:Name="userJobTitleTextBlock" TextWrapping="Wrap" Margin="5,0" Width="200"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,2">
            <TextBlock Text="Location:" Width="70" TextAlignment="Right" VerticalAlignment="Center"/>
            <TextBlock x:Name="userLocationTextBlock" TextWrapping="Wrap" Margin="5,0" Width="200"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
        <Button x:Name="addUserAddButton" Content="Add" Margin="10" Width="100" IsEnabled="False"/>
        <Button x:Name="addUserCancelButton" Content="Cancel" Margin="10" Width="100"/>
    </StackPanel>
</StackPanel>
</Window>
"@

    # Feed XAML in to XMLNodeReader
    $XMLReader = (New-Object System.Xml.XmlNodeReader $xaml)

    # Create a Window Object
    $WindowObject = [Windows.Markup.XamlReader]::Load($XMLReader)

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {

        try {
            $window.Add($_.Name, $WindowObject.FindName($_.Name))
        }
        catch {

        }

    }

    # Events
    # User text changed
    $window.userTextBox.add_TextChanged( {

            $upn = $window.userTextBox.Text

            # Validate UPN
            if ($upn) {

                $validUPN = ValidateRegex $upn "([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)"

                if ($validUPN -eq $true) {

                    $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

                }

                # If user id is found
                if ($user.id) {

                    $window.userDisplayNameTextBlock.Text = $user.displayName
                    $window.userJobTitleTextBlock.Text = $user.jobTitle
                    $window.userLocationTextBlock.Text = $user.officeLocation
                    $window.addUserAddButton.IsEnabled = $true

                    return $user

                }
                else {

                    $window.userDisplayNameTextBlock.Text = $null
                    $window.userJobTitleTextBlock.Text = $null
                    $window.userLocationTextBlock.Text = $null
                    $window.addUserAddButton.IsEnabled = $false

                }

            }

        })

    # Add User Clicked
    $window.addUserAddButton.add_Click( {
    
            # Get ID of UPN in textbox
            $upn = $window.userTextBox.Text
            $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

            # If user is found
            if ($user.id -and $script:lastAPICallSuccess -eq $true) {

                $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/users/$($user.id)" }

                # Add User
                InvokeGraphAPICall -method "POST" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/$type`s/`$ref" -body $body

                # Check if success
                if ($script:lastAPICallSuccess -eq $true) {

                    OKPrompt -messageBody "$($user.displayname) has been added as a $type to the Team." -messageTitle "User Added To Team"

                    # Refresh Team
                    GetTeamInformation

                    # Close Window
                    $window.AddUserWindow.Close()

                } 

            }

        })

    # Cancel Clicked
    $window.addUserCancelButton.add_Click( {
    
            # Close Window
            $window.AddUserWindow.Close()


        })

    # Show Window
    $window.AddUserWindow.ShowDialog()

}

function RemoveUserFromTeam {

    param (

        [Parameter(mandatory = $true)][PSCustomObject]$user,
        [Parameter(mandatory = $true)][ValidateSet('member', 'owner')][string]$type

    )

    # Check a user is selected (otherwise, do nothing)
    if ($user.id) {

        # Confirm deletion
        $buttonType = [System.Windows.MessageBoxButton]::YesNo
        $messageIcon = [System.Windows.MessageBoxImage]::Question
        $messageBody = "Are you sure you want to remove '$($user.displayName)' as a $type of this Team?"
        $messageTitle = "Confirm Removal"
 
        $result = [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

        # If yes
        if ($result -eq "Yes") {

            # Remove user from Team
            InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/$type`s/$($user.id)/`$ref"

            # Reload Team
            GetTeamInformation

        }

    }

}

function AddWikiTab {
    param (
        
        [Parameter(mandatory = $true)][string]$wikiName

    )
    
    # If there is a current team
    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and -not [string]::IsNullOrEmpty($script:currentChannelId)) {

        $body = @{

            "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.wiki')"
            name                  = $wikiName

        }

        $tab = InvokeGraphAPICall -Method "POST" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs" -body $body

        return $tab

    }

}

function RemoveWikiTab {
    param (
        
    )

    # Get Wiki Tab Id
    $tab = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs?`$filter=displayName eq 'Wiki'"

    if ($tab.value.id) {

        $script:currentTabId = $tab.value.id

        # Remove autocreated Wiki
        InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs/$script:currentTabId"

    }
    
}

function AddOneNoteTab {

    param (

        [Parameter(mandatory = $true)][string]$notebookName,
        [Parameter(mandatory = $true)][string]$tabName

    )

    # If there is a current team
    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and -not [string]::IsNullOrEmpty($script:currentChannelId)) {

        $notebookBody = @{

            displayName = $notebookName

        }

        # If using user permissions, if not already a member, temporarily add logged in user to the group (as membership is required to create a group notebook)
        if ($script:me.id) {

            # Check if already a member or not
            $alreadyMember = checkIfMember
            
        }

        $notebook = InvokeGraphAPICall -Method "POST" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/onenote/notebooks" -body $notebookBody

        # Check notebook was created
        if ($script:lastAPICallSuccess -eq $true) {

            $randomGuid = New-Guid

            # URL Encode Parameters
            $oneNoteWebUrl = [System.Web.HttpUtility]::UrlEncode($notebook.links.oneNoteWebUrl.href)
            $notebookId = [System.Web.HttpUtility]::UrlEncode($notebook.id)

            $tabBody = @{

                "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('0d820ecd-def2-4297-adad-78056cde7c78')"
                displayName           = $tabName

                configuration         = @{

                    entityId   = "$randomGuid`_$($notebook.id)"
                    contentUrl = "https://www.onenote.com/teams/TabContent?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&notebookSource=New&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$script:currentTeamId%2Fnotes%2Fnotebooks%2F$notebookId&oneNoteWebUrl=$oneNoteWebUrl&notebookName=note&ui={locale}&tenantId={tid}"
                    removeUrl  = "https://www.onenote.com/teams/TabRemove?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&notebookSource=New&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$script:currentTeamId%2Fnotes%2Fnotebooks%2F$notebookId&oneNoteWebUrl=$oneNoteWebUrl&notebookName=note&ui={locale}&tenantId={tid}"
                    websiteUrl = "https://www.onenote.com/teams/TabRedirect?redirectUrl=$oneNoteWebUrl"

                }

            }

            $tab = InvokeGraphAPICall -Method "POST" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs" -body $tabBody

            # If not a member before, remove
            if ($alreadyMember -eq $false -and $script:me.id) {

                # Remove User
                InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/members/$($script:me.id)/`$ref"

            }

            return $tab

        }

    }

}

function AddWebsiteTab {

    param (

        [Parameter(mandatory = $true)][string]$WebsiteName,
        [Parameter(mandatory = $true)][string]$WebsiteURL

    )

    # If there is a current team
    if (-not [string]::IsNullOrEmpty($script:currentTeamId) -and -not [string]::IsNullOrEmpty($script:currentChannelId)) {

        $tabBody = @{

            "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.web')"
            displayName           = $WebsiteName

            configuration         = @{

                entityId   = $null
                contentUrl = $WebsiteURL
                removeUrl  = $null
                websiteUrl = $WebsiteURL

            }

        }

        $tab = InvokeGraphAPICall -Method "POST" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs" -body $tabBody

        return $tab

    }

}
function ErrorPrompt {
    param (
        
        [Parameter(mandatory = $true)][string]$messageBody,
        [Parameter(mandatory = $true)][string]$messageTitle

    )

    $buttonType = [System.Windows.MessageBoxButton]::OK
    $messageIcon = [System.Windows.MessageBoxImage]::Error

    Write-Warning $messageBody
    [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

}


function OKPrompt {
    param (
        
        [Parameter(mandatory = $true)][string]$messageBody,
        [Parameter(mandatory = $true)][string]$messageTitle

    )

    $buttonType = [System.Windows.MessageBoxButton]::OK
    $messageIcon = [System.Windows.MessageBoxImage]::Information

    [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

}

function checkIfOwner {

    param (


    )

    $alreadyOwner = $null

    $owners = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/owners"

    $owners.value | ForEach-Object {

        if ($_.id -eq $script:me.id) {
            
            # Already a owner
            $alreadyOwner = $true

        }

        # If not already a owner, temporarily add
        if ([string]::IsNullOrEmpty($alreadyOwner)) {
                
            $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/users/$($script:me.id)" }

            # Add User
            InvokeGraphAPICall -method "POST" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/owners/`$ref" -body $body -silent

            # Wait a little bit before proceeding
            Start-Sleep -Seconds 2

            return $false

        }
        else {
            
            return $true

        }

    }

}

function checkIfMember {

    param (


    )

    $alreadyMember = $null

    $members = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/members"

    $members.value | ForEach-Object {

        if ($_.id -eq $script:me.id) {
            
            # Already a member
            $alreadyMember = $true

        }

        # If not already a member, temporarily add
        if ([string]::IsNullOrEmpty($alreadyMember)) {
                
            $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/users/$($script:me.id)" }

            # Add User
            InvokeGraphAPICall -method "POST" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/members/`$ref" -body $body -silent

            # Wait a little bit before proceeding
            Start-Sleep -Seconds 2

            return $false

        }
        else {
            
            return $true

        }

    }

}

function TeamsReport {
    param (
        
    )

    # Inpiriation from @MattEllisUC script - https://gist.github.com/moreaboutmatt/9dc6868955b5146ad815b8e044610047

    $teamsReport = @()

    # Save dialog
    $saveAs = New-Object Microsoft.Win32.SaveFileDialog
    $saveAs.Filter = "CSV|*.csv"
    $saveAs.ShowDialog()

    # Warn can take a while
    OKPrompt -messageBody "Click OK to start compiling report - check PowerShell console for progress.`n`rThis may take some time, please be patient." -messageTitle "Please Wait"

    # Get all groups that are Teams
    $groups = InvokeGraphAPICall -method "GET" -uri "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')" -silent
    $totalGroups = (($groups).value).Count

    $groups.value | ForEach-Object {

        # Progress counter
        $counter++

        # Progress
        Write-Progress -Activity "Compiling Report" -Status "Processing Team $counter of $totalGroups" -CurrentOperation $_.displayName -PercentComplete (($counter / $totalGroups) * 100)

        # More info from Graph
        $team = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$($_.id)" -silent
        $members = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$($_.id)/members" -silent
        $owners = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$($_.id)/owners" -silent
        $channels = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$($_.id)/channels" -silent

        $teamsObject = @{

            # General
            Id                                = $_.id
            DisplayName                       = $_.displayName
            Description                       = $_.description
            Mail                              = $_.mail
            Visibility                        = $_.visibility
            CreatedDateTime                   = $_.createdDateTime
            ExpirationDateTime                = $_.expirationDateTime
            Archived                          = $team.isArchived

            # Member Settings
            AllowCreateUpdateChannels         = $team.memberSettings.allowCreateUpdateChannels
            AllowDeleteChannels               = $team.memberSettings.allowDeleteChannels
            AllowAddRemoveApps                = $team.memberSettings.allowAddRemoveApps
            AllowCreateUpdateRemoveTabs       = $team.memberSettings.allowCreateUpdateRemoveTabs
            AllowCreateUpdateRemoveConnectors = $team.memberSettings.allowCreateUpdateRemoveConnectors

            # Guest Settings
            AllowGuestCreateUpdateChannels    = $team.guestSettings.allowCreateUpdateChannels
            AllowGuestDeleteChannels          = $team.guestSettings.allowDeleteChannels

            # Messaging Settings
            AllowUserEditMessages             = $team.messagingSettings.allowUserEditMessages
            AllowUserDeleteMessages           = $team.messagingSettings.allowUserDeleteMessages
            AllowOwnerDeleteMessages          = $team.messagingSettings.allowOwnerDeleteMessages
            AllowTeamMentions                 = $team.messagingSettings.allowTeamMentions
            AllowChannelMentions              = $team.messagingSettings.allowChannelMentions

            # Fun Settings
            AllowGiphy                        = $team.funSettings.allowGiphy
            GiphyContentRating                = $team.funSettings.giphyContentRating
            AllowStickersAndMemes             = $team.funSettings.allowStickersAndMemes
            AllowCustomMemes                  = $team.funSettings.allowCustomMemes

            # Discovery Settings
            ShowInTeamsSearchAndSuggestions   = $team.discoverySettings.showInTeamsSearchAndSuggestions

            # Totals
            NumberChannels                    = (($channels).value).Count
            NumberOwners                      = (($owners).value).Count
            NumberMembers                     = (($members).value).Count

        }

        $teamsReport += New-Object PSObject -Property $teamsObject

    }

    # Progress Complete
    Write-Progress -Activity "Compiling Report" -Completed

    if ($teamsReport) {

        try {
            
            $teamsReport | Export-CSV -Path $saveAs.Filename -NoTypeInformation
            OKPrompt -messageBody "Report completed and saved at $($saveAs.Filename)." -messageTitle "Report Completed"

        }
        catch {

            ErrorPrompt -messageBody "Unable to save report at $($saveAs.Filename)." -messageTitle "Unable To Save Report"

        }

    }   

}

# Load MainWindow XAML
LoadMainWindow

# Populate UI
# Privacy
$privacyOptions = @("Private", "Public") | Sort-Object
$script:mainWindow.teamPrivacyComboBox.ItemsSource = $privacyOptions

# Giphy
$giphyOptions = @("Strict", "Moderate") | Sort-Object
$script:mainWindow.giphyContentRatingComboBox.ItemsSource = $giphyOptions

# WPF Events
# Connect Clicked
$script:mainWindow.connectButton.add_Click( {

        GetAuthToken

        # If there is an issued token
        if ($script:issuedToken.access_token) {

            if ($script:me.userPrincipalName) {

                $loggedInUser = "`n`rLogged in as: $($script:me.userPrincipalName)"

            }

            OKPrompt -messageBody "Connected successfully to MS Graph $($loggedInUser)" -messageTitle "Connected to MS Graph"

            # Change UI State
            $script:mainWindow.disconnectButton.IsEnabled = $true
            $script:mainWindow.connectButton.IsEnabled = $false
            $script:mainWindow.teamsTabItem.IsEnabled = $true
            $script:mainWindow.connectionSettingsStackPanel.IsEnabled = $false
        
            ListTeams

        }

    })

# Disonnect Clicked
$script:mainWindow.disconnectButton.add_Click( {

        # Change UI State
        $script:mainWindow.disconnectButton.IsEnabled = $false
        $script:mainWindow.connectButton.IsEnabled = $true
        $script:mainWindow.teamsTabItem.IsEnabled = $false
        $script:mainWindow.connectionSettingsStackPanel.IsEnabled = $true

    })

# Refresh button clicked
$script:mainWindow.teamsRefreshButton.add_Click( {
        
        ListTeams

        # Reset filter
        $script:mainWindow.teamsFilterTextBox.text = $null

    })

# Team Selection Changed
$script:mainWindow.teamsListBox.add_SelectionChanged( {

        GetTeamInformation

    })

# Filter Teams List when text entered in Search box
$script:mainWindow.teamsFilterTextBox.add_TextChanged( {

        FilterTeams

    })

# Add Team clicked
$script:mainWindow.addTeamButton.add_Click( {

        AddTeam

    })

# Delete Team clicked
$script:mainWindow.deleteTeamButton.add_Click( {

        DeleteTeam

    })

# Archive Team clicked
$script:mainWindow.archiveTeamButton.add_Click( {

        ArchiveTeam -task archive

    })

# Unarchive Team clicked
$script:mainWindow.unArchiveTeamButton.add_Click( {

        ArchiveTeam -task unarchive

    })

# Clone Team clicked
$script:mainWindow.cloneTeamButton.add_Click( {

        CloneTeam

    })

# Add Member clicked
$script:mainWindow.addTeamMemberButton.add_Click( {

        AddUserToTeam -type member

    })

# Add Owner clicked
$script:mainWindow.addTeamOwnerButton.add_Click( {

        AddUserToTeam -type owner

    })

# Remove Member clicked
$script:mainWindow.removeTeamMemberButton.add_Click( {
    
        # Get Current Selection
        $selectedItem = $script:mainWindow.teamMembersDataGrid.SelectedItem

        RemoveUserFromTeam -user $selectedItem -type member

    })

# Remove Owner clicked
$script:mainWindow.removeTeamOwnerButton.add_Click( {
    
        # Get Current Selection
        $selectedItem = $script:mainWindow.teamOwnersDataGrid.SelectedItem

        RemoveUserFromTeam -user $selectedItem -type owner

    })

# Update Team clicked
$script:mainWindow.updateTeamButton.add_Click( {

        UpdateTeam

    })

# Refresh Team clicked
$script:mainWindow.refreshTeamButton.add_Click( {

        GetTeamInformation

    })

# Channel Selection Changed
$script:mainWindow.channelsComboBox.add_SelectionChanged( {

        GetChannelInformation

    })

# Add Channel clicked
$script:mainWindow.addChannelButton.add_Click( {

        AddChannel

    })

# Delete Channel clicked
$script:mainWindow.deleteChannelButton.add_Click( {

        DeleteChannel

    })

# Update Channel clicked
$script:mainWindow.updateChannelButton.add_Click( {

        UpdateChannel

    })

# Tab Selection Changed
$script:mainWindow.channelTabsComboBox.add_SelectionChanged( {

        GetTabInformation

    })

# Add Tab
$script:mainWindow.addTabButton.add_Click( {

        AddTab

    })

# Delete Tab
$script:mainWindow.deleteTabButton.add_Click( {

        DeleteTab

    })

# Update Tab
$script:mainWindow.updateTabButton.add_Click( {

        UpdateTab

    })

# Custom Azure AD App Selected
$script:mainWindow.useCustomAzureADApplicationRadioButton.add_Click( {

        # Enable UI
        $script:mainWindow.customAzureADAppStackPanel.IsEnabled = $true

    })

# Shared Azure AD App Selected
$script:mainWindow.useSharedAzureADApplicationRadioButton.add_Click( {

        # Enable UI
        $script:mainWindow.customAzureADAppStackPanel.IsEnabled = $false

    })

# Report
$script:mainWindow.teamsReportButton.add_Click( {

        TeamsReport

    })

# Team Visibility Changed
$script:mainWindow.teamPrivacyComboBox.add_SelectionChanged( {

        # Change discovery settings based on visibility
        if ($script:mainWindow.teamPrivacyComboBox.SelectedItem -eq "Public") {
        
            # Enable discovery
            $script:mainWindow.showInTeamsSearchAndSuggestionsCheckBox.IsChecked = $true
            # Disable Checkbox
            $script:mainWindow.showInTeamsSearchAndSuggestionsCheckBox.IsEnabled = $false

        }
        elseif ($script:mainWindow.teamPrivacyComboBox.SelectedItem -eq "Private") {

            # Enable discovery
            $script:mainWindow.showInTeamsSearchAndSuggestionsCheckBox.IsChecked = $false
            # Enable Checkbox
            $script:mainWindow.showInTeamsSearchAndSuggestionsCheckBox.IsEnabled = $true

        }

    })

# Show WPF MainWindow
$script:mainWindow.MainWindow.ShowDialog() | Out-Null