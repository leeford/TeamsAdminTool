# Add required assemblies
Add-Type -AssemblyName System.Web, PresentationFramework, PresentationCore

function LoadMainWindow {

    param (
        
    )

    # Declare Objects
    $script:mainWindow = @{}

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="MainWindow"
    Title="TeamsAdminTool" Width="1000" SizeToContent="WidthAndHeight" ResizeMode="CanMinimize" 
>
<StackPanel>
    <TabControl x:Name="mainTabControl" Margin="3">
        <TabItem x:Name="connectTabItem" Header="Connect" Width="100" Height="30">
            <StackPanel Orientation="Vertical">
                <GroupBox Header="Azure AD Application" Margin="10,0" HorizontalAlignment="Left" Width="900">
                    <StackPanel Orientation="Horizontal">
                        <Label x:Name="clientIdLabel" Content="Application (Client) ID:" HorizontalAlignment="Right"/>
                        <TextBox x:Name="clientIdTextBox" Margin="4" Width="270" HorizontalAlignment="Left" Text=""/>
                        <Label x:Name="tenantIdLabel" Content="Directory (Tenant) ID:" HorizontalAlignment="Right"/>
                        <TextBox x:Name="tenantIdTextBox" Margin="4" Width="270" HorizontalAlignment="Left" Text=""/>
                    </StackPanel>
                </GroupBox>
                <GroupBox Header="Graph API Permissions" Margin="10,0" HorizontalAlignment="Left" Width="900">
                    <StackPanel Orientation="Horizontal">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <RadioButton x:Name="applicationPermissionsRadioButton" Content="Application Permissions" GroupName="Permissions" IsChecked="True"/>
                            <StackPanel Orientation="Vertical">
                                <StackPanel Orientation="Horizontal">
                                    <Label x:Name="clientSecretLabel" Content="Secret:"/>
                                    <PasswordBox x:Name="clientSecretPasswordBox" Margin="4" Width="270" />
                                </StackPanel>
                            </StackPanel>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" Margin="10">
                            <RadioButton x:Name="userPermissionsRadioButton" Content="(Delegated) User Permissions" GroupName="Permissions"/>
                            <StackPanel Orientation="Horizontal">
                                <Label x:Name="redirectUriLabel" Content="Redirect URI:"/>
                                <TextBox x:Name="redirectUriTextBox" Margin="4" Width="400" Text="https://login.microsoftonline.com/common/oauth2/nativeclient" />
                            </StackPanel>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>
                <Button x:Name="connectButton" Content="Connect to Graph" Margin="15" Width="120" Height="20" HorizontalAlignment="Left"/>
            </StackPanel>
        </TabItem>
        <TabItem x:Name="teamsTabItem" Header="Teams" IsEnabled="False" Width="100" Height="30">
            <StackPanel Orientation="Horizontal">
                <StackPanel Orientation="Vertical" Margin="10">
                    <StackPanel Orientation="Horizontal">
                        <Button x:Name="teamsRefreshButton" Content="Refresh Teams" Width="90" Height="30" HorizontalAlignment="Left" Margin="5"/>
                        <Button x:Name="createTeamButton" Content="Create Team" Width="90" Height="30" HorizontalAlignment="Left" Margin="5"/>
                    </StackPanel>
                    <TextBlock x:Name="totalTeamsTextBlock" Text="Teams (0):"/>
                    <TextBox x:Name="teamsFilterTextBox" Margin="0,5,0,0"/>
                    <ListBox x:Name="teamsListBox" Width="200" Height="Auto" HorizontalAlignment="Left" MaxHeight="500"/>
                </StackPanel>
                <StackPanel Orientation="Vertical" Margin="0,10,10,10">
                    <StackPanel Orientation="Horizontal">
                        <Rectangle Height="50" Width="50">
                            <Rectangle.Fill>
                                <ImageBrush ImageSource="" />
                            </Rectangle.Fill>
                        </Rectangle>
                        <StackPanel Orientation="Vertical" Margin="10">
                            <TextBlock x:Name="teamDisplayNameTextBlock" FontSize="24" />
                            <TextBlock x:Name="teamDescriptionTextBlock" FontSize="16" />
                        </StackPanel>
                    </StackPanel>
                    <TabControl x:Name="teamTabControl" Width="750" Height="600" IsEnabled="False">
                        <TabItem x:Name="teamOverviewTabItem" Header="Overview" Width="100" Height="30">
                            <StackPanel Orientation="Vertical">
                                <StackPanel Orientation="Horizontal" Margin="10">
                                    <StackPanel Orientation="Vertical" Margin="5">
                                        <TextBlock Text="Visibility:" HorizontalAlignment="Right" />
                                        <TextBlock Text="Creation Date:" HorizontalAlignment="Right" />
                                        <TextBlock Text="Expiration Date:" HorizontalAlignment="Right" />
                                        <TextBlock Text="Mail:" HorizontalAlignment="Right" />
                                        <TextBlock Text="Archived:" HorizontalAlignment="Right" />
                                        <TextBlock Text="Members:" HorizontalAlignment="Right" />
                                        <TextBlock Text="Owners:" HorizontalAlignment="Right" />
                                        <TextBlock Text="Channels:" HorizontalAlignment="Right" />
                                    </StackPanel>
                                    <StackPanel Orientation="Vertical" HorizontalAlignment="Right" Margin="5">
                                        <TextBlock x:Name="teamVisibilityTextBlock" HorizontalAlignment="Left" />
                                        <TextBlock x:Name="teamCreatedDateTextBlock" HorizontalAlignment="Left" />
                                        <TextBlock x:Name="teamExpirationDateTextBlock" HorizontalAlignment="Left" />
                                        <TextBlock x:Name="teamMailTextBlock" HorizontalAlignment="Left" />
                                        <TextBlock x:Name="teamArchivedTextBlock" HorizontalAlignment="Left" />
                                        <TextBlock x:Name="totalTeamMembersTextBlock" HorizontalAlignment="Left" />
                                        <TextBlock x:Name="totalTeamOwnersTextBlock" HorizontalAlignment="Left" />
                                        <TextBlock x:Name="totalTeamChannelsTextBlock" HorizontalAlignment="Left" />
                                    </StackPanel>
                                </StackPanel>
                            </StackPanel>
                        </TabItem>
                        <TabItem x:Name="teamMembersTabItem" Header="Members" Width="100" Height="30">
                            <StackPanel Orientation="Vertical" Margin="10">
                                <TextBlock x:Name="teamMembersTextBlock" Text="Members (0):"/>
                                <DataGrid x:Name="teamMembersDataGrid" AutoGenerateColumns="False" Margin="0,10" Height="190" SelectionMode="Single" AlternatingRowBackground="#FFE5E5F1" AlternationCount="2" GridLinesVisibility="None">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Name" Binding="{Binding displayName}" IsReadOnly="True"/>
                                        <DataGridTextColumn Header="Title" Binding="{Binding jobTitle}" IsReadOnly="True" />
                                        <DataGridTextColumn Header="User Name" Binding="{Binding userPrincipalName}" IsReadOnly="True" />
                                        <DataGridTextColumn Header="Location" Binding="{Binding officeLocation}" IsReadOnly="True" />
                                        <DataGridTextColumn Binding="{Binding id}" IsReadOnly="True" Visibility="Hidden"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                    <Button x:Name="addTeamMemberButton" Content="Add Member" Width="120" Height="30" Margin="5"/>
                                    <Button x:Name="removeTeamMemberButton" Content="Remove Member" Width="120" Height="30" Margin="5"/>
                                </StackPanel>
                                <TextBlock x:Name="teamOwnersTextBlock" Text="Owners (0):" Margin="0,5,0,0"/>
                                <DataGrid x:Name="teamOwnersDataGrid" AutoGenerateColumns="False" Margin="0,10" Height="190" SelectionMode="Single" AlternatingRowBackground="#FFE5E5F1" AlternationCount="2" GridLinesVisibility="None">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Name" Binding="{Binding displayName}" IsReadOnly="True"/>
                                        <DataGridTextColumn Header="Title" Binding="{Binding jobTitle}" IsReadOnly="True" />
                                        <DataGridTextColumn Header="User Name" Binding="{Binding userPrincipalName}" IsReadOnly="True" />
                                        <DataGridTextColumn Header="Location" Binding="{Binding officeLocation}" IsReadOnly="True" />
                                        <DataGridTextColumn Binding="{Binding id}" IsReadOnly="True" Visibility="Hidden"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                    <Button x:Name="addTeamOwnerButton" Content="Add Owner" Width="120" Height="30" Margin="5"/>
                                    <Button x:Name="removeTeamOwnerButton" Content="Remove Owner" Width="120" Height="30" Margin="5"/>
                                </StackPanel>
                            </StackPanel>
                        </TabItem>
                        <TabItem x:Name="teamChannelsTabItem" Header="Channels" Width="100" Height="30">
                            <StackPanel Orientation="Vertical" Margin="10">
                                <StackPanel Orientation="Horizontal">
                                    <StackPanel Orientation="Vertical">
                                        <TextBlock x:Name="teamChannelsTextBlock" Text="Channels (0):"/>
                                        <ListBox x:Name="channelsListBox" Width="150" HorizontalAlignment="Left" MaxHeight="500" Margin="0,10" />
                                    </StackPanel>
                                    <StackPanel Orientation="Vertical" Margin="10,0">
                                        <TextBlock x:Name="channelDisplayNameTextBlock" FontSize="18" />
                                        <TextBlock x:Name="channelDescriptionTextBlock" FontSize="14" />
                                        <TabControl x:Name="channelTabControl" Width="565" Height="490" IsEnabled="False" Margin="0,10,0,0">
                                            <TabItem x:Name="channelOverviewTabItem" Header="Overview" Width="100">
                                                <StackPanel Orientation="Horizontal" Margin="10">
                                                    <StackPanel Orientation="Vertical" Margin="5">
                                                        <TextBlock Text="Mail:" HorizontalAlignment="Right" />
                                                        <TextBlock Text="Tabs:" HorizontalAlignment="Right" />
                                                        <TextBlock Text="Is Favourite By Default:" HorizontalAlignment="Right" />
                                                    </StackPanel>
                                                    <StackPanel Orientation="Vertical" HorizontalAlignment="Right" Margin="5">
                                                        <TextBlock x:Name="channelMailTextBlock" HorizontalAlignment="Left" />
                                                        <TextBlock x:Name="totalChannelTabsTextBlock" HorizontalAlignment="Left" />
                                                        <CheckBox x:Name="channelIsFavouriteByDefault1CheckBox" HorizontalAlignment="Left" />
                                                    </StackPanel>
                                                </StackPanel>
                                            </TabItem>
                                            <TabItem x:Name="channelTabsTabItem" Header="Tabs" Width="100">
                                                <StackPanel Orientation="Vertical" Margin="10">
                                                    <TextBlock x:Name="channelTabsTextBlock" Text="Tabs (0):"/>
                                                    <ComboBox x:Name="channelTabsComboBox" Width="200" Margin="0,10" HorizontalAlignment="Left"/>
                                                    <StackPanel Orientation="Horizontal">
                                                    </StackPanel>
                                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                                        <Button x:Name="deleteChannelTabButton" Content="Delete Tab" Width="120" Height="30" Margin="5"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </TabItem>
                                            <TabItem x:Name="channelSettingsTabItem" Header="Settings" Width="100">
                                                <StackPanel Orientation="Vertical" Margin="10">
                                                    <GroupBox Header="General">
                                                        <StackPanel Orientation="Horizontal" Margin="10,10">
                                                            <StackPanel Orientation="Vertical">
                                                                <Label Content="Channel Name:" HorizontalAlignment="Right"/>
                                                                <Label Content="Channel Description:" HorizontalAlignment="Right"/>
                                                                <Label Content="Is Favourite By Default:" HorizontalAlignment="Right"/>
                                                            </StackPanel>
                                                            <StackPanel Orientation="Vertical" Margin="5,0">
                                                                <TextBox x:Name="channelDisplayNameTextBox" Width="300" Height="23" Margin="0,1.5"/>
                                                                <TextBox x:Name="channelDescriptionTextBox" Width="300" Height="23" Margin="0,1.5"/>
                                                                <CheckBox x:Name="channelIsFavouriteByDefault2CheckBox" MinWidth="100" HorizontalAlignment="Left" Margin="0,5"/>
                                                            </StackPanel>
                                                        </StackPanel>
                                                    </GroupBox>
                                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom">
                                                        <Button x:Name="deleteChannelButton" Content="Delete Channel" Width="120" Height="30" Margin="5"/>
                                                        <Button x:Name="updateChannelButton" Content="Update Channel" Margin="5" Width="120" Height="30"/>
                                                        <Button x:Name="refreshChannelButton" Content="Refresh" Margin="5" Width="120" Height="30"/>
                                                    </StackPanel>
                                                </StackPanel>
                                            </TabItem>
                                        </TabControl>
                                    </StackPanel>
                                </StackPanel>
                            </StackPanel>
                        </TabItem>
                        <TabItem x:Name="teamSettingsTabItem" Header="Settings" Width="100" Height="30" >
                            <StackPanel Orientation="Vertical" Margin="10">
                                <StackPanel x:Name="teamSettingsStackPanel" Orientation="Vertical">
                                    <GroupBox Header="General">
                                        <StackPanel Orientation="Horizontal" Margin="30,10">
                                            <StackPanel Orientation="Vertical">
                                                <Label Content="Team Name:" HorizontalAlignment="Right"/>
                                                <Label Content="Team Description:" HorizontalAlignment="Right"/>
                                                <Label Content="Privacy:" HorizontalAlignment="Right"/>
                                            </StackPanel>
                                            <StackPanel Orientation="Vertical" Margin="5,0">
                                                <TextBox x:Name="teamDisplayNameTextBox" Width="400" Height="23" Margin="0,1.5"/>
                                                <TextBox x:Name="teamDescriptionTextBox" Width="400" Height="23" Margin="0,1.5"/>
                                                <ComboBox x:Name="teamPrivacyComboBox" MinWidth="100" HorizontalAlignment="Left" Margin="0,1.5"/>
                                            </StackPanel>
                                        </StackPanel>
                                    </GroupBox>
                                    <GroupBox Header="Member Settings">
                                        <StackPanel Orientation="Horizontal" Margin="30,10">
                                            <StackPanel Orientation="Vertical">
                                                <CheckBox x:Name="allowCreateUpdateChannelsCheckBox" Content="Create/Update Channels"/>
                                                <CheckBox x:Name="allowDeleteChannelsCheckBox" Content="Delete Channels"/>
                                                <CheckBox x:Name="allowAddRemoveAppsCheckBox" Content="Add/Remove Apps"/>
                                            </StackPanel>
                                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                                <CheckBox x:Name="allowCreateUpdateRemoveTabsCheckBox" Content="Create/Update/Remove Tabs"/>
                                                <CheckBox x:Name="allowCreateUpdateRemoveConnectorsCheckBox" Content="Create/Update/Remove Connectors"/>
                                            </StackPanel>
                                        </StackPanel>
                                    </GroupBox>
                                    <GroupBox Header="Guest Settings">
                                        <StackPanel Orientation="Horizontal" Margin="30,10">
                                            <StackPanel Orientation="Vertical">
                                                <CheckBox x:Name="allowGuestCreateUpdateChannelsCheckBox" Content="Create/Update Channels"/>
                                            </StackPanel>
                                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                                <CheckBox x:Name="allowGuestDeleteChannelsCheckBox" Content="Delete Channels"/>
                                            </StackPanel>
                                        </StackPanel>
                                    </GroupBox>
                                    <GroupBox Header="Messaging Settings">
                                        <StackPanel Orientation="Horizontal" Margin="30,10">
                                            <StackPanel Orientation="Vertical">
                                                <CheckBox x:Name="allowUserEditMessagesCheckBox" Content="Members Edit Messages"/>
                                                <CheckBox x:Name="allowUserDeleteMessagesCheckBox" Content="Members Delete Messages"/>
                                                <CheckBox x:Name="allowOwnerDeleteMessagesCheckBox" Content="Owners Delete Messages"/>
                                            </StackPanel>
                                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                                <CheckBox x:Name="allowTeamMentionsCheckBox" Content="Team Mentions"/>
                                                <CheckBox x:Name="allowChannelMentionsCheckBox" Content="Channel Mentions"/>
                                            </StackPanel>
                                        </StackPanel>
                                    </GroupBox>
                                    <GroupBox Header="Fun Settings">
                                        <StackPanel Orientation="Horizontal" Margin="30,10">
                                            <StackPanel Orientation="Vertical">
                                                <CheckBox x:Name="allowGiphyCheckBox" Content="Allow Giphy"/>
                                                <StackPanel Orientation="Horizontal">
                                                    <Label Content="Giphy Rating:"/>
                                                    <ComboBox x:Name="giphyContentRatingComboBox" MinWidth="100" Margin="5,0,0,0"/>
                                                </StackPanel>
                                            </StackPanel>
                                            <StackPanel Orientation="Vertical" Margin="10,0,0,0">
                                                <CheckBox x:Name="allowStickersAndMemesCheckBox" Content="Allow Stickers and Memes"/>
                                                <CheckBox x:Name="allowCustomMemesCheckBox" Content="Allow Custom Memes"/>
                                            </StackPanel>
                                        </StackPanel>
                                    </GroupBox>

                                </StackPanel>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom">
                                    <Button x:Name="archiveTeamButton" Content="Archive Team" Width="120" Height="30" Margin="5"/>
                                    <Button x:Name="unArchiveTeamButton" Content="Un-Archive Team" Width="120" Height="30" Margin="5"/>
                                    <Button x:Name="deleteTeamButton" Content="Delete Team" Width="120" Height="30" Margin="5"/>
                                    <Button x:Name="updateTeamButton" Content="Update Team" Width="120" Height="30" Margin="5"/>
                                    <Button x:Name="refreshTeamButton" Content="Refresh" Width="120" Height="30" Margin="5"/>
                                </StackPanel>
                            </StackPanel>
                        </TabItem>
                    </TabControl>
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

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {

        $script:mainWindow.Add($_.Name, $WindowObject.FindName($_.Name))

    }

}
function GetAuthToken {

    param (

    )

    # User permissions checked
    if ($script:mainWindow.userPermissionsRadioButton.IsChecked -eq $true) {

        # Check all required fields are valid
        ValidateTextBox "clientIdTextBox" "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
        ValidateTextBox "tenantIdTextBox" "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
        ValidateTextBox "redirectUriTextBox" "(.+)"

        if ($script:inputs.clientIdTextBox -eq $true -and $script:inputs.tenantIdTextBox -eq $true -and $script:inputs.redirectUriTextBox -eq $true) {

            # Get Issued Token for User
            $script:issuedToken = GetAuthTokenUser

        }
        else {

            ErrorPrompt -messageBody "Not all fields populated to request token. Ensure Client ID, Tenant ID and Redirect URI are populated correcty." -messageTitle "Not all fields populated"

        }

    }

    # Application permissions selected
    if ($script:mainWindow.applicationPermissionsRadioButton.IsChecked -eq $true) {

        # Check all required fields are valid
        ValidateTextBox "clientIdTextBox" "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
        ValidateTextBox "tenantIdTextBox" "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
        ValidateTextBox "clientSecretPasswordBox" "(.+)"

        if ($script:inputs.clientIdTextBox -eq $true -and $script:inputs.tenantIdTextBox -eq $true -and $script:inputs.clientSecretPasswordBox -eq $true) {

            # Get Issued Token for Application
            $script:issuedToken = GetAuthTokenApplication

        }
        else {

            ErrorPrompt -messageBody "Not all fields populated to request token. Ensure Client ID, Tenant ID and Client Secret are populated correcty." -messageTitle "Not all fields populated"

        }
    }

    # If there is an issued token
    if ($script:issuedToken.access_token) {

        # Set the token timer
        $script:tokenTimer = Get-Date

    }

}

function GetAuthCodeUser {
    param (

    )

    $clientId = [string]$script:mainWindow.clientIdTextBox.Text
    $tenantId = [string]$script:mainWindow.tenantIdTextBox.Text

    # Random State
    $state = Get-Random

    # Encode scope to fit inside query string
    $script:scope = "Group.ReadWrite.All User.ReadWrite.All"
    $encodedScope = [System.Web.HttpUtility]::UrlEncode($script:scope)

    # Redirect URI (encode it to fit inside query string)
    $redirectUri = [System.Web.HttpUtility]::UrlEncode($script:mainWindow.redirectUriTextBox.Text)

    # Construct URI
    $uri = [uri]"https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize?client_id=$clientId&response_type=code&redirect_uri=$redirectUri&response_mode=query&scope=$encodedScope&state=$state&prompt=login"

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

    $clientId = [string]$script:mainWindow.clientIdTextBox.Text
    $tenantId = [string]$script:mainWindow.tenantIdTextBox.Text
    
    # Get Auth Code needed for Token
    $authCode = GetAuthCodeUser

    Write-Host $authCode[1]

    # Construct URI
    $uri = [uri]"https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id    = $clientId
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

    $clientId = [string]$script:mainWindow.clientIdTextBox.Text
    $tenantId = [string]$script:mainWindow.tenantIdTextBox.Text

    # Construct URI
    $uri = [uri]"https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $clientId
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
        [Parameter(mandatory = $false)][string]$body,
        [Parameter(mandatory = $false)][string]$contentType,
        [Parameter(mandatory = $false)][switch]$silent

    )

    # Calculate current token age
    $tokenAge = New-TimeSpan $script:tokenTimer (Get-Date)

    # Check token has not expired
    if ($tokenAge.TotalSeconds -gt 3500) {

        Write-Warning "Token Expired!"

        # If last token issued included a refresh token
        if ($script:issuedToken.refresh_token) {

            # Get new token using refresh token
            GetAuthTokenUserRefresh

            # Otherwise authenticate without
        }
        else {

            GetAuthToken

        }

    }

    # If no content type, use application/json
    if (-not $contentType) {

        $contentType = "application/json"

    }
        
    # Construct headers
    $Headers = @{"Authorization" = "Bearer $($script:issuedToken.access_token)"}

    $apiCall = try {

        # Assume success
        $script:lastAPICallSuccess = $true

        # If there is a body (and it's not a GET), use it
        if ($body -and $method -ne "GET") {

            Invoke-RestMethod -Method $method -Uri $uri -ContentType $contentType -Headers $Headers -Body $body -ErrorAction Stop

        }
        else {

            Invoke-RestMethod -Method $method -Uri $uri -ContentType $contentType -Headers $Headers -ErrorAction Stop

        }

    }
    catch [System.Net.WebException] {

        if (-not $silent) {

            ErrorPrompt -messageBody "Exception was caught: $($_.Exception.Message) URI $uri" -messageTitle "Error on Graph API call"
            $script:lastAPICallSuccess = $false

        }

    }

        return $apiCall

}
function ValidateTextBox {
    param (
        
        [Parameter(mandatory = $true)][string]$textbox,
        [Parameter(mandatory = $true)][string]$regex

    )
    
    # Remove existing validation from hashtable
    $script:inputs.Remove($textbox)

    if ($script:mainWindow.$textbox.Text -match $regex -or $script:mainWindow.$textbox.Password -match $regex) {

        $script:mainWindow.$textbox.BorderBrush = "Green"
        $script:inputs.Add($textbox, $true)

    }
    else {

        $script:mainWindow.$textbox.BorderBrush = "Red"
        $script:inputs.Add($textbox, $false)
        Write-Warning "$textbox is not validated."

    }

}

function GetTeamInformation {
    param (


    )

    if ($script:mainWindow.teamsListBox.SelectedValue) {

        $script:currentTeamId = $script:mainWindow.teamsListBox.SelectedValue

    }

    # Reset current channel ID and disable channel tab control (as Team may have changed)
    $script:currentChannelId = $null
    $script:mainWindow.channelTabControl.IsEnabled = $false

    if (-not [string]::IsNullOrEmpty($script:currentTeamId)) {

        # Team #
        ########
        $team = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId"
        $script:mainWindow.teamArchivedTextBlock.Text = $team.isArchived
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
        $script:mainWindow.giphyContentRatingComboBox.SelectedItem = (Get-culture).TextInfo.ToTitleCase($team.funSettings.giphyContentRating)
        $script:mainWindow.allowStickersAndMemesCheckBox.IsChecked = $team.funSettings.allowStickersAndMemes
        $script:mainWindow.allowCustomMemesCheckBox.IsChecked = $team.funSettings.allowCustomMemes

        # Group #
        #########
        $group = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId"
        $script:mainWindow.teamDisplayNameTextBlock.Text = $group.displayName
        $script:mainWindow.teamDescriptionTextBlock.Text = $group.description
        $script:mainWindow.teamVisibilityTextBlock.Text = $group.visibility
        $script:mainWindow.teamPrivacyComboBox.SelectedItem = (Get-culture).TextInfo.ToTitleCase($group.visibility)
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

        # Team Photo/Image #
        ####################
        #$teamPhoto = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/photo/`$value" -contentType "image/jpg"
        #$teamPhoto | Out-File -FilePath "team.jpg"
        #$script:mainWindow.teamPhoto.ImageSource = $teamPhoto

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
        $script:mainWindow.channelsListBox.DisplayMemberPath = "displayName"
        $script:mainWindow.channelsListBox.SelectedValuePath = "id"
        $script:mainWindow.channelsListBox.ItemsSource = $channels.value

        # Enable Tabs
        $script:mainWindow.teamTabControl.IsEnabled = $true

        # Set UI based on archive status
        # Archived
        if ($team.isArchived -eq $true) {
            
            $script:mainWindow.teamSettingsStackPanel.IsEnabled = $false
            $script:mainWindow.archiveTeamButton.IsEnabled = $false
            $script:mainWindow.updateTeamButton.IsEnabled = $false
            $script:mainWindow.unarchiveTeamButton.IsEnabled = $true

        # Not archived
        } elseif ($team.isArchived -eq $false) {
            
            $script:mainWindow.teamSettingsStackPanel.IsEnabled = $true
            $script:mainWindow.archiveTeamButton.IsEnabled = $true
            $script:mainWindow.updateTeamButton.IsEnabled = $true
            $script:mainWindow.unarchiveTeamButton.IsEnabled = $false

        }
    
    }
}

function GetChannelInformation {
    param (


    )

    if ($script:mainWindow.channelsListBox.SelectedValue) {

        $script:currentChannelId = $script:mainWindow.channelsListBox.SelectedValue

    }

    if (-not [string]::IsNullOrEmpty($script:currentChannelId)) {

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

        } else {

            $script:mainWindow.channelIsFavouriteByDefault1CheckBox.IsChecked = $false
            $script:mainWindow.channelIsFavouriteByDefault2CheckBox.IsChecked = $false

        }
        
        Write-Host $channel

        # Tabs #
        ########
        $tabs = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/channels/$script:currentChannelId/tabs"
        $totalTabs = (($tabs).value).Count
        $script:mainWindow.totalChannelTabsTextBlock.Text = $totalTabs
        $script:mainWindow.channelTabsTextBlock.Text = "Tabs ($totalTabs):"

        # Enable Tabs
        $script:mainWindow.channelTabControl.IsEnabled = $true
    
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
    $filteredTeams = $script:teams.value | Where-Object {$_.displayName -like "*$teamsFilterText*"} | Sort-Object -Property displayName
    
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
    $addTeamWindow = @{}

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="AddTeamWindow"
    Title="Create Team" Width="640" SizeToContent="Height"
    >
    <StackPanel Orientation="Vertical">
        <StackPanel Orientation="Vertical" Margin="10,0">
            <GroupBox Header="Standard Team" Margin="0,5">
                <StackPanel>
                    <RadioButton x:Name="standardTeamRadioButton" Content="Create a Standard Team" GroupName="teamType" IsChecked="True" Margin="0,10,0,0"/>
                    <GroupBox Header="General" Margin="10">
                        <StackPanel Orientation="Horizontal" Margin="5">
                            <StackPanel Orientation="Vertical">
                                <Label Content="Team Name:" HorizontalAlignment="Right"/>
                                <Label Content="Team Description:" HorizontalAlignment="Right"/>
                                <Label Content="Privacy:" HorizontalAlignment="Right"/>
                                <Label Content="Team Owner:" HorizontalAlignment="Right"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" Margin="5,0">
                                <TextBox x:Name="teamNameTextBox" Width="400" Height="23" HorizontalAlignment="Left" Margin="0,1.5"/>
                                <TextBox x:Name="teamDescriptionTextBox" Width="400" HorizontalAlignment="Left" Height="23" Margin="0,1.5"/>
                                <ComboBox x:Name="teamPrivacyComboBox" MinWidth="100" HorizontalAlignment="Left" Margin="0,1.5"/>
                                <StackPanel Orientation="Horizontal">
                                    <TextBox x:Name="teamOwnerTextBox" Width="250" Height="23" Margin="0,1.5"/>
                                    <TextBlock x:Name="teamOwnerDisplayNameTextBlock" Width="200" Margin="5,3" Text="User: "/>
                                </StackPanel>
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>
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
                                    <Label Content="Favourite:"/>
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
                </StackPanel>
            </GroupBox>
            <GroupBox Header="Custom Team" Margin="0,5">
                <StackPanel Orientation="Vertical" Margin="10">
                    <RadioButton x:Name="customTeamRadioButton" Content="Create a Custom Team using JSON" GroupName="teamType" IsChecked="False" Margin="0,10"/>
                    <TextBlock Text="Example JSON for creating Teams can be found at the Graph API documentation: https://docs.microsoft.com/en-gb/graph/api/team-post?view=graph-rest-beta" TextWrapping="Wrap" FontStyle="Italic" Margin="10,0"/>
                    <Expander Header="JSON" Margin="0,10">
                        <TextBox x:Name="customJSONTextBox" MaxLines="30" MinLines="10" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" Margin="5"/>
                    </Expander>
                </StackPanel>
            </GroupBox>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="addTeamButton" Content="Add" Margin="10" Width="100"/>
            <Button x:Name="cancelButton" Content="Cancel" Margin="10" Width="100"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    # Feed XAML in to XMLNodeReader
    $XMLReader = (New-Object System.Xml.XmlNodeReader $xaml)

    # Create a Window Object
    $WindowObject = [Windows.Markup.XamlReader]::Load($XMLReader)

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {

        $addTeamWindow.Add($_.Name, $WindowObject.FindName($_.Name))

    }

    # Populate UI
    # Privacy
    $privacyOptions = @("Private", "Public") | Sort-Object
    $addTeamWindow.teamPrivacyComboBox.ItemsSource = $privacyOptions
    $addTeamWindow.teamPrivacyComboBox.SelectedItem = "Private"
    # Giphy
    $giphyOptions = @("Strict", "Moderate") | Sort-Object
    $addTeamWindow.giphyContentRatingComboBox.ItemsSource = $giphyOptions
    $addTeamWindow.giphyContentRatingComboBox.SelectedItem = "Moderate"

    # Events
    # Cancel Clicked
    $addTeamWindow.cancelButton.add_Click( {

            # Close Window
            $addTeamWindow.AddTeamWindow.Close()

        })

    # Owner text changed
    $addTeamWindow.teamOwnerTextBox.add_TextChanged( {

            $upn = $addTeamWindow.teamOwnerTextBox.Text
            $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

            # If user is found
            if ($user) {

                $addTeamWindow.teamOwnerDisplayNameTextBlock.Text = "User: $($user.displayName)"

            }
            else {

                $addTeamWindow.teamOwnerDisplayNameTextBlock.Text = "User: "

            }

        })

    # Add Clicked
    $addTeamWindow.addTeamButton.add_Click( {

            # Standard Team
            if ($addTeamWindow.standardTeamRadioButton.IsChecked -eq $true) {

                # Check owner is valid
                $upn = $addTeamWindow.teamOwnerTextBox.Text
                $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

                # Check the bare minumum is valid
                if ($addTeamWindow.teamNameTextBox.Text -and $addTeamWindow.teamDescriptionTextBox.Text -and $addTeamWindow.teamPrivacyComboBox.SelectedItem -and $user.id) {

                    # Build Channel and Tabs
                    $channels = @()
                    Foreach ($i in 1..5) {

                        # Current values
                        $currentTeamChannelName = "teamChannelDisplayName$($i)TextBox"
                        $currentTeamChannelDescription = "teamChannelDescription$($i)TextBox"
                        $currentTeamFavourite = "teamChannelFavourite$($i)CheckBox"
                        $currentTeamWiki = "teamChannelWiki$($i)CheckBox"
                        $currentTeamOneNote = "teamChannelOneNote$($i)CheckBox"

                        # If channel has a name, let's use it
                        if ($addTeamWindow.$currentTeamChannelName.Text) {

                            $tabs = @()

                            Write-Host $addTeamWindow.$currentTeamWiki.IsChecked

                            # Wiki Tab
                            if ($addTeamWindow.$currentTeamWiki.IsChecked -eq $true) {

                                $wikiTab = @{

                                    "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.wiki')"
                                    name                  = "Setup Wiki"

                                }

                                $tabs += $wikiTab

                            }

                            $channel = @{

                                displayName         = $addTeamWindow.$currentTeamChannelName.Text
                                description         = $addTeamWindow.$currentTeamChannelDescription.Text
                                isFavoriteByDefault = $addTeamWindow.$currentTeamFavourite.IsChecked

                                tabs                = $tabs

                            }

                            $channels += $channel

                        }
            
                    }

                    # Build base request body
                    $body = @{

                        # General
                        'template@odata.bind' = "https://graph.microsoft.com/beta/teamsTemplates('standard')"
                        displayName           = $addTeamWindow.teamNameTextBox.Text
                        description           = $addTeamWindow.teamDescriptionTextBox.Text
                        'owners@odata.bind'   = @(
                            "https://graph.microsoft.com/beta/users('$($user.id)')"
                        )
                        visibility            = $addTeamWindow.teamPrivacyComboBox.SelectedItem

                        # Member Settings
                        memberSettings        = @{

                            allowCreateUpdateChannels         = $addTeamWindow.allowCreateUpdateChannelsCheckBox.IsChecked
                            allowDeleteChannels               = $addTeamWindow.allowDeleteChannelsCheckBox.IsChecked
                            allowAddRemoveApps                = $addTeamWindow.allowAddRemoveAppsCheckBox.IsChecked
                            allowCreateUpdateRemoveTabs       = $addTeamWindow.allowCreateUpdateRemoveTabsCheckBox.IsChecked
                            allowCreateUpdateRemoveConnectors = $addTeamWindow.allowCreateUpdateRemoveConnectorsCheckBox.IsChecked

                        }

                        # Guest Settings
                        guestSettings         = @{

                            allowCreateUpdateChannels = $addTeamWindow.allowGuestCreateUpdateChannelsCheckBox.IsChecked
                            allowDeleteChannels       = $addTeamWindow.allowGuestDeleteChannelsCheckBox.IsChecked

                        }

                        # Messaging Settings
                        messagingSettings     = @{

                            allowUserEditMessages    = $addTeamWindow.allowUserEditMessagesCheckBox.IsChecked
                            allowUserDeleteMessages  = $addTeamWindow.allowUserDeleteMessagesCheckBox.IsChecked
                            allowOwnerDeleteMessages = $addTeamWindow.allowOwnerDeleteMessagesCheckBox.IsChecked
                            allowTeamMentions        = $addTeamWindow.allowTeamMentionsCheckBox.IsChecked
                            allowChannelMentions     = $addTeamWindow.allowChannelMentionsCheckBox.IsChecked

                        }

                        # Fun Settings
                        funSettings           = @{

                            allowGiphy            = $addTeamWindow.allowGiphyCheckBox.IsChecked
                            giphyContentRating    = $addTeamWindow.giphyContentRatingComboBox.SelectedItem
                            allowStickersAndMemes = $addTeamWindow.allowStickersAndMemesCheckBox.IsChecked
                            allowCustomMemes      = $addTeamWindow.allowCustomMemesCheckBox.IsChecked

                        }

                        # Channels
                        channels              = $channels

                    }

                    # Convert body to JSON
                    $bodyJson = $body | ConvertTo-Json -Depth 5 | ForEach-Object { [regex]::Unescape($_) }

                }
                else {

                    ErrorPrompt -messageBody "Not all fields are complete. Ensure your Team has a name, a description and an owner." -messageTitle "Incomplete"

                }

                # Custom Team
            }
            elseif ($addTeamWindow.customTeamRadioButton.IsChecked -eq $true) {
                
                # If text in JSON text box
                if ($addTeamWindow.customJSONTextBox.Text) {

                    $bodyJson = $addTeamWindow.customJSONTextBox.Text

                }
                else {

                    ErrorPrompt -messageBody "Not all fields are complete. Ensure you have added some JSON for the Team." -messageTitle "Incomplete"

                }

            }

            # If JSON body is available
            if ($bodyJson) {

                # Create Team
                InvokeGraphAPICall -method "POST" -uri "https://graph.microsoft.com/beta/teams" -body $bodyJson

                # Check if success
                if($script:lastAPICallSuccess -eq $true) {

                    $buttonType = [System.Windows.MessageBoxButton]::OK
                    $messageIcon = [System.Windows.MessageBoxImage]::Information
                    $messageBody = "Team has been created."
                    $messageTitle = "Team Created"
            
                    [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

                }
                
                # Close Window
                $addTeamWindow.AddTeamWindow.Close()
        
                # Refresh Team list
                ListTeams

            }

        })

    # Show Window
    $addTeamWindow.AddTeamWindow.ShowDialog()
}

function DeleteTeam {

    param (

    )

    # If there is a current team
    if (-not [string]::IsNullOrEmpty($script:currentTeamId)) {

        # Get current Team name from group
        $group = InvokeGraphAPICall -Method "GET" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId"

        # Confirm deletion
        $buttonType = [System.Windows.MessageBoxButton]::YesNo
        $messageIcon = [System.Windows.MessageBoxImage]::Question
        $messageBody = "Are you sure you want to delete the Team '$($group.displayName)'?"
        $messageTitle = "Confirm Deletion"
 
        $result = [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

        # If yes
        if ($result -eq "Yes") {

            # Delete Team
            InvokeGraphAPICall -Method "DELETE" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId"

            # Check if success
            if($script:lastAPICallSuccess -eq $true) {

                $buttonType = [System.Windows.MessageBoxButton]::OK
                $messageIcon = [System.Windows.MessageBoxImage]::Information
                $messageBody = "Team has been deleted."
                $messageTitle = "Team Deleted"
        
                [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

            }

            # Clear current Team as it no longer exists
            $script:currentTeamId = $null

            # Disable UI
            $script:mainWindow.teamTabControl.IsEnabled = $false

            # Reload Team List
            ListTeams            

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
            displayName           = $script:mainWindow.teamDisplayNameTextBox.Text
            description           = $script:mainWindow.teamDescriptionTextBox.Text
            visibility            = $script:mainWindow.teamPrivacyComboBox.SelectedItem

        }

        # Build request body for Team
        $teamBody = @{

            # Member Settings
            memberSettings        = @{

                allowCreateUpdateChannels         = $script:mainWindow.allowCreateUpdateChannelsCheckBox.IsChecked
                allowDeleteChannels               = $script:mainWindow.allowDeleteChannelsCheckBox.IsChecked
                allowAddRemoveApps                = $script:mainWindow.allowAddRemoveAppsCheckBox.IsChecked
                allowCreateUpdateRemoveTabs       = $script:mainWindow.allowCreateUpdateRemoveTabsCheckBox.IsChecked
                allowCreateUpdateRemoveConnectors = $script:mainWindow.allowCreateUpdateRemoveConnectorsCheckBox.IsChecked

            }

            # Guest Settings
            guestSettings         = @{

                allowCreateUpdateChannels = $script:mainWindow.allowGuestCreateUpdateChannelsCheckBox.IsChecked
                allowDeleteChannels       = $script:mainWindow.allowGuestDeleteChannelsCheckBox.IsChecked

            }

            # Messaging Settings
            messagingSettings     = @{

                allowUserEditMessages    = $script:mainWindow.allowUserEditMessagesCheckBox.IsChecked
                allowUserDeleteMessages  = $script:mainWindow.allowUserDeleteMessagesCheckBox.IsChecked
                allowOwnerDeleteMessages = $script:mainWindow.allowOwnerDeleteMessagesCheckBox.IsChecked
                allowTeamMentions        = $script:mainWindow.allowTeamMentionsCheckBox.IsChecked
                allowChannelMentions     = $script:mainWindow.allowChannelMentionsCheckBox.IsChecked

            }

            # Fun Settings
            funSettings           = @{

                allowGiphy            = $script:mainWindow.allowGiphyCheckBox.IsChecked
                giphyContentRating    = $script:mainWindow.giphyContentRatingComboBox.SelectedItem
                allowStickersAndMemes = $script:mainWindow.allowStickersAndMemesCheckBox.IsChecked
                allowCustomMemes      = $script:mainWindow.allowCustomMemesCheckBox.IsChecked

            }

        }

        # Convert body to JSON
        $groupBodyJson = $groupBody | ConvertTo-Json -Depth 5 | ForEach-Object { [regex]::Unescape($_) }
        $teamBodyJson = $teamBody | ConvertTo-Json -Depth 5 | ForEach-Object { [regex]::Unescape($_) }

    }
    else {

        ErrorPrompt -messageBody "Not all fields are complete. Ensure your Team has a name and a description." -messageTitle "Incomplete"

    }

    # If JSON body is available
    if ($groupBodyJson -and $teamBodyJson) {

        # Update Group
        InvokeGraphAPICall -method "PATCH" -uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId" -body $groupBodyJson

        # Update Team
        InvokeGraphAPICall -method "PATCH" -uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId" -body $teamBodyJson

        # Check if success
        if($script:lastAPICallSuccess -eq $true) {

            $buttonType = [System.Windows.MessageBoxButton]::OK
            $messageIcon = [System.Windows.MessageBoxImage]::Information
            $messageBody = "Team has been updated."
            $messageTitle = "Team Updated"
    
            [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

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

    if(-not [string]::IsNullOrEmpty($script:currentTeamId)) {

        InvokeGraphAPICall -method "POST" -uri "https://graph.microsoft.com/beta/teams/$script:currentTeamId/$task"

    }

    # Check if success
    if($script:lastAPICallSuccess -eq $true) {

        $buttonType = [System.Windows.MessageBoxButton]::OK
        $messageIcon = [System.Windows.MessageBoxImage]::Information
        $messageBody = "Team has been $task`d."
        $messageTitle = "Team $task`d"

        [System.Windows.MessageBox]::Show($messageBody, $messageTitle, $buttonType, $messageIcon)

    }

    # Reload Team
    GetTeamInformation
    
}

function AddUserToTeam {
    param (

        [Parameter(mandatory = $true)][ValidateSet('member', 'owner')][string]$type

    )
    
    # Load Window
    # Declare Objects
    $addUserWindow = @{}

    # Load XAML
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="AddUserWindow"
    Title="Add User" Width="400" SizeToContent="WidthAndHeight"
    >
<StackPanel Orientation="Vertical" Margin="10">
    <StackPanel Orientation="Horizontal">
        <Label x:Name="addUserLabel" Content="User (UPN):"/>
        <TextBox x:Name="addUserTextBox" Width="300" Height="23"/>
    </StackPanel>
    <StackPanel Orientation="Horizontal" Margin="0,10">
        <StackPanel Orientation="Vertical">
            <TextBlock Text="Name:" HorizontalAlignment="Right"/>
            <TextBlock Text="Title:" HorizontalAlignment="Right"/>
            <TextBlock Text="Location:" HorizontalAlignment="Right"/>
        </StackPanel>
        <StackPanel Orientation="Vertical" Margin="5,0">
            <TextBlock x:Name="userDisplayNameTextBlock" HorizontalAlignment="Left"/>
            <TextBlock x:Name="userJobTitleTextBlock" HorizontalAlignment="Left"/>
            <TextBlock x:Name="userLocationTextBlock" HorizontalAlignment="Left"/>
        </StackPanel>
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

    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {

        $addUserWindow.Add($_.Name, $WindowObject.FindName($_.Name))

    }

    # Events
    # User Text Changed
    $addUserWindow.addUserTextBox.add_TextChanged( {

            $upn = $addUserWindow.addUserTextBox.Text
            $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

            # If user is found
            if ($user) {

                $addUserWindow.userDisplayNameTextBlock.Text = $user.displayName
                $addUserWindow.userJobTitleTextBlock.Text = $user.jobTitle
                $addUserWindow.userLocationTextBlock.Text = $user.officeLocation
                $addUserWindow.addUserAddButton.IsEnabled = $true

            }
            else {

                $addUserWindow.userDisplayNameTextBlock.Text = $null
                $addUserWindow.userJobTitleTextBlock.Text = $null
                $addUserWindow.userLocationTextBlock.Text = $null
                $addUserWindow.addUserAddButton.IsEnabled = $false

            }

        })

    # Add User Clicked
    $addUserWindow.addUserAddButton.add_Click( {
    
            # Get ID of UPN in textbox
            $upn = $addUserWindow.addUserTextBox.Text
            $user = InvokeGraphAPICall -method "GET" -Uri "https://graph.microsoft.com/beta/users/$upn" -silent

            # If user is found
            if ($user) {

                $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/users/$($user.id)" } | ConvertTo-Json

                # Add User
                InvokeGraphAPICall -method "POST" -Uri "https://graph.microsoft.com/beta/groups/$script:currentTeamId/$type`s/`$ref" -body $body

                # Refresh Team
                GetTeamInformation

                # Close Window
                $addUserWindow.AddUserWindow.Close()

            }

        })

    # Cancel Clicked
    $addUserWindow.addUserCancelButton.add_Click( {
    
            # Close Window
            $addUserWindow.AddUserWindow.Close()


        })

    # Show Window
    $addUserWindow.AddUserWindow.ShowDialog()

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

$script:inputs = @{}

# Load MainWindow XAML
LoadMainWindow

# Populate UI
# Privacy
$privacyOptions = @("Private", "Public") | Sort-Object
$script:mainWindow.teamPrivacyComboBox.ItemsSource = $privacyOptions
$script:mainWindow.teamPrivacyComboBox.SelectedItem = "Private"

# Giphy
$giphyOptions = @("Strict", "Moderate") | Sort-Object
$script:mainWindow.giphyContentRatingComboBox.ItemsSource = $giphyOptions
$script:mainWindow.giphyContentRatingComboBox.SelectedItem = "Moderate"

# WPF Events
# Connect Clicked
$script:mainWindow.connectButton.add_Click( {

        GetAuthToken

        # If there is an issued token
        if ($script:issuedToken.access_token) {

            # Enable Tabs
            $script:mainWindow.teamsTabItem.IsEnabled = $true
        
            # List Teams
            ListTeams

        }

    })

$script:mainWindow.teamsRefreshButton.add_Click( {
        
        # List Teams
        ListTeams

        # Reset filter
        $script:mainWindow.teamsFilterTextBox.text = $null

    })

# Team Selection Changes
$script:mainWindow.teamsListBox.add_SelectionChanged( {

        GetTeamInformation

    })

# Filter Teams List when text entered in Search box
$script:mainWindow.teamsFilterTextBox.add_TextChanged( {

        FilterTeams

    })

# Create Team clicked
$script:mainWindow.createTeamButton.add_Click( {

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

    # Update Team
    UpdateTeam
})

# Refresh Team clicked
$script:mainWindow.refreshTeamButton.add_Click( {

    # Reload Team
    GetTeamInformation

})

# Channel Selection Changes
$script:mainWindow.channelsListBox.add_SelectionChanged( {

    GetChannelInformation

})

# Show WPF MainWindow
$script:mainWindow.MainWindow.ShowDialog() | Out-Null