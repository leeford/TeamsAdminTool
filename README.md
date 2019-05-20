# TeamsAdminTool

_**Disclaimer:** This script is provided ‘as-is’ without any warranty or support. Use of this script is at your own risk and I accept no responsibility for any damage caused._

## What is it?
![TeamsAdminTool](https://www.lee-ford.co.uk/images/TeamsAdminTool.png)

TeamsAdminTool is a tool that allows you to manage Teams (inc. settings), Channels, Tabs etc. from a single GUI interface. It is written in Windows PowerShell (relies on WPF so won't work with Core) and requires no additional modules to be installed. To use the tool, simply clone/download the _TeamsAdminTool.ps1_ file and run from PowerShell.

## What can it be used for?
TeamsAdminTool allows you to manage many aspects of a Team that cannot currently be done using Teams Admin Centre or Teams PowerShell Module. To best explain it here is a table of all the features:

| Feature                            | TeamsAdminTool | Teams Admin Centre | Teams PowerShell |
| ---------------------------------- | :------------: | :----------------: | :--------------: |
| **View created Teams**             |      Yes       |        Yes         |       Yes        |
| **Create a Team**                  |      Yes       |        Yes         |       Yes        |
| _- Using a JSON Template_          |      Yes       |         No         |        No        |
| _- With defined Channels and Tabs_ |      Yes       |         No         |        No        |
| _- With defined settings_          |      Yes       |         No         |        No        |
| **Modify a Team**                  |      Yes       |        Yes         |       Yes        |
| _- Rename Team_                    |      Yes       |        Yes         |       Yes        |
| _- Modify Team membership_         |      Yes       |        Yes         |       Yes        |
| _- Modify Team Member settings_    |      Yes       |        Yes         |       Yes        |
| _- Modify Team Guest settings_     |      Yes       |         No         |       Yes        |
| _- Modify Team Messaging settings_ |      Yes       |     Partially      |       Yes        |
| _- Modify Team Fun settings_       |      Yes       |         No         |       Yes        |
| _- Modify Team Discovery settings_ |      Yes       |         No         |        No        |
| **Archive a Team**                 |      Yes       |        Yes         |        No        |
| **Delete a Team**                  |      Yes       |        Yes         |       Yes        |
| **Clone a Team**                   |      Yes       |         No         |        No        |
| **View created Channels**          |      Yes       |        Yes         |       Yes        |
| **Create a Channel**               |      Yes       |        Yes         |       Yes        |
| _- With defined Tabs_              |      Yes       |         No         |        No        |
| **Modify a Channel**               |      Yes       |         No         |       Yes        |
| _- Rename Channel_                 |      Yes       |         No         |       Yes        |
| _- Modify Channel settings_        |      Yes       |         No         |        No        |
| **Delete a Channel**               |      Yes       |        Yes         |       Yes        |
| **View created Tabs**              |      Yes       |         No         |        No        |
| **Create a Tab**                   |      Yes       |         No         |        No        |
| _- OneNote Tab_                    |      Yes       |         No         |        No        |
| _- Wiki Tab_                       |      Yes       |         No         |        No        |
| _- Website Tab_                    |      Yes       |         No         |        No        |
| **Modify a Tab**                   |      Yes       |         No         |        No        |
| _- Rename Tab_                     |      Yes       |         No         |        No        |
| **Delete a Tab**                   |      Yes       |         No         |        No        |
| **Teams Report**                   |      Yes       |         No         |        No        |

## How is this done?
TeamsAdminTool is built using Microsoft's Graph API which from what I gather, what the Admin Centre and PowerShell Module use. **Currently, this tool is using Graph API beta endpoints** to achieve full functionality. The plan is to roll this back to v1.0 endpoints as functionality becomes available.

_Note: Whilst there may be bugs with using the beta endpoint, in testing most functions work as expected (see Known Issues section), although being a beta this could change. If you are unsure of impact of the tool making changes to your Teams whilst it used beta endpoints - It is strongly advised you use the tool on some test Teams or a test Tenant._

## Pre-requisites
To connect to Graph, you will need to use an Azure AD v2.0 Application. The application requires that it is granted the following Graph API permissions (delegated user or application):

* Group.ReadWrite.All - Allows creation and modification of Groups/Teams
* Directory.Read.All - Allows read-only access to directory for Group/Team membership
* Notes.ReadWrite.All - Allows access to Notebooks _Note: When using user permissions to create Notebooks the user is temporarily added to the group to facilitate this_

There are two ways this can be achieved:

  1. **Connect using a shared, pre-configured Azure AD Application** - This is the easiest option and no set up is required, using a shared application all you will need to do is login using your O365 account and consent the application against your tenant (admin consent required):

    You can grant consent when signing in with the tool or by going to this URL: https://login.microsoftonline.com/common/adminconsent?client_id=6d84adaa-2a01-4f45-964b-180cbdbfd20d
     
      _Grant Consent Prompt:_

      ![](https://www.lee-ford.co.uk/images/TeamsAdminToolGrantConsent.png)

      _Application Consent Granted in Azure AD:_

      ![](https://www.lee-ford.co.uk/images/2019-05-15%2013_54_24-Enterprise%20applications%20-%20Microsoft%20Azure.png)

      _Note: This does method does not provide access to your tenant for anyone other than the users you grant it to in your own tenant. It is essentially a template of permissions._
      
  2. **Create your own Azure AD Application** - Using the permissions mentioned above, create an Azure AD application ensuring you:
      * Grant the permissions mentioned above as delegated user or application permissions
      * If using application permissions, create a secret
      * If using delegated user permissions, ensure the Redirect URI https://login.microsoftonline.com/common/oauth2/nativeclient is checked

## Usage

1. Download or clone the repository:
   
    ```git clone https://github.com/leeford/TeamsAdminTool.git```

2. Run the script from a PowerShell prompt:
    
    ```.\TeamsAdminTool.ps1```

3. Choose an Azure AD Application type and click **Connect**. If prompted, login using your O365 credentials
    ![](https://www.lee-ford.co.uk/images/ConnectTeamsAdminTool.png)

## Known Issues

* When cloning a Team, the mail nickname specified will not be used - This is a Graph bug
* When retrieving or setting the Channel "IsFavoriteByDefault" value, null is returned - This is a [Graph bug](https://github.com/microsoftgraph/microsoft-graph-docs/issues/4241)
* When creating a OneNote tab (including the Tab name), once a Teams user interfaces with the Tab for the first time, the Tab is renamed to 'note'. Unsure what (if anything) can be done about this as this occurs outside this tool.