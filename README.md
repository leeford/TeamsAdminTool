# TeamsAdminTool

_Disclaimer: _

## What is it?
![TeamsAdminTool](/images/2019-05-14 21_36_41-TeamsAdminTool.png)

TeamsAdminTool is a tool that allows you to manage Teams (inc. settings), Channels, Tabs etc. from a single GUI interface. It is written in Windows PowerShell (relies on WPF so won't work with Core) and requires no additional modules to be installed. To use the tool, simply clone/download the _TeamsAdminTool.ps1_ file and run from PowerShell.


## What can it be used for?
TeamsAdminTool allows you to manage many aspects of a Team that cannot currently be done using Teams Admin Centre or Teams PowerShell Module. To best explain it here is a table of all the features:

| Feature                            | TeamsAdminTool | Teams Admin Centre | Teams PowerShell |
| ---------------------------------- | :------------: | :----------------: | :--------------: |
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
| **Create a Channel**               |      Yes       |        Yes         |       Yes        |
| _- With defined Tabs_              |      Yes       |         No         |        No        |
| **Modify a Channel**               |      Yes       |         No         |       Yes        |
| _- Rename Channel_                 |      Yes       |         No         |       Yes        |
| _- Modify Channel settings_        |      Yes       |         No         |        No        |
| **Delete a Channel**               |      Yes       |        Yes         |       Yes        |
| **Create a Tab**                   |      Yes       |         No         |        No        |
| _- OneNote Tab_                    |      Yes       |         No         |        No        |
| _- Wiki Tab_                       |      Yes       |         No         |        No        |
| _- Website Tab_                    |      Yes       |         No         |        No        |
| **Modify a Tab**                   |      Yes       |         No         |        No        |
| _- Rename Tab_                     |      Yes       |         No         |        No        |
| **Delete a Tab**                   |      Yes       |         No         |        No        |
<br />


## How is this done?
TeamsAdminTool is built using Microsoft's Graph API which from what I gather, what the Admin Centre and PowerShell Module use. **Currently, this tool is using Graph API beta endpoints** to acheive full functionality. The plan is to roll this back to v1.0 endpoints as functionality becomes available.

_Note: Whilst there may be bugs with using the beta endpoint, in testing most functions work as expected (see Known Issues section), although being a beta this could change. If you are unsure of impact of the tool making changes to your Teams whilst it used beta endpoints - It is strongly advised you use the tool on some test Teams or a test Tenant._

## Usage

1. Download or clone the repository:
        git clone https://github.com/leeford/TeamsAdminTool.git
2. Run the script from a PowerShell prompt
        .\TeamsAdminTool.ps1
