<p align="center">

  <h3 align="center"Optimize-WsusServer</h3>

</p>

<!-- TABLE OF CONTENTS -->
## Table of Contents

* [About the Project](#about-the-project)
  * [Built With](#built-with)
* [Getting Started](#getting-started)
  * [Prerequisites](#prerequisites)
  * [Installation](#installation)
* [Usage](#usage)
* [License](#license)
* [Contact](#contact)
* [References](#references)

<!-- ABOUT THE PROJECT -->
## About The Project
Comprehensive Windows Server Update Services (WSUS) cleanup, optimization, maintenance, and configuration PowerShell script.

Free and open source: [MIT License](https://github.com/awarre/Optimize-WsusServer/blob/master/LICENSE)

####Features
* Deep cleaning search and removal of unnecessary updates and drives by product title and update title.
* IIS Configuration validation and optimization.
* Disable device driver syncronization and caching.
* WSUS integrated update and computer cleanup
* Microsoft best practice WSUS database optimization and re-indexing
* Creation of daily and weekly optimization scheduled tasks.

### Built With
This section should list any major frameworks that you built your project using. Leave any add-ons/plugins for the acknowledgements section. Here are a few examples.
* [Powershell](https://docs.microsoft.com/en-us/powershell/)

<!-- GETTING STARTED -->
## Getting Started

### Prerequisites
* PowerShell
* Windows Server Update Services (WSUS)

### Installation
1. Download [Optimize-WsusServer.ps1](https://github.com/awarre/Optimize-WsusServer/blob/master/Optimize-WsusServer.ps1)
2. From PowerShell run
```powershell
Optimize-WsusServer.ps1 -FirstRun
```
<!-- USAGE EXAMPLES -->
## Usage

```powershell
Optimize-WsusServer.ps1 -FirstRun
```
Presents a series of prompts for user to initiate all recommended first run optimization tasks.

```
Optimize-WsusServer.PS1 -DisableDrivers
```
Disable device driver syncronization and caching.

```
Optimize-WsusServer.PS1 -DeepClean
```
Searches through most likely categories for unneeded updates and drivers to free up massive amounts of storage and improve database responsiveness. Prompts user to approve removal before deletion.

```
Optimize-WsusServer.PS1 -CheckConfig
```
Validates current WSUS IIS configuration against recommended settings. Helps prevent frequent WSUS/IIS/SQL service crashes and the "RESET SERVER NODE" error.

```
Optimize-WsusServer.PS1 -OptimizeServer
```
Runs all of Microsoft's built-in WSUS cleanup processes.

```
Optimize-WsusServer.PS1 -OptimizeDatabase
```
Runs Microsoft's recommended SQL reindexing script.

```
Optimize-WsusServer.PS1 -InstallDailyTask
```
Creates a scheduled task to run the OptimizeServer function nightly.

```  
Optimize-WsusServer.PS1 -InstallWeeklyTask
```
Creates a scheduled task to run the OptimizeDatabase function weekly.

<!-- LICENSE -->
## License
Distributed under the MIT License. See `LICENSE` for more information.

<!-- CONTACT -->
## Contact

Project Link: [https://github.com/awarre/Optimize-WsusServer](https://github.com/awarre/Optimize-WsusServer)

<!-- REFERENCES -->
## References
* [The complete guide to Microsoft WSUS and Configuration Manager SUP maintenance](https://support.microsoft.com/en-us/help/4490644/complete-guide-to-microsoft-wsus-and-configuration-manager-sup-maint)
* [Reindex the WSUS Database](https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/dd939795(v=ws.10))
* [10 Tips for the SQL Server PowerShell Scripter](https://devblogs.microsoft.com/scripting/10-tips-for-the-sql-server-powershell-scripter/)
* [How to Check if an Index Exists on a Table in SQL Server](https://littlekendra.com/2016/01/28/how-to-check-if-an-index-exists-on-a-table-in-sql-server/)
* [Getting 2016 updates to work on WSUS](https://www.reddit.com/r/sysadmin/comments/996xul/getting_2016_updates_to_work_on_wsus/)
* [Examples of Comment-Based Help](https://docs.microsoft.com/en-us/powershell/scripting/developer/help/examples-of-comment-based-help?view=powershell-7)
* [Get-IISSite](https://docs.microsoft.com/en-us/powershell/module/iisadministration/get-iissite?view=win10-ps)
* [Get-WebApplication](https://docs.microsoft.com/en-us/powershell/module/webadminstration/get-webapplication?view=winserver2012-ps)
* [Get-WsusClassification](https://docs.microsoft.com/en-us/powershell/module/wsus/get-wsusclassification?view=win10-ps)
* [Get-WsusProduct](https://docs.microsoft.com/en-us/powershell/module/wsus/get-wsusproduct?view=win10-ps)
* [Invoke-Sqlcmd](https://docs.microsoft.com/en-us/powershell/module/sqlserver/invoke-sqlcmd?view=sqlserver-ps)
* [Invoke-WsusServerCleanup](https://docs.microsoft.com/en-us/powershell/module/wsus/Invoke-WsusServerCleanup?view=win10-ps)
* [ScheduledTasks](https://docs.microsoft.com/en-us/powershell/module/scheduledtasks/?view=win10-ps)
* [Set-WsusClassification](https://docs.microsoft.com/en-us/powershell/module/updateservices/set-wsusclassification?view=win10-ps)
* [WSUS GetUpdates Method](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/aa350127(v=vs.85))
* [WSUS IUpdate Properties](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms752741(v=vs.85))