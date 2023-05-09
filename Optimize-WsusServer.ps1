#Requires -modules SqlServer

<#
.SYNOPSIS
    Comprehensive Windows Server Update Services (WSUS) configuration and optimization script.
.DESCRIPTION
    Comprehensive Windows Server Update Services (WSUS) configuration and optimization script.
    Features:
        -Deep cleaning search and removal of unnecessary updates and drives by product title and update title.
        -IIS Configuration validation and optimization.
        -Disable device driver syncronization and caching.
        -WSUS integrated update and computer cleanup
        -Microsoft best practice WSUS database optimization and re-indexing
        -Creation of daily and weekly optimization scheduled tasks.

.PARAMETER FirstRun
    Presents a series of prompts for user to initiate all recommended first run optimization tasks. Additional parameters will be ignored, as they will be redundant.

.PARAMETER DeclineSupersededUpdates
Declines all updates that have been approved and are superseded by other updates. The update will only be declined if a superseding update has been approved.

.PARAMETER DeepClean
    Searches through most likely categories for unneeded updates and drivers to free up massive amounts of storage and improve database responsiveness. Prompts user to approve removal before deletion.

.PARAMETER DisableDrivers
    Disable device driver syncronization and caching.

.PARAMETER CheckConfig
    Validates current WSUS IIS configuration against recommended settings. Helps prevent frequent WSUS/IIS/SQL service crashes and the "RESET SERVER NODE" error.

.PARAMETER OptimizeServer
    Runs all of Microsoft's built-in WSUS cleanup processes.

.PARAMETER OptimizeDatabase
    Runs Microsoft's recommended SQL reindexing script.

.PARAMETER InstallDailyTask
    Creates a scheduled task to run the OptimizeServer function nightly.

.PARAMETER InstallWeeklyTask
    Creates a scheduled task to run the OptimizeDatabase function weekly.

.NOTES
  Version:        1.2.1
  Author:         Austin Warren
  Creation Date:  2020/07/31

.EXAMPLE
  Optimize-WsusServer.ps1 -FirstRun
  Optimize-WsusServer.ps1 -DeepClean
  Optimize-WsusServer.ps1 -InstallDailyTask -CheckConfig -OptimizeServer
#>


[CmdletBinding()]
param (
    [Parameter()]
    [switch]
    $FirstRun,
    [Parameter()]
    [switch]
    $DisableDrivers,
    [Parameter()]
    [switch]
    $DeepClean,
    [Parameter()]
    [switch]
    $CheckConfig,
    [Parameter()]
    [switch]
    $InstallDailyTask,
    [Parameter()]
    [switch]
    $InstallWeeklyTask,
    [Parameter()]
    [switch]
    $OptimizeServer,
    [Parameter()]
    [switch]
    $OptimizeDatabase,
    [switch]
    $DeclineSupersededUpdates
)
#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Recommended IIS settings: https://www.reddit.com/r/sysadmin/comments/996xul/getting_2016_updates_to_work_on_wsus/
$recommendedIISSettings = @{
    QueueLength              = 25000
    LoadBalancerCapabilities = 'TcpLevel'
    CpuResetInterval         = 15
    RecyclingMemory          = 0
    RecyclingPrivateMemory   = 0
    ClientMaxRequestLength   = 204800
    ClientExecutionTimeout   = 7200
}

<#
DeepClean

To find potentially unneeded updates:
    1. WSUS management console
    2. Updates > All Updates
    3. Approval: Approved, Status: No Status
    4. Look for unused products
    5. Add titles to respective arrays below

Get-WsusProduct - Lists all Microsoft WSUS product categories.
#>

# Common unneeded updates by ProductTitles
$unneededUpdatesbyProductTitles = @(
    "Forefront Identity Manager 2010",
    "Microsoft Lync Server 2010",
    "Microsoft Lync Server 2013",
    "Office 2003",
    "Office 2007",
    "Office 2010",
    "Office 2002/XP",
    "SQL Server 2000",
    "SQL Server 2005",
    "SQL Server 2008",
    "Virtual PC",
    "Windows 2000",
    "Windows 7",
    "Windows 8 Embedded",
    "Windows 8.1",
    "Windows 8",
    "Windows Server 2003 R2",
    "Windows Server 2003",
    "Windows Server 2008 R2",
    "Windows Server 2008",
    "Windows Ultimate Extras",
    "Windows Vista",
    "Windows XP Embedded",
    "Windows XP x64 Edition",
    "Windows XP"
)

# Common unneeded updates by Title
$unneededUpdatesbyTitle = @(
    "Internet Explorer 6",
    "Internet Explorer 7",
    "Internet Explorer 8",
    "Internet Explorer 9",
    "Language Interface Pack",
    "Windows 10 (consumer editions)",
    "Windows 10 Education",
    "Windows 10 Enterprise N",
    "Itanium",
    "ARM64"
)

<#
REFERENCES
    The complete guide to Microsoft WSUS and Configuration Manager SUP maintenance
    https://support.microsoft.com/en-us/help/4490644/complete-guide-to-microsoft-wsus-and-configuration-manager-sup-maint

    Invoke-WsusServerCleanup
    https://docs.microsoft.com/en-us/powershell/module/wsus/Invoke-WsusServerCleanup?view=win10-ps

    Reindex the WSUS Database
    https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/dd939795(v=ws.10)

    Invoke-Sqlcmd
    https://docs.microsoft.com/en-us/powershell/module/sqlserver/invoke-sqlcmd?view=sqlserver-ps

    How to Check if an Index Exists on a Table in SQL Server
    https://littlekendra.com/2016/01/28/how-to-check-if-an-index-exists-on-a-table-in-sql-server/
#>

<#
    "[U]sed to create custom indexes in the SUSDB database. This is a one-time process, which is optional but recommended, as doing so will greatly improve performance during subsequent cleanup operations."
    Modified to check if indexes already exist before creating them.
#>
$createCustomIndexesSQLQuery = @"
USE [SUSDB]
IF 0 = (SELECT COUNT(*) as index_count
    FROM sys.indexes
    WHERE object_id = OBJECT_ID('[dbo].[tbLocalizedPropertyForRevision]')
    AND name='nclLocalizedPropertyID')
BEGIN
-- Create custom index in tbLocalizedPropertyForRevision
	CREATE NONCLUSTERED INDEX [nclLocalizedPropertyID] ON [dbo].[tbLocalizedPropertyForRevision]
	(
		 [LocalizedPropertyID] ASC
	)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
END
ELSE
BEGIN
	PRINT '[nclLocalizedPropertyID] ON [dbo].[tbLocalizedPropertyForRevision] already exists'
END ;
GO
IF 0 = (SELECT COUNT(*) as index_count
    FROM sys.indexes
    WHERE object_id = OBJECT_ID('[dbo].[tbRevisionSupersedesUpdate]')
    AND name='nclSupercededUpdateID')
BEGIN
-- Create custom index in tbRevisionSupersedesUpdate
	CREATE NONCLUSTERED INDEX [nclSupercededUpdateID] ON [dbo].[tbRevisionSupersedesUpdate]
	(
		 [SupersededUpdateID] ASC
	)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
END
ELSE
BEGIN
	PRINT '[nclSupercededUpdateID] ON [dbo].[tbRevisionSupersedesUpdate] already exists'
END ;
GO
"@

<#
    Microsoft recommended database maintenance script

    "The performance of large Windows Server Update Services (WSUS) deployments will degrade over time if the WSUS database
    is not maintained properly. The WSUSDBMaintenance script is a T-SQL script that can be run by SQL Server administrators
    to re-index and defragment WSUS databases. It should not be used on WSUS 2.0 databases.This script contributed by the
    Microsoft WSUS team."

    Reference: https://support.microsoft.com/en-us/help/4490644/complete-guide-to-microsoft-wsus-and-configuration-manager-sup-maint
#>
$wsusDBMaintenanceSQLQuery = @"
/******************************************************************************
This sample T-SQL script performs basic maintenance tasks on SUSDB
1. Identifies indexes that are fragmented and defragments them. For certain
   tables, a fill-factor is set in order to improve insert performance.
   Based on MSDN sample at http://msdn2.microsoft.com/en-us/library/ms188917.aspx
   and tailored for SUSDB requirements
2. Updates potentially out-of-date table statistics.
******************************************************************************/

USE SUSDB;
GO
SET NOCOUNT ON;

-- Rebuild or reorganize indexes based on their fragmentation levels
DECLARE @work_to_do TABLE (
    objectid int
    , indexid int
    , pagedensity float
    , fragmentation float
    , numrows int
)

DECLARE @objectid int;
DECLARE @indexid int;
DECLARE @schemaname nvarchar(130);
DECLARE @objectname nvarchar(130);
DECLARE @indexname nvarchar(130);
DECLARE @numrows int
DECLARE @density float;
DECLARE @fragmentation float;
DECLARE @command nvarchar(4000);
DECLARE @fillfactorset bit
DECLARE @numpages int

-- Select indexes that need to be defragmented based on the following
-- * Page density is low
-- * External fragmentation is high in relation to index size
PRINT 'Estimating fragmentation: Begin. ' + convert(nvarchar, getdate(), 121)
INSERT @work_to_do
SELECT
    f.object_id
    , index_id
    , avg_page_space_used_in_percent
    , avg_fragmentation_in_percent
    , record_count
FROM
    sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL , NULL, 'SAMPLED') AS f
WHERE
    (f.avg_page_space_used_in_percent < 85.0 and f.avg_page_space_used_in_percent/100.0 * page_count < page_count - 1)
    or (f.page_count > 50 and f.avg_fragmentation_in_percent > 15.0)
    or (f.page_count > 10 and f.avg_fragmentation_in_percent > 80.0)

PRINT 'Number of indexes to rebuild: ' + cast(@@ROWCOUNT as nvarchar(20))

PRINT 'Estimating fragmentation: End. ' + convert(nvarchar, getdate(), 121)

SELECT @numpages = sum(ps.used_page_count)
FROM
    @work_to_do AS fi
    INNER JOIN sys.indexes AS i ON fi.objectid = i.object_id and fi.indexid = i.index_id
    INNER JOIN sys.dm_db_partition_stats AS ps on i.object_id = ps.object_id and i.index_id = ps.index_id

-- Declare the cursor for the list of indexes to be processed.
DECLARE curIndexes CURSOR FOR SELECT * FROM @work_to_do

-- Open the cursor.
OPEN curIndexes

-- Loop through the indexes
WHILE (1=1)
BEGIN
    FETCH NEXT FROM curIndexes
    INTO @objectid, @indexid, @density, @fragmentation, @numrows;
    IF @@FETCH_STATUS < 0 BREAK;

    SELECT
        @objectname = QUOTENAME(o.name)
        , @schemaname = QUOTENAME(s.name)
    FROM
        sys.objects AS o
        INNER JOIN sys.schemas as s ON s.schema_id = o.schema_id
    WHERE
        o.object_id = @objectid;

    SELECT
        @indexname = QUOTENAME(name)
        , @fillfactorset = CASE fill_factor WHEN 0 THEN 0 ELSE 1 END
    FROM
        sys.indexes
    WHERE
        object_id = @objectid AND index_id = @indexid;

    IF ((@density BETWEEN 75.0 AND 85.0) AND @fillfactorset = 1) OR (@fragmentation < 30.0)
        SET @command = N'ALTER INDEX ' + @indexname + N' ON ' + @schemaname + N'.' + @objectname + N' REORGANIZE';
    ELSE IF @numrows >= 5000 AND @fillfactorset = 0
        SET @command = N'ALTER INDEX ' + @indexname + N' ON ' + @schemaname + N'.' + @objectname + N' REBUILD WITH (FILLFACTOR = 90)';
    ELSE
        SET @command = N'ALTER INDEX ' + @indexname + N' ON ' + @schemaname + N'.' + @objectname + N' REBUILD';
    PRINT convert(nvarchar, getdate(), 121) + N' Executing: ' + @command;
    EXEC (@command);
    PRINT convert(nvarchar, getdate(), 121) + N' Done.';
END

-- Close and deallocate the cursor.
CLOSE curIndexes;
DEALLOCATE curIndexes;

IF EXISTS (SELECT * FROM @work_to_do)
BEGIN
    PRINT 'Estimated number of pages in fragmented indexes: ' + cast(@numpages as nvarchar(20))
    SELECT @numpages = @numpages - sum(ps.used_page_count)
    FROM
        @work_to_do AS fi
        INNER JOIN sys.indexes AS i ON fi.objectid = i.object_id and fi.indexid = i.index_id
        INNER JOIN sys.dm_db_partition_stats AS ps on i.object_id = ps.object_id and i.index_id = ps.index_id

    PRINT 'Estimated number of pages freed: ' + cast(@numpages as nvarchar(20))
END
GO

--Update all statistics
PRINT 'Updating all statistics.' + convert(nvarchar, getdate(), 121)
EXEC sp_updatestats
PRINT 'Done updating statistics.' + convert(nvarchar, getdate(), 121)
GO
"@

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Confirm-Prompt ($prompt) {
    <#
    .SYNOPSIS
    Y/N confirmation prompt.

    .DESCRIPTION
    Displays Y/N confirmation prompt and returns true or false.

    .PARAMETER prompt
    String displayed as prompt

    .EXAMPLE
    Confirm-Prompt "Is this a question?"
    #>
    Write-Host "$prompt Y/N: " -BackgroundColor Blue -ForegroundColor White -NoNewline
    $confirm = Read-Host

    if ($confirm.ToLower() -eq 'y') {
        return $true
    } else {
        return $false
    }
}

function Optimize-WsusUpdates {
    <#
    .SYNOPSIS
    Runs all built-in WSUS cleanup processes.

    .DESCRIPTION
    Runs all built-in WSUS cleanup processes.

    .LINK
    https://docs.microsoft.com/en-us/powershell/scripting/developer/help/examples-of-comment-based-help?view=powershell-7
    #>

    Write-Host "Deleting obsolete computers from WSUS database"
    Invoke-WsusServerCleanup -CleanupObsoleteComputers

    Write-Host "Deleting obsolete updates"
    Invoke-WsusServerCleanup -CleanupObsoleteUpdates

    Write-Host "Deleting unneeded content files"
    Invoke-WsusServerCleanup -CleanupUnneededContentFiles

    Write-Host "Deleting obsolete update revisions"
    Invoke-WsusServerCleanup -CompressUpdates

    Write-Host "Declining expired updates"
    Invoke-WsusServerCleanup -DeclineExpiredUpdates

    Write-Host "Declining superceded updates"
    Invoke-WsusServerCleanup -DeclineSupersededUpdates

    Write-Host "Declining additional superceded updates"
    Decline-SupersededUpdates $TRUE
}

function Optimize-WsusDatabase {
    <#
    .SYNOPSIS
    Runs WSUS database optimization.

    .DESCRIPTION
    Runs Microsoft's recommended WSUS database optimization.

    .LINK
    https://support.microsoft.com/en-us/help/4490644/complete-guide-to-microsoft-wsus-and-configuration-manager-sup-maint

    .LINK
    https://devblogs.microsoft.com/scripting/10-tips-for-the-sql-server-powershell-scripter/
    #>

    # Check registry for WSUS database install type (SQL or WID)
    $wsusSqlServerName = (get-itemproperty "HKLM:\Software\Microsoft\Update Services\Server\Setup" -Name "SqlServername").SqlServername

    # Set the named pipe to use based on WSUS db type
    switch -Regex ($wsusSqlServerName) {
        'SQLEXPRESS' { $serverInstance = 'np:\\.\pipe\MSSQL$SQLEXPRESS\sql\query'; break }
        '##WID' { $serverInstance = 'np:\\.\pipe\MICROSOFT##WID\tsql\query'; break }
        '##SSEE' { $serverInstance = 'np:\\.\pipe\MSSQL$MICROSOFT##SSEE\sql\query'; break }
        default { $serverInstance = $wsusSqlServerName }
    }

    # Setting query timeout value because both of these scripts are prone to timeout
    # https://devblogs.microsoft.com/scripting/10-tips-for-the-sql-server-powershell-scripter/

    Write-Host "Creating custom indexes in WSUS index if they don't already exist. This will speed up future database optimizations."
    #Create custom indexes in the database if they don't already exist
    Invoke-Sqlcmd -query $createCustomIndexesSQLQuery -ServerInstance $serverInstance -QueryTimeout 120 -Encrypt Optional

    Write-Host "Running WSUS SQL database maintenence script. This can take an extremely long time on the first run."
    #Run the WSUS SQL database maintenance script
    Invoke-Sqlcmd -query $wsusDBMaintenanceSQLQuery -ServerInstance $serverInstance -QueryTimeout 40000 -Encrypt Optional
}

function New-WsusMaintainenceTask($interval) {
    <#
    .SYNOPSIS
    Creates a new WSUS optimization scheduled tasks.

    .DESCRIPTION
    Creates or overwrites daily or weekly scheduled tasks for WSUS update and database optimization.

    .PARAMETER interval
    Specifies "Daily" or "Weekly" tasks

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/scheduledtasks/?view=win10-ps
    #>

    $taskName = "Optimize WSUS Server ($interval)"
    $scriptPath = 'C:\Scripts'

    # Delete scheduled task with the same name if it already exists
    If (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Unregistered Schedule Task: $taskName"
    }

    # Change scheduled action based on Daily or Weekly
    switch ($interval) {
        'Daily' {
            $trigger = New-ScheduledTaskTrigger -Daily -At "12pm"
            $scriptAction = "-OptimizeServer"
            Break
        }
        'Weekly' {
            $trigger = New-ScheduledTaskTrigger -Weekly -At "2am" -DaysOfWeek Sunday
            $scriptAction = "-OptimizeDatabase"
            Break
        }
        Default {}
    }

    $scriptName = Split-Path $MyInvocation.PSCommandPath -Leaf

    #Create "C:\Scripts" to store PS script
    $null = New-Item -Path "$scriptPath" -ItemType Directory -Force
    Write-Host "Created Directory: $scriptPath"

    # Copy current script to script
    Copy-Item -Path $PSCommandPath -Destination $scriptPath -Force
    Write-Host "Copied Script: $scriptName"

    # Create and register the scheduled task
    $task = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-Command `"&'$($scriptPath)`\$($scriptName)'$scriptAction`""

    $settings = New-ScheduledTaskSettingsSet
    $principal = New-ScheduledTaskPrincipal `
        -UserId "NT AUTHORITY\SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    # Sending to $null to supress output
    $null = Register-ScheduledTask $taskName -Action $task -Trigger $trigger -Settings $settings -Principal $principal

    Write-Host "Registered Scheduled Task: $taskName"
}

function Get-WsusIISConfig {
    <#
    .SYNOPSIS
    Returns a hash of all WSUS optimization related IIS settings.

    .DESCRIPTION
    Determines WSUS IIS Site and Pool, and then forms hash of all relevant optimization settings.

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/webadminstration/get-webapplication?view=winserver2012-ps

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/iisadministration/get-iissite?view=win10-ps
    #>

    # Get WSUS IIS Index from registry
    $iisSiteIndex = Get-ItemPropertyValue "HKLM:\Software\Microsoft\Update Services\Server\Setup" -Name "IISTargetWebSiteIndex"

    # IIS Site
    $iisSiteName = Get-IISSite | Where-Object -Property "Id" -Eq $iisSiteIndex | Select-Object -ExpandProperty "Name"

    # Site Application Pool
    $iisAppPool = Get-WebApplication -site $iisSiteName -Name "ClientWebService" | Select-Object -ExpandProperty "applicationPool"

    # Application Pool Config
    $iisApplicationPoolConfig = Get-IISConfigCollection -ConfigElement (Get-IISConfigSection -SectionPath "system.applicationHost/applicationPools")

    # WSUS Pool Config Root
    $wsusPoolConfig = Get-IISConfigCollectionElement -ConfigCollection $iisApplicationPoolConfig -ConfigAttribute @{"name" = "$iisAppPool" }

    # Queue Length
    $queueLength = Get-IISConfigAttributeValue -ConfigElement $wsusPoolConfig -AttributeName "queueLength"

    #Load Balancer Capabilities
    $wsusPoolFailureConfig = Get-IISConfigElement -ConfigElement $wsusPoolConfig -ChildElementName "failure"
    $loadBalancerCapabilities = Get-IISConfigAttributeValue -ConfigElement $wsusPoolFailureConfig -AttributeName "loadBalancerCapabilities"

    # CPU Reset Interval
    $wsusPoolCpuConfig = Get-IISConfigElement -ConfigElement $wsusPoolConfig -ChildElementName "cpu"
    $cpuResetInterval = (Get-IISConfigAttributeValue -ConfigElement $wsusPoolCpuConfig -AttributeName "resetInterval").TotalMinutes

    # Recycling Config Root
    $wsusPoolRecyclingConfig = Get-IISConfigElement -ConfigElement $wsusPoolConfig -ChildElementName "recycling" | Get-IISConfigElement -ChildElementName "periodicRestart"

    $recyclingMemory = Get-IISConfigAttributeValue -ConfigElement $wsusPoolRecyclingConfig -AttributeName "memory"
    $recyclingPrivateMemory = Get-IISConfigAttributeValue -ConfigElement $wsusPoolRecyclingConfig -AttributeName "privateMemory"

    $clientWebServiceConfig = Get-WebConfiguration -PSPath $iisPath -Filter "system.web/httpRuntime"

    $clientMaxRequestLength = $clientWebServiceConfig | select-object -ExpandProperty maxRequestLength
    $clientExecutionTimeout = ($clientWebServiceConfig | select-object -ExpandProperty executionTimeout).TotalSeconds

    # Return hash of IIS settings
    @{
        QueueLength              = $queueLength
        LoadBalancerCapabilities = $loadBalancerCapabilities
        CpuResetInterval         = $cpuResetInterval
        RecyclingMemory          = $recyclingMemory
        RecyclingPrivateMemory   = $recyclingPrivateMemory
        ClientMaxRequestLength   = $clientMaxRequestLength
        ClientExecutionTimeout   = $clientExecutionTimeout
    }
}

function Get-WsusIISLocalizedNamespacePath {
    # Get localized WSUS IIS web site path: https://docs.microsoft.com/fr-fr/security-updates/windowsupdateservices/18127277 - Document is in English but posted in the French docs
    $iisSitePhysicalPath = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Update Services\Server\Setup\' -Name "TargetDir"
    $iisLocalizedString = Get-Website | Where-Object {$($_.PhysicalPath).StartsWith($iisSitePhysicalPath)} | Select-Object -ExpandProperty Name
    $iisLocalizedNamespacePath = "IIS:\Sites\$iisLocalizedString\ClientWebService"
    return $iisLocalizedNamespacePath
}

function Test-WsusIISConfig ($settings, $recommended) {
    <#
    .SYNOPSIS
    Compares current WSUS IIS settings to recommended values.

    .DESCRIPTION
    Compares current WSUS IIS settings to recommended values. Prompts user to commit changes.

    .PARAMETER settings
    Hash of current WSUS IIS settings.

    .PARAMETER recommended
    Hash of recommended WSUS IIS settings.
    #>

    # Delay IIS configuration commits until we're done updating all necessary settings
    Start-IISCommitDelay

    foreach ($key in $recommended.Keys) {
        # If the current configuration setting doesn't match the recommended value, prompt the user to update
        # This could be better designed to match minimum requirements instead of specific values, but it isn't.
        If ($recommended[$key] -ne $settings[$key]) {
            Write-Host "$key`n`tCurrent:`t$($settings[$key])`n`tRecommended:`t$($recommended[$key])" -BackgroundColor Black -ForegroundColor Red

            if (Confirm-Prompt "Update $key to recommended value?") {
                Update-WsusIISConfig $key $recommended[$key]
            }
        }
        else {
            Write-Host "$key`n`tCurrent:`t$($settings[$key])`n`tRecommended:`t$($recommended[$key])" -BackgroundColor Black -ForegroundColor Green
        }
    }

    # Allow IIS config commits again
    Stop-IISCommitDelay
}

function Update-WsusIISConfig ($settingKey, $recommendedValue) {
    <#
    .SYNOPSIS
    Modifies IIS configuration for specified setting.

    .DESCRIPTION
    Modifies specified IIS setting for WSUS IIS Site/App Pool optimization.

    .PARAMETER settingKey
    String used to reference specific IIS configuration setting.

    .PARAMETER recommendedValue
    Recommended value for WSUS IIS configuration setting.
    #>

    # WSUS IIS Index
    $iisSiteIndex = Get-ItemPropertyValue "HKLM:\Software\Microsoft\Update Services\Server\Setup" -Name "IISTargetWebSiteIndex"

    # IIS Site
    $iisSiteName = Get-IISSite | Where-Object -Property "Id" -Eq $iisSiteIndex | Select-Object -ExpandProperty "Name"

    # Site Application Pool
    $iisAppPool = Get-WebApplication -site $iisSiteName -Name "ClientWebService" | Select-Object -ExpandProperty "applicationPool"

    # Application Pool Config
    $iisApplicationPoolConfig = Get-IISConfigCollection -ConfigElement (Get-IISConfigSection -SectionPath "system.applicationHost/applicationPools")

    # WSUS Pool Config Root
    $wsusPoolConfig = Get-IISConfigCollectionElement -ConfigCollection $iisApplicationPoolConfig -ConfigAttribute @{"name" = "$iisAppPool" }

    # Recycling Config Root
    $wsusPoolRecyclingConfig = Get-IISConfigElement -ConfigElement $wsusPoolConfig -ChildElementName "recycling" | Get-IISConfigElement -ChildElementName "periodicRestart"

    switch ($settingKey) {
        'QueueLength' {
            # Queue Length
            Set-IISConfigAttributeValue -ConfigElement $wsusPoolConfig -AttributeName "queueLength" -AttributeValue $recommendedValue
            Break
        }
        'LoadBalancerCapabilities' {
            # Failure Config Root
            $wsusPoolFailureConfig = Get-IISConfigElement -ConfigElement $wsusPoolConfig -ChildElementName "failure"

            # Load Balancer Capabilities
            Set-IISConfigAttributeValue -ConfigElement $wsusPoolFailureConfig -AttributeName "loadBalancerCapabilities" -AttributeValue $recommendedValue
            Break
        }
        'CpuResetInterval' {
            # CPU Reset Interval
            $wsusPoolCpuConfig = Get-IISConfigElement -ConfigElement $wsusPoolConfig -ChildElementName "cpu"
            Set-IISConfigAttributeValue -ConfigElement $wsusPoolCpuConfig -AttributeName "resetInterval" -AttributeValue ([timespan]::FromMinutes($recommendedValue))
            Break
        }
        'RecyclingMemory' {
            Set-IISConfigAttributeValue -ConfigElement $wsusPoolRecyclingConfig -AttributeName "memory" -AttributeValue $recommendedValue
            Break
        }
        'RecyclingPrivateMemory' {
            Set-IISConfigAttributeValue -ConfigElement $wsusPoolRecyclingConfig -AttributeName "privateMemory" -AttributeValue $recommendedValue
            Break
        }
        'ClientMaxRequestLength' {
            # Check if the IIS WSUS Client Web Service web.config is read only and make it RW if so
            Unblock-WebConfigAcl
            Set-WebConfigurationProperty -PSPath $iisPath -Filter "system.web/httpRuntime" -Name "maxRequestLength" -Value $recommendedValue
            Break
        }
        'ClientExecutionTimeout' {
            # Check if the IIS WSUS Client Web Service web.config is read only and make it RW if so
            Unblock-WebConfigAcl
            Set-WebConfigurationProperty -PSPath $iisPath -Filter "system.web/httpRuntime" -Name "executionTimeout" -Value ([timespan]::FromSeconds($recommendedValue))
            Break
        }
        Default {}
    }

    Write-Host "Updated IIS Setting: $settingKey, $recommendedValue" -BackgroundColor Green -ForegroundColor Black
}

function Remove-Updates ($searchStrings, $updateProp, $force=$false) {
    [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null
    $wsusServer = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();
    $scope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $updates = $wsusServer.GetUpdates($scope)
    $declinedCount = 0
    $searchCount = 0
    $userMsg = 'Found'
    $color = 'Yellow'

    if ($force) {
        $userMsg = 'Declined'
        $color = 'DarkGreen'
    }

    Write-Host "Update Property: $updateProp"

    foreach ($searchString in $searchStrings)
    {
        $confirm = $false
        Write-Host " - Update Search: $searchString"
        $searchCount = 0
        foreach ($update in $updates){
            if ($update.$($updateProp) -match "$searchString"){
                if($update.IsApproved){

                    if ($force){
                        $update.Decline()
                    }
                    $searchCount = $searchCount + 1
                    Write-Host "   [*]$($userMsg): $($update.Title), $($update.ProductTitles) ($searchString)" -ForegroundColor $color
                }
            }
        }

        if ($searchCount -gt 0) {
            Write-Host "$searchCount `"$searchString`" Updates $userMsg!" -ForegroundColor "Blue" -BackgroundColor White
        } else {
            Write-Host "      $searchCount `"$searchString`" Updates $userMsg" -ForegroundColor "White"
        }

        #Prompt user to confirm declining updates. Do no prompt if force flag is enable to prevent loop
        if ((-not $force) -and ($searchCount -ne 0)){
            $confirm = Confirm-Prompt "Are you sure you want to decline all ($searchCount) listed ($searchString) updates?"

            if ($confirm) {
                Remove-Updates @($searchString) $updateProp $true | out-null
            }
        }

        if (($confirm) -or $force){
            $declinedCount = ($declinedCount + $searchCount)
        }
    }

    return $declinedCount
}

function Invoke-DeepClean ($titles, $productTitles) {
    <#
    .SYNOPSIS
    Checks for unneeded WSUS updates to be deleted.

    .DESCRIPTION
    Checks for unneeded WSUS updates by product category to be deleted.

    .PARAMETER titles
    Array of titles of WSUS titles to search and prompt for removal

    .PARAMETER productTitles
    Array of WSUS product titles to search and prompt for removal

    .EXAMPLE
    DeepClean $titles $products

    .NOTES
    WSUS GetUpdates Method
    https://docs.microsoft.com/en-us/previous-versions/windows/desktop/aa350127(v=vs.85)

    WSUS IUpdate Properties
    https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms752741(v=vs.85)

    WSUS Product List
    Get-WsusProduct
    https://docs.microsoft.com/en-us/powershell/module/wsus/get-wsusproduct?view=win10-ps

    WSUS Classification List
    Get-WsusClassification
    https://docs.microsoft.com/en-us/powershell/module/wsus/get-wsusclassification?view=win10-ps
    #>

    $declinedTotal = 0

    Write-Host "Make certain to carefully read the listed updates before choosing to remove them!" -BackgroundColor White -ForegroundColor Green

    #Remove updates by Title
    Write-Host "Searching for unneeded updates by Title. This process can take a long time. Please wait." -BackgroundColor White -ForegroundColor Blue
    $declinedTotal += Remove-Updates $titles 'Title'

    #Remove updates by ProductTitles
    Write-Host "Searching for unneeded updates by ProductTitle. This process can take a long time. Please wait." -BackgroundColor White -ForegroundColor Blue
    $declinedTotal += Remove-Updates $productTitles 'ProductTitles'

    #Remove drivers
    Write-Host "Searching for drivers to be removed from WSUS. This process can take a long time. Please wait." -BackgroundColor White -ForegroundColor Blue
    $declinedTotal += Remove-Updates @('Drivers') 'UpdateClassificationTitle'

    Write-Host "Searching for unneeded updates superseded by newer updates. This process can take a long time. Please wait." -BackgroundColor White -ForegroundColor Blue
    $declinedTotal += Decline-SupersededUpdates

    Write-Host "================DEEPCLEAN COMPLETE==================" -BackgroundColor White -ForegroundColor Blue
    Write-Host "$declinedTotal Total Updates Declined" -BackgroundColor White -ForegroundColor Blue
}

function Disable-WsusDriverSync {
    <#
    .SYNOPSIS
    Disable WSUS device driver syncronization and caching.

    .DESCRIPTION
    Disable WSUS device driver syncronization and caching. Automatic driver sychronization is one of the primary causes of WSUS slowness, crashing, and wasted storage space.

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/updateservices/set-wsusclassification?view=win10-ps
    #>

    Get-WsusClassification | Where-Object -FilterScript {$_.Classification.Title -Eq "Drivers"} | Set-WsusClassification -Disable
    Get-WsusClassification | Where-Object -FilterScript {$_.Classification.Title -Eq "Driver Sets"} | Set-WsusClassification -Disable
}


function Unblock-WebConfigAcl {
    <#
    .SYNOPSIS
    Grants local admins access to web.config

    .DESCRIPTION
    Grants BUILTIN\Administrators ownership and read write access to ClientWebService web.config. Also removes Read Only flag.

    .LINK
    https://devblogs.microsoft.com/scripting/use-powershell-to-translate-a-users-sid-to-an-active-directory-account-name/
    https://docs.microsoft.com/en-us/dotnet/api/system.security.principal.securityidentifier.-ctor?view=windowsdesktop-5.0#System_Security_Principal_SecurityIdentifier__ctor_System_String_
    #>

    $wsusWebConfigPath = Get-WebConfigFile -PSPath $iisPath | Select-Object -ExpandProperty 'FullName'

    # Get localized BUILTIN\Administrators group
    $builtinAdminGroup = ([System.Security.Principal.SecurityIdentifier]'S-1-5-32-544').Translate([System.Security.Principal.NTAccount]).Value

    Set-FileAclOwner $wsusWebConfigPath $builtinAdminGroup
    Set-FileAclPermissions $wsusWebConfigPath $builtinAdminGroup 'FullControl' 'None' 'None' 'Allow'
    Set-ItemProperty -Path $wsusWebConfigPath -Name IsReadOnly -Value $false
}

function Set-FileAclOwner ($file, $owner) {
    <#
    .SYNOPSIS
    Sets NTFS file owner

    .DESCRIPTION
    Sets NTFS file owner

    .PARAMETER file
    File path as string

    .PARAMETER owner
    Account as string to set as owner

    .LINK
    https://stackoverflow.com/questions/22988384/powershell-change-owner-of-files-and-folders
    #>

    $acl = Get-Acl($file)
    $account = New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList $owner
    $acl.SetOwner($account)
    Set-Acl -Path $file -AclObject $acl
}

function Set-FileAclPermissions ($file, $accString, $rights, $inheritanceFlags, $propagationFlags, $type) {
    <#
    .SYNOPSIS
    Set NTFS file permissions

    .DESCRIPTION
    Set NTFS permissions for specified file

    .PARAMETER file
    File path as string

    .PARAMETER accString
    Account to set permissions for as string

    .PARAMETER rights
    Access Rights - https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.filesystemrights?view=dotnet-plat-ext-3.1

    .PARAMETER inheritanceFlags
    Inheritence flags - https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.inheritanceflags?view=dotnet-plat-ext-3.1

    .PARAMETER propagationFlags
    Propagation flags - https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.propagationflags?view=dotnet-plat-ext-3.1

    .PARAMETER type
    Access control type - https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.accesscontroltype?view=dotnet-plat-ext-3.1

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-acl?view=powershell-7

    .LINK
    https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.filesystemaccessrule.-ctor?view=dotnet-plat-ext-3.1#System_Security_AccessControl_FileSystemAccessRule__ctor_System_String_System_Security_AccessControl_FileSystemRights_System_Security_AccessControl_InheritanceFlags_System_Security_AccessControl_PropagationFlags_System_Security_AccessControl_AccessControlType_
    #>

    $acl = Get-Acl($file)
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $accString, $rights, $inheritanceFlags, $propagationFlags, $type
    $acl.SetAccessRule($accessRule)
    Set-Acl -Path $file -AclObject $acl
}

function Decline-SupersededUpdates ($verbose){
    <#
    .SYNOPSIS
    Declines approved updates that have been approved and are superseded by other updates.

    .DESCRIPTION
    Declines all updates that have been approved and are superseded by other updates. The update will only be declined if a superseding update has been approved.

    .LINK
    ApprovedStates - https://docs.microsoft.com/en-us/previous-versions/windows/desktop/aa354257(v=vs.85)

    .LINK
    IUpdate - https://docs.microsoft.com/en-us/previous-versions/windows/desktop/bb313429(v=vs.85)

    .LINK
    UpdateCollection - https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms752803(v=vs.85)
    #>
    $declineCount = 0
    [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null
    $wsusServer = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();
    $scope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

    $scope.ApprovedStates = "LatestRevisionApproved"
    $updates = $wsusServer.GetUpdates($scope)

    foreach ($update in $updates){
        $updatesThatSupersede = $update.GetRelatedUpdates("UpdatesThatSupersedeThisUpdate")
        if($updatesThatSupersede.Count -gt 0) {
            foreach ($super in $updatesThatSupersede)
            {
                if ($super.IsApproved){
                    $update.Decline()
                    $declineCount++
                    break
                }
            }
        }
    }

    if($verbose) {
        Write-Host "Osbolete Updates Declined: $declineCount"
    } else {
        return $declineCount
    }
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------
$iisPath = Get-WsusIISLocalizedNamespacePath

# Check commandline parameters.
switch($true) {
    ($FirstRun) {
        Write-Host "All of the following processes are highly recommended!" -ForegroundColor Blue -BackgroundColor White

        switch($true) {
            (Confirm-Prompt "Run WSUS IIS configuration optimization?") {
                $wsusIISConfig = Get-WsusIISConfig
                Test-WsusIISConfig $wsusIISConfig $recommendedIISSettings
            }
            (Confirm-Prompt "Run WSUS database optimization?") {
                Optimize-WsusDatabase
            }
            (Confirm-Prompt "Run WSUS server optimization?") {
                Optimize-WsusUpdates
            }
            (Confirm-Prompt "Create daily WSUS server optimization scheduled task?") {
                New-WsusMaintainenceTask('Daily')
            }
            (Confirm-Prompt "Create weekly WSUS database optimization scheduled task?") {
                New-WsusMaintainenceTask('Weekly')
            }
            (Confirm-Prompt "Disable device driver synchronization?") {
                Disable-WsusDriverSync
            }
        }
        Break
    }
    ($DisableDrivers) {
        Disable-WsusDriverSync
    }
    ($DeclineSupersededUpdates) {
        Decline-SupersededUpdates
    }
    ($DeepClean) {
        Invoke-DeepClean $unneededUpdatesbyTitle $unneededUpdatesbyProductTitles
    }
    ($InstallDailyTask) {
        New-WsusMaintainenceTask('Daily')
    }
    ($InstallWeeklyTask) {
        New-WsusMaintainenceTask('Weekly')
    }
    ($CheckConfig) {
        $wsusIISConfig = Get-WsusIISConfig
        Test-WsusIISConfig $wsusIISConfig $recommendedIISSettings
    }
    ($OptimizeServer) {
        Optimize-WsusUpdates
    }
    ($OptimizeDatabase) {
        Optimize-WsusDatabase
    }
}
