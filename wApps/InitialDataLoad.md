---
title: Initial Data Load
layout: custom
keywords: []
description: 
---

## Prerequisites

1. Verify SQL Server Version is > 2008

    *To check your MSSQL version, run ```SELECT @@VERSION``` in a query window.* 

2. SQL Management Studio 
3. Interject financial 
4. User admin on SQL server

## Steps Required to Install

### 1. Execute Script DB to initialize Epicor Enterprise Data

```SQL
--Import configuration setup from Epicor and initial setup of Interject
EXEC [Custom].[EPR_InstallScript1_DatabaseConfig]
	 @MasterEpicorDatabase          = '[MasterDatbase]'
	,@DefaultDatabaseNameSource     = '[DefaultDatabase]'

--Import data 
EXEC [Custom].[EPR_InstallScript2_EpicorImport]

EXEC [Custom].[EPR_InstallScript3_ReportingImport]
	 @ReportingImport_YearBegin     = '2000'
	,@ReportingImport_YearEnd       = '2016'

EXEC [Custom].[EPR_InstallScript4_GroupingImport]

```

### 2. Execute Script SQL Agent Jobs

```SQL
EXEC [Custom].[EPR_InstallScript5_SetupJobs]
```

Executing script will create three jobs: 
* [Interject_Reporting_CheckSchedule_ImportActual] - Syncs Actual data between Epicor tables to Interject table
* [Interject_Reporting_CheckSchedule_ImportBudget] - Syncs Budget data between Epicor tables to Interject table
* [Interject_Reporting_AddJobsFromScheduler] - Process data and distribute it to interject tables 
* [Interject_Reporting_ImporEpicor_DeletesRecords] - Validates data and remove records nightly if data was removed form epicor tables 


### 3. Setup DB Connections in Interject Portal

**Step 1:** Navigate to [portal.gointerject.com](https://portal.gointerject.com). After logging on select Data Connections left side menu.

**Step 2:** In the Data Connections page select the New Connection button in the top right-hand corner.

**Step 3:** In the Connection Type field, make sure Database is selected.

**Step 4:** The Connection Details page will contain the following information for the new connection.

**Name:** A unique friendly name used when connecting a Data Portal to the Data Connection

**Description (optional):** description of what the connection string is connecting to

**Connection String:** used by INTERJECT to connect to the specified server & database

![](/images/Database/04.png)

Detail instructions how to setup and how to test [Database Connection](/wPortal/L-Database-Connection.html)

### 4. Redirect DB Connection to new DB in Interject portal

Step 1:Navigate to [portal.gointerject.com](https://portal.gointerject.com). After logging on select Data Connections left side menu.

Step 2: In the My Apps page select Epicor Tools. 

Step 3: At the bottom of My Apps page there will be a Connection Redirect section. Select the connection to overide in the left side and new connection to your data on the right.

![](/images/A-InitialDataLoad/01.png)

