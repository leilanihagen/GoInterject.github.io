---
title: SQL Database Installation
layout: custom
keywords: []
description: 
---

# Prequisites

1. Verify SQL Server Version is > 2008

    *To check your MSSQL version, run ```SELECT @@VERSION``` in a query window.* 

2. Obtain Git Labs Account and Permissions granted
3. Obtain Git Application
4. SQL Management Studio and Windows Powershell


# Steps Required to Install

### 1. Establish connection to Appropiate DB Server
### 2. Verify SQL Database Mail Profile
    To check the Mail profile, run ```SELECT name FROM msdb.dbo.sysmail_profile```.

### 3. Create Interject Reporting Database Schema

- CREATE [Interject_Reporting] database, if not already exist 

   ![](/images/A-SQL-Installation/01.png)


### 4. Deploy DB Objects via PoSH scripting

   *Requires git application and access to repository* 

    a) Download PoShDbToolGUI application by cloning the follow repo using terminal (ie: Powershell): 

```PowerShell
git clone https://gitlab.com/Open-Interject/PowershellDBToolsGui.git -b feature/payload_file

cd .\PowershellDBToolsGui\PoShDbToolGUI\bin\Release\
.\DbTools.exe
```

[PoShDbToolGUI]:https://gitlab.com/Open-Interject/PowershellDBToolsGui/raw/feature/payload_file/PoShDbToolGUI/bin/Release/DbTools.exe?inline=false

    b) Create copy of the Interject reporting repos using the "Epicor" branch and generate a payload from git repo. 

```PowerShell
git clone https://gitlab.com/Interject/Interject_Reporting.git -b Epicor

cd Interject_Reporting

git ls-tree --full-tree -r --name-only HEAD > ..\payload.txt
```

#### c)  Open PoSh GUI 

1c.  Assign Repo Folder Location

2c.  Assign Payload File Location

3c.  Assign Server Name

4c.  Assign Database Name

5c.  Execute SQL changes

![](/images/A-SQL-Installation/03.png)

### 5. Script DB to initialize Epicor Enterprise Data

    Execute below scripts to initialize Epicor Enterprise data for Interject Financials for Spreadsheets Within using SQL Server on Interject Reporting database from step 2.

```SQL
EXECUTE [Custom].[EPR_InstallScript1_DatabaseConfig]
	 @MasterEpicorDatabase = '[DemoControl]'
	,@DefaultDatabaseNameSource = '[DemoHold]'

EXECUTE [Custom].[EPR_InstallScript2_EpicorImport]
EXECUTE [Custom].[EPR_InstallScript3_ReportingImport]

EXECUTE [Custom].[EPR_InstallScript3_ReportingImport]
	 @ReportingImport_YearBegin = '1996' --Accont beginning
	,@ReportingImport_YearEnd = '2000'

EXECUTE [Custom].[EPR_InstallScript4_GroupingImport] 
```

### 6. Assign Interject Database Role and permissions

    a) Execute **ReportingDB_Permissions.sql** in **[Interject_Reporting]** db to implement the following on the SQL Server:

[download script][1] 

-	CREATE **[db_Interject]** database role
-	IF SQL-AUTH: CREATE **[InterjectAppUser]** database user and add to database role
-	IF WIN-AUTH:  CREATE **[INTERJECT\InterjectUsers]** database user and add to database role
-	APPLY PERMISSIONS TO SCHEMAS for database role
    - **[Client]** – EXECUTE
    - **[Custom]** – EXECUTE
    - **[Report]** – EXECUTE

    b) Execute **ReportingDB_AddSignaturePermissions.sql** in **[Interject_Reporting]** db to implement the following on the SQL Server:

 [download script][2] 

    ![](/images/A-SQL-Installation/02.png)

### 7. Enable Security Certificates

    *Edit the script to set an alternative password for the certificate *

-	CREATE CERTIFICATE **[InterjectCertificate]**
-	CREATE CERTIFICATE USER **[InterjectCertificateUser]**
-	Add certificate to stored procedures with dynamic sql.
    - [Custom].[GL_COA_withBalances]
    - [Custom].[GL_JEQuery]
    - [Custom].[EPR_Grouping_Import]
    - [Custom].[GL_Segment]
    - [FinCube].[FinalCalculation]
    - [FinCube].[FinalCalculation]
    - [Note].[Note_Save]
    - [Note].[Note_Fixed_Get]
    - [Note].[Note_Save]
    - [Report].[MembersByCategory_Pull]
    - [Client].[FinCube_DynamicRow]

-	APPLY PERMISSIONS TO SCHEMAS for certificate user
    - [FSGroup] – SELECT   
    - [FSData] – SELECT
    - [ImportERP] – SELECT

### 8. Install SQL Agent Jobs
     Edit and Execute each of the following scripts to install the SQL Server Agent Jobs that accompany the Interject Solution. 

```SQL
EXECUTE [Custom].[EPR_InstallScript5_SetupJobs]
```

-	10.SQLAgentJob_Interject_Reporting_AddJobsFromScheduler.sql
-	10.SQLAgentJob_Interject_Reporting_CheckSchedule_ImportActual.sql
-	10.SQLAgentJob_Interject_Reporting_CheckSchedule_ImportBudget.sql
-	10.SQLAgentJob_Interject_Reporting_ImportEpicor_DeleteRecords.sql

The scripts include default schedules and assumes to be executed on same server as the **[Interject_Reporting]** and Epicor company databases.



[1]:{{ site.url }}/images/A-SQL-Installation/ReportingDB_Permissions.sql
[2]:{{ site.url }}/images/A-SQL-Installation/ReportingEpicor_AddSignature_Permission.sql

### 9. Setup DB Connections in Interject Portal
Setup database connection. [ Database Connection ](/wPortal/L-Database-Connection.html)

**PROD:** Data Source = ""; Initial Catalog=""; Integrated Security=SSPI

**BETA:** Data Source = ""; Initial Catalog=""; Integrated Security=SSPI 

**DEV:** Data Source = ""; Initial Catalog=""; Integrated Security=SSPI

**Other Connections:** 
- "": Used for initial SYSDATA migration only in DEV.
- "": Used to point to Laminin DEV Lab.

### 10. Enable Subscriptions to Interject Financials
