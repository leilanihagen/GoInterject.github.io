---
title: Technical Install
layout: custom
keywords: []
description: 
---

## Prerequisites

1. Verify SQL Server Version is > 2008

    *To check your MSSQL version, run ```SELECT @@VERSION``` in a query window.* 

2. SQL Management Studio 
3. Download Initial Deploy Scripts using the following link [Initial.Interject_Reporting.sql](https://drive.google.com/a/gointerject.com/uc?authuser=0&id=1fdddeCsvwNwF5VqICLoAZU4KSkcZKSyx&export=download). If unable to download email [help@gointerject.com](help@gointerject.com) to get access to file.

## Steps Required to Install

### 1. Establish connection to Appropiate DB Server

Using MSSMS connect to your Epicor server using sysadmin user. 

### 2. Create Interject Reporting Database

- CREATE [Interject_Reporting] database, if not already exist 

   ![](/images/A-SQL-Installation/01.png)

### 3. Execute Initial Deploy Script on New Database

Open [Initial.Interject_Reporting.sql](https://drive.google.com/a/gointerject.com/uc?authuser=0&id=1fdddeCsvwNwF5VqICLoAZU4KSkcZKSyx&export=download) script. Executing script on the newly created database. This script creates all DB Object.

### 4. Create security objects and grand read only access to Eipor tables

Pass the following 2 parameters to [Custom].[Interject_SetupScript1_Security]    
* MasterEpicorDatabase - Specify the master Epicor Batabase
* CertificatePassword - Create a certificate with a custom password. 

**Example:**
```SQL
EXEC [Custom].[Interject_SetupScript1_Security]
	  @MasterEpicorDatabase = '[DemoControl]'
	 ,@CertificatePassword =  'myPassword1234'
```







