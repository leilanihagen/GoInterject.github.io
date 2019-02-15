---
title: Technical Install
layout: custom
keywords: []
description: 
---

# Prerequisites

1. Verify SQL Server Version is > 2008

    *To check your MSSQL version, run ```SELECT @@VERSION``` in a query window.* 

2. SQL Management Studio 
3. Download Initial Deploy Scripts from secured portal

# Steps Required to Install

## 1. Establish connection to Appropiate DB Server
## 2. Create Interject Reporting Database Schema

- CREATE [Interject_Reporting] database, if not already exist 

   ![](/images/A-SQL-Installation/01.png)

## 3. Execute Initial Deploy Script on New Database

### DB Object Creation
### DB permissions and roles

The security model can be setup in a few ways, by location or other segment.  The security in epicor is by database, which normally has a group of location segments.
Execute Read Only Access Setup Scripts
Applies Certificates on Epicor DBs for select tables


## 4. Execute Read Only Access Setup Scripts

Execute Read Only Access Setup Scripts

   
