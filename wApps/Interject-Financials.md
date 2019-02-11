---
title: INTERJECT Financials
layout: custom
keywords: []
description: INTERJECT™ Financials specifics (Topics that are unique/specific to the Financials Application) 
---
## Overview

## My Apps
**Interject Financials:** Includes 44 data portals related to *Financials for Spreadsheets.* Data connection is redirected through "" 
**Epicor Tools** Includes 3 data portals to help migrate FRx reports. 

## Data Connections

**Connection String:**

Setup database connection. [ Database Connection ](/wPortal/L-Database-Connection.html)

**PROD:** Data Source = ""; Initial Catalog=""; Integrated Security=SSPI

**BETA:** Data Source = ""; Initial Catalog=""; Integrated Security=SSPI 

**DEV:** Data Source = ""; Initial Catalog=""; Integrated Security=SSPI

**Other Connections:** 
- "": Used for initial SYSDATA migration only in DEV.
- "": Used to point to Laminin DEV Lab.

## Data Portals

""
- Import yearly numbers into the Interject Reporting DB from a list of Epicor DBs
- Connection: ""

""
- import yearly budget numbers into the Interject Reporting DB from a list of Epicor DBs
- Connection: ""

""
- Summary & detail report of Epicor Chart of Accounts with balances
- Override of "" from offering
- Connection: ""