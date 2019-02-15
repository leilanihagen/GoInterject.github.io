---
title: FRx Migration
layout: custom
keywords: []
description: 
---

# Prerequisites

1. Verify ODBC connection to Microsoft Access has already existed (if you installed MS Office before, it should be already done)

![](/images/FRxReplacement/02.png)

2. Download Initial FRX Migration Script
3. SQL Management Studio 

# Steps Required to Install

## 1. Open SQL Server and Create a new database

## 2.  Import MDB file to the new database
Verbose Version:

![](/images/FRxReplacement/06.png)

__Note:__ _If Access database is protected you will need to remove the database password before gouping through with the import._

![](/images/FRxReplacement/07.png)

![](/images/FRxReplacement/03.png)

![](/images/FRxReplacement/05.png)

## 3. Execute  FRx Migration Script 

![](/images/FRxReplacement/08.png)

## 4. Setup DB Connections in Interject Portal 

![](/images/FRxReplacement/09.png)
  
## 5. Redirect DB Connection to new DB  in interject process

![](/images/FRxReplacement/10.png)


