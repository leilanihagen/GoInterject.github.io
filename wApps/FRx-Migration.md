---
title: FRx Migration
layout: custom
keywords: []
description: 
---

# Prerequisites

1. Make sure that ODBC connection to Microsoft Access has already existed (if you installed MS Office before, it should be already done)

![](/images/FRxReplacement/02.png)

# Steps Required to Install

### 1. Open SQL Server and Create a new database

### 2.  Import MDB file to the new database
Verbose Version:

![](/images/FRxReplacement/06.png)

__Note:__ _If Access database is protected you will need to remove the database password before gouping through with the import._

![](/images/FRxReplacement/07.png)

![](/images/FRxReplacement/03.png)

![](/images/FRxReplacement/05.png)


### 3. Add FRx Conversion Scripts

   a) Download PoShDbToolGUI application by cloning the follow repo using terminal (ie: Powershell): 

```PowerShell
git clone https://gitlab.com/Open-Interject/PowershellDBToolsGui.git -b feature/payload_file
cd .\PowershellDBToolsGui\PoShDbToolGUI\bin\Release\
.\DbTools.exe
```

    b) Clone the the scripts into retro folder

```
git clone https://gitlab.com/Open-Interject/FRxExportColumnRow.git
git ls-tree --full-tree -r --name-only HEAD > ..\payload.txt
```
### Open PoSh GUI 


a.  Assign Repo Folder Location

b.  Assign Payload File Location

c.  Assign Server Name

d.  Assign Database Name

e.  Execute SQL changes

![](/images/A-SQL-Installation/03.png)


