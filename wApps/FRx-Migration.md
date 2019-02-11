---
title: FRx Migration
layout: custom
keywords: []
description: 
---

## FRx Implementation Overview 

### Import MDB FRx file to SQL Server 

1. Make sure that ODCB connection to Microsoft Access has already existed (if you installed MS Office before, it should be already done)

![](/images/FRxReplacement/02.png)

2. Open SQL Server and Create a new database

3. Import MDB file to the new database
Verbose Version:

![](/images/FRxReplacement/06.png)

__Note:__ _If Access database is protected you will need to remove the database password before gouping through with the import._

![](/images/FRxReplacement/07.png)

![](/images/FRxReplacement/03.png)

![](/images/FRxReplacement/05.png)


### Add FRx Conversion Scripts

Download PoShDbToolGUI application by cloning the follow repo. 

```PowerShell
git clone https://gitlab.com/Open-Interject/PowershellDBToolsGui.git -b feature/payload_file
cd .\PowershellDBToolsGui\PoShDbToolGUI\bin\Release\
.\DbTools.exe
```

Clone the the scripts into retro folder

```
git clone https://gitlab.com/Open-Interject/FRxExportColumnRow.git
git ls-tree --full-tree -r --name-only HEAD > ..\payload.txt
```
Open PoSh GUI 

![](/images/A-SQL-Installation/03.png)


