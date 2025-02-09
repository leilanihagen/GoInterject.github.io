---
title: "Lab Drill: Inventory Report"
layout: custom
keywords: [drill,unfreezing,unfreeze,report,hyperlinks]
description: In this example, you will view drilling between reports using the same Inventory reports created during the Real World Inventory Walk-through.
---
* * *

##  **Overview**

In this example, you will view drilling between reports using the same Inventory reports created during the [ Real World Inventory Walk-through ](/wAbout/Inventory-Reports.html) . You will setup a drill between the **Inventory by Category** and **Inventory by Detail** pages of the workbook, then you will set up a hyperlink so the drill can be more intuitive to users. 




###  Unfreezing the Report 

**Step 1:** First, open the Inventory Report Template. This file has been prepared specifically for this exercise. Once opened, it should look like the screenshot below. 

![](/images/L-Drill-Inventory/01.png)
<br>

**Step 2:** Now unfreeze the panes to access the report formulas. 

![](/images/L-Drill-Inventory/02.png)
<br>
  


###  Setting Up The Drill 

**Step 1:** Type [ **=ReportDrill()** ](/wIndex/ReportDrill.html) into cell C5. Then, click the **fx** button to bring up the Function Wizard. 

![](/images/L-Drill-Inventory/03.png)

**Step 2:** Now type **InventoryByDetail!C4** into the ReportCellToRun argument to specify the range you want to navigate to. You will skipping the ReportCodeToRun argument since that is only used for drilling to other workbooks in the Report Library. 

![](/images/L-Drill-Inventory/04.png)
<br>
  


**Step 3:** Next use the  TransferPairs argument to note which cell values in the source worksheet will be transferred to the target worksheet during the drill operation. To do this, use special functions to pair the source cells to the target cells. Type  [ **PairGroup(Pair())** ](/wIndex/PairGroup.html) in the TransferPairs argument to get it started. 

![](/images/L-Drill-Inventory/05.png)
<br>
  


**Step 4:** In the Formula Bar, click within the word **Pair()** inside the text **PairGroup(Pair())** while the Function Wizard is open. See the illustration below. Once this is done, the Function Wizard will automatically change to help with the Pair() function. Type **B15:B23** in the From argument as shown below. Column B is where the CustomerID will be shown in the source report. By noting a range from row 22 to 24 in column B, INTERJECT will expand that range to the data that is presented in this source report. 

![](/images/L-Drill-Inventory/06.png)
<br>

**Step 5:** Next, select the Target argument and navigate to the **InventoryByDetail** tab. You want to place the CustomerID in cell E11 during the drill operation. Excel will fill in the formula automatically based on where you click. Click **OK** to finish updating the function and it will take you back to the source report. 

![](/images/L-Drill-Inventory/07.png)
<br>
  


After pressing OK, the report formula should look as it does in the image below. 

![](/images/L-Drill-Inventory/08.png)
<br>

###  Creating Hyperlinks as Drills 

**Step 1:** Now you are going to create hyperlinks for the drill. First, highlight the cells you want to setup the hyperlink, then right click and choose **Link**. In some versions of Excel it will show as **Hyperlink**. 

![](/images/L-Drill-Inventory/09.jpg)
<br>

**Note:** Each drill will need to be linked individually, not all at once. If they are linked all at once then the drills will not work as it will drill everything at once, rather than one at a time. 

**Step 2:** In the Hyperlink pop-up window, you will select **Place in This Document**. Then select **ScreenTip**, type **Interject Drill**, and press OK. Although this technically sets up a hyperlink to cell A1 in the same tab, INTERJECT will override the event so the INTERJECT drill will activate. 

![](/images/L-Drill-Inventory/10.png)
<br>
  


Once you select **OK**, the cells will be linked to the drill, as shown below. 

![](/images/L-Drill-Inventory/11.png)
<br>
  


When the panes are refrozen, the report should look like the image below. 

![](/images/L-Drill-Inventory/12.png)
<br>
  


**Step 3:** [ Pull Data ](/wGetStarted/INTERJECT-Ribbon-Menu-Items.html) to see data for each category you just linked. 

![](/images/L-Drill-Inventory/13.png)
<br>
  


Here, you have the report pulled and are ready to go. 

![](/images/L-Drill-Inventory/14.png)
<br>
  


**Step 4:** Now that you have the data, and can click the hyperlink. As shown in the animated GIF below, click on **Totals for Grains/Cereals** and INTERJECT will drill to the detail of that category in the target worksheet. Hyperlinks only show the Drill window when there is more than one drill option setup. In this case, you only setup one drill and it goes there automatically. 

![](/images/L-Drill-Inventory/15.gif)
<br>
  


Hyperlinking Drills is a simple way to make INTERJECT reports faster and more user-friendly. Click [ here ](/wGetStarted/L-Drill-TheThreeWays.html#the-hyperlink-method) for the Financial Report Drill walk-through. 

  


