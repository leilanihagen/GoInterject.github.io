Start with a blank Excel workbook:

![](/images/L-Dev-Report_from_Scratch/01.png)

### Formatting the Behind the Scenes Section

We’ll start by setting up the titles of the sections that hold the different report functionality formulas. This formatting is standard to INTERJECT reports.

**Step 1:** Start by selecting row 1 and coloring it dark blue (#1F4E78). This is the color that INTERJECT uses for titles of report definition sections.

1. First, click on the “1” that denotes row 1 to highlight the entire row.
2. Click the paint bucket to fill the color.
3. Choose the darkest blue (#1F4E78).

![](/images/L-Dev-Report_from_Scratch/02.png)

For this report, we will need 5 different titled sections. Now that you have the color selected in your paint bucket, simply click on every other row and then click on the paint bucket until you have 5 dark blue rows:

![](/images/L-Dev-Report_from_Scratch/03.png)

Now let’s name the title sections. We will need names: “Column Definitions,” “Formatting Range,” “Report Formulas,” “Hidden Parameters and Notes,” and “Report Area Below.”  We will enter “Column Definitions” and make it **white** and **bold** as follows:

![](/images/L-Dev-Report_from_Scratch/04.png)

Now enter the names “Formatting Range,” “Report Formulas,” “Hidden Parameters and Notes,” and “Report Area Below” in the next 4 title rows. Don’t worry about the formatting of these 4 for now.

![](/images/L-Dev-Report_from_Scratch/05.png)

Now, we can use the format painter to copy the formatting of the first title to the remaining 4. 

1. Select **row 1**,
2. click the **format painter**,
3. click **row 3**.

![](/images/L-Dev-Report_from_Scratch/06.png)

Repeat for **rows 5, 7 and 9.**

As you may have noticed, the jFreezePanes() is out of place. And our hidden freeze panes sections goes all the way down to row 17, so the space where our titles are laid out should occupy all of this space. Let’s insert some more empty rows under our titles to put more space for formula definitions and the like.

Copy two empty rows from somewhere in the sheet:

![](/images/L-Dev-Report_from_Scratch/07.png)

Paste them above row 2 by right clicking on row 2:

![](/images/L-Dev-Report_from_Scratch/08.png)

Now copy and paste 2 more rows under each title so that your report looks like this:

![](/images/L-Dev-Report_from_Scratch/09.png)

Now let’s add the standard light blue color to the titled sections:

    1. Select the 3 rows under Column Definitions.
    2. Click the paint bucket.
    3. Select the lightest blue color (#DDEBF7).

![](/images/L-Dev-Report_from_Scratch/10.png)

Repeat this step for the three other report definition areas by **selecting each area** and **clicking the paint bucket** (it should already have your previously selected color in it.

![](/images/L-Dev-Report_from_Scratch/11.png)

### Setting the Freeze Panes

Freeze panes are important for INTERJECT reports because they allow us to a) hide the report definitions section to ensure that we aren’t confusing our end users with details that they do not need to see and b) keep a header with column titles visible to the user as they scroll through report data.

jFreezePanes() is an INTERJECT function that takes advantage of the Excel native Freeze/Unfreeze panes option in the Quick Tools menu. The jFreezePanes() function allows us to specify a) which worksheets in a workbook will be frozen (whichever worksheets have the jFreezePanes() function in their Report Formulas section) and b) where to freeze the panes in the workbook (which cells to freeze between). [Read more about jFreezePanes() here](https://docs.gointerject.com/wIndex/jFreezePanes.html).

We will start off by setting the Freeze Panes at the correct location.

**Step 1:** To start, type “=jFreezePanes” in cell **F10**:

![](/images/L-Dev-Report_from_Scratch/12.png)

Click on the function builder(?):

![](/images/L-Dev-Report_from_Scratch/13.png)

**Step 2:** There are two parameters of jFreezePanes(), FreezePanesCell and AnchorViewCell. AnchorViewCell specifies the very top row that will be visible when the panes are frozen. The cells above AnchorViewCell will be hidden when the panes are frozen. The cells between AnchorViewCell and FreezePanesCell is the block that is frozen at the top of the sheet as you scroll down the sheet.

Set **FreezePanesCell = A26**:

![](/images/L-Dev-Report_from_Scratch/14.png)

Then set **AnchorViewCell = A18**:

![](/images/L-Dev-Report_from_Scratch/15.png)

Now that we have our freeze pane set up, we can start with formatting the spreadsheet.

INTERJECT uses the hidden area of the frozen pane to define INTERJECT report functions and to set up the formatting of the report.

Let’s try freezing the panes to see how it works.

1. Press and hold **CTRL + SHIFT + T** or click on **Quick Tools** to open the Quick Tools menu.
2. Press **Enter** or click **Freeze/Unfreeze Panes (current tab)**.

![](/images/L-Dev-Report_from_Scratch/16.png)

Your report should now look like the following. The sectioned off block from rows 18-25 (ends at highlighted line) is the frozen pane section that will stay at the top as you scroll down. This is where our header with the name of the report and filter parameters will go later. The cells above row 18, which contain our report definitions area, are hidden.

![](/images/L-Dev-Report_from_Scratch/17.png)

### Formatting the Report Area

Now let’s format the report area. We’ll start by putting a report title in cell **B19** “Customer Orders” and formatting it to be **bold** and of text **size 14**.

![](/images/L-Dev-Report_from_Scratch/18.png)

Next, let’s name the report filters for this report. The report filters act as a way to specify which data is being pulled into the report from the data portal by specifying a set of characters that the pulled in data must contain. In cells **B21, B22 and B23**, respectively, type in: **“Company Name:”**, **“Contact Name:”**, and **“Customer ID:”**

![](/images/L-Dev-Report_from_Scratch/19.png)

Now, let’s resize column A to be smaller, and extend column B and C by a bit.

1. Drag column A back.
2. Drag column B forward.
3. Drag column C forward.

![](/images/L-Dev-Report_from_Scratch/20.png)

Now let’s color the input fields for the report filters. Apply the lightest orange color () to cells **C21, C22 and C23**:

![](/images/L-Dev-Report_from_Scratch/21.png)

Now let’s right-align **cells C21-C23**,

1. Select cells **C21-C23**.
2. Select right-align.

![](/images/L-Dev-Report_from_Scratch/22.png)

Now that we’ve titled and styled our report filters, let’s make the spreadsheet look better by removing the gridlines in Excel.

Go to the **File** tab in Excel:

![](/images/L-Dev-Report_from_Scratch/23.png)

Go to **Options**:

![](/images/L-Dev-Report_from_Scratch/24.png)

    1. Go to the **Advanced** tab.
    2. Scroll down until you see **“Display options for this worksheet”**.
    3. **Uncheck** the **“Show gridlines”** checkbox.

![](/images/L-Dev-Report_from_Scratch/25.png)

Now there should be no gridlines on the current worksheet.

Let’s rename the current worksheet **CustomerOrderHistory** and delete any other worksheets you have in the workbook.

Right-click on the first sheet’s title and select **Rename**.

![](/images/L-Dev-Report_from_Scratch/26.png)

Type in **CustomerOrderHistory**.

![](/images/L-Dev-Report_from_Scratch/27.png)

Right-click on any other sheets that you have in the workbook and select **Delete**.

![](/images/L-Dev-Report_from_Scratch/28.png)

**Step #:** Let’s change the font on our entire worksheet to be **Century Gothic**.

Select all the cells in the sheet by clicking the tab in the top left corner.

![](/images/L-Dev-Report_from_Scratch/29.png)

Then type **Century Gothic** into the font selector in the **Home** tab at the top.

![](/images/L-Dev-Report_from_Scratch/30.png)

You can reduce the scale of the worksheet to your monitors needs. Since the new font is a little bigger, we’ll reduce the demo worksheet to 90%.

31 (32 and 33 deleted, so shift 34->32 etc)

### Adding ReportRange() to the Report

**Step 1:** Let’s add our first INTERJECT report formula to the report. We’ll start with **ReportRange()**. ReportRange() is a report formula used to PULL data into a defined *range* of a report from the Data Portal. ReportRange() can be used with formatting to format the data returned from the Data Portal into the spreadsheet. Read more about ReportRange() [here](https://docs.gointerject.com/wIndex/ReportRange.html#function-summary).

Type **=ReportRange()** in cell **C10** then click on the function builder icon.

![](/images/L-Dev-Report_from_Scratch/32.png)

As you can see, DataPortal is the first parameter that will must provide ReportRange() so that it knows where to pull in the data from. Type **NorthwindCustomerOrders_MyName** into the DataPortal parameter box for now.

![](/images/L-Dev-Report_from_Scratch/33.png)

We will now switch to configuring an INTERJECT Data Connection, and a Data Portal that we can pull from using ReportRange().

We will fill in the rest of the ReportRange() parameters once we have set up our Data Portal and Connection.

### Setting Up the Data Connection

In order to continue our work from here, we need to set up the back-end Data Portal that ReportRange() will be using. For now, we will pause working on the front-end Excel report to configure the Data Portal and Data Connection that ReportRange() will use in our report.

We’ll start with the Data Connection. INTERJECT Data Connections enable users to connect to a database in order to pull data out of that database based on criteria specified in stored procedures which are set up with Data Portals. An overview of Data Connections and Data Portals can be found [here](https://docs.gointerject.com/wPortal/The-INTERJECT-Website-Portal.html#overview).

**Step 1: Logging in** Start by navigating to the INTERJECT portal site ([here](https://portal.gointerject.com/)) and logging in.

![](/images/L-Dev-Report_from_Scratch/34.png)

**Step 2: Create the connection:** Create a new data connection by clicking the New Connection button.

![](/images/L-Dev-Report_from_Scratch/35.png)

Name you connection **NorthwindDB_MyName** (substitute for your name) and give it a quick description.

![](/images/L-Dev-Report_from_Scratch/36.png)

Select **Database** from the dropdown list for **Connection Type**.

![](/images/L-Dev-Report_from_Scratch/37.png)

For the connection string, you will need to have your own sample Northwind database to use. You can download a Northwind sample database from Microsoft [here](https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sql/linq/downloading-sample-databases).

Substitute in your server and database name in italicized parts of the following sample connection string: ”Server=*MyServerAddress*;Database=*MyDatabase*;Trusted_Connection=True;” Once you have your connection string entered, press Save to continue.

![](/images/L-Dev-Report_from_Scratch/38.png)

### Setting Up the Data Portal

**Step 1: Create the Data Portal** Now, we will create the Data Portal that allows us to actually pull data from the Data Connection that we made above.

Data Portals are provided as a way to connect to specific stored procedures within the Data Connection to an existing database. It is a finer-grain level of control, and connects to a single stored procedure on the database you connect to through the provided Data Connection. You can have multiple Data Portals connected to one Data Connection, but not vice-versa. For more, see [the website portal documentation](https://docs.gointerject.com/wPortal/The-INTERJECT-Website-Portal.html#-data-connections-).

Navigate again to [the portal site](https://portal.gointerject.com/) and choose Data Portals.

![](/images/L-Dev-Report_from_Scratch/39.png)

Create a new data portal.

![](/images/L-Dev-Report_from_Scratch/40.png)

Start by naming your Data Portal **”NorthwindCustomerOrders_YourName”** (substitute in your name) and giving it a description.

![](/images/L-Dev-Report_from_Scratch/41.png)

For the **Connection**, we will use the Data Connection we created in the last section, **NorthwindDB_YourName**. It should appear in the dropdown list when clicked

![](/images/L-Dev-Report_from_Scratch/42.png)

Now we will specify the stored procedure that this data portal will be referencing. We will write the stored procedure itself shortly. Name your stored procedure **”[demo].[northwind_customer_orders_myname]”**.

![](/images/L-Dev-Report_from_Scratch/43.png)

For the **Category**, enter **Demo** and for the **Command Type**, choose **Stored Procedure Name** from the dropdown list.

![](/images/L-Dev-Report_from_Scratch/44.png)

Make sure **Data Portal Status** is set to **Enabled** and **Is Custom Command?** is set to **No**, then save the new Data Portal:

![](/images/L-Dev-Report_from_Scratch/45.png)

Now we will add our formula parameters. If you look at the report, you will remember we have 3 filters on our report, **Company Name**, **Contact Name** and **Customer ID**. Our Data Portal and stored procedures need to know that we have these filter parameters to affect the data that they pull out.

It will be important later on, when writing the stored procedure, that the order the parameters are listed in the Data Portal is the same as their order listed in the stored procedure. Because of this, we recommend keeping the order consistent between all three platforms, the report, the Data Portal and the stored procedure, to minimize opportunities for confusion. We will use the order of the filter parameters from the report everywhere and recommend you do the same.

Parameter order of filters in report:

![](/images/L-Dev-Report_from_Scratch/46.png)

Click on the **Click here to add a Formula Parameter** link and enter the first parameter name **CompanyName**. This will directly reference a parameter that we will code into our stored procedure with the same name later on.

![](/images/L-Dev-Report_from_Scratch/47.png)

Set the **TYPE** to **nvarchar** so that a character string can be entered by the user, and set the **DIRECTION** to **input** since this will be an input parameter and not an output.

![](/images/L-Dev-Report_from_Scratch/48.png)

Press the save button so that it turns from red to green.

![](/images/L-Dev-Report_from_Scratch/49.png)

Add the next Formula Parameter for **ContactName**.

![](/images/L-Dev-Report_from_Scratch/50.png)

Add the last Formula Parameter for **CustomerID**.

![](/images/L-Dev-Report_from_Scratch/51.png)

Now we’ll add the necessary System Parameters. System Parameters are used to pass information from the user’s system to the stored procedure/Data Portal. Here, we will be adding 2 System Parameters, **Interject_NTLogin**, which is used to capture the user’s Windows login, and **Interject_LocalTimeZoneOffset**, which is used to capture the difference from the user’s local time zone to the universal time. You can read more about System Parameters (and these specific ones) [here](https://docs.gointerject.com/wGetStarted/L-Dev-CustomerAging.html#system-parameters).

Create a new System Parameter and choose **Interject_NTLogin** from the dropdown menu.

![](/images/L-Dev-Report_from_Scratch/52.png)

Press the save button and wait until it turns green.

![](/images/L-Dev-Report_from_Scratch/53.png)

Add a second System Parameter, choose **Interject_LocalTimeZoneOffset** from the dropdown menu and make sure you save the new parameter.

![](/images/L-Dev-Report_from_Scratch/54.png)

Verify that you have all your parameter information correct and that you have saved them all before moving on. Your screen should look as follows.

![](/images/L-Dev-Report_from_Scratch/55.png)

### Setting up ReportRange() with the Data Portal

Now, we have a Data Connection to a database, and a Data Portal which specifies a stored procedure to provide data to it; but we need to write the stored procedure in order to actually get anything back from our ReportRange() call in the report.

In order to show how the front-end Excel interface ties into the writing of the back-end stored procedure, let’s start by going back to the report and figuring out what data we want to display to the user.

**Step 1:** Go back to the report, click in cell **C10** and open the function builder.

![](/images/L-Dev-Report_from_Scratch/56.png)

Enter **2:4** into the **ColDefRange** to tell ReportRange() that all of its column definitions can be found in this range of rows. You can read more about ColDefRange here.

![](/images/L-Dev-Report_from_Scratch/57.png)

Now, we can specify the columns that we want to get back from our Data Portal via ReportRange() in the Column Definitions section of our report. Let’s fill these values in.

Starting with row 2, type **CustomerID** into cell **B2**, **CompanyName** into cell **C2**, **ContactName** into cell **E2**, **OrderID** into cell **F2**, **OrderDate** into cell **G2**, **OrderAmount** into cell **H2**, **Freight** into cell **I2**, **TotalAmount** into cell **J2**.

![](/images/L-Dev-Report_from_Scratch/58.png)

In row 3, we just need **ShipVia** in cell **C3** and **ShippedDate** in cell **E3**.

![](/images/L-Dev-Report_from_Scratch/59.png)

Now, let’s add the other parameters. Open the function arguments for ReportRange() again.

![](/images/L-Dev-Report_from_Scratch/60.png)

ReportRange() works by inserting the result set returned from the Data Portal *in between* two or more rows. These rows are specified by the TargetDataRange argument. Input **27:28** for our **TargetDataRange** (down in the Report Area, below the filter parameters).

![](/images/L-Dev-Report_from_Scratch/61.png)

The Formatting Range is the part of the report definitions section that specifies how final output will be formatted when returned to the end user. Our formatting range occupies rows 6:8, so input **6:8** in **FormatRange**.

![](/images/L-Dev-Report_from_Scratch/62.png)

The **Parameters** parameter specifies which cells will be the “filter” cells whose values are sent to the Data Portal to filter results to the user’s specifications. The Param() function ([read more here](https://docs.gointerject.com/wIndex/Param.html)) is used here to capture the cells. Type **Param(C21,C22,C23)** into **Parameters**.

![](/images/L-Dev-Report_from_Scratch/63.png)

As a best practice, we recommend you set **UseEntireRow** to **TRUE** and **PutFieldNamesAtTop** to **FALSE**

![](/images/L-Dev-Report_from_Scratch/64.png)

### Writing the SQL Stored Procedure Behind ReportRange()

Using a SQL editor like [SQL Server Management Studio](https://docs.microsoft.com/en-us/sql/ssms/sql-server-management-studio-ssms?view=sql-server-2017), copy and paste in the following code:

code

Here is the SELECT statement in the code. The columns returned from the SELECT statement are the ones that populate into the report.

## (screenshot including both SELECT and column definition area)

Save your stored procedure, making sure that it’s name matches the name you specified for it in the Data Portal you created and that it is in the same database that you specified in the Data Connection you created.

### Designing the Formatting Range

Because we have multiple rows in our Column Definition for ReportRange(), we need a Formatting Range to specify how each of these fields will look in the output from the pull.

Formatting Ranges work by letting you define the formatting you’d like to apply to your output data (specified in the Column Definition area); they let you do it concisely, in one place, such that the formatting can be copied down and repeated for each data record set pulled from the Data Portal. You don’t need to design a Formatting Range for every report you will write, but when you are using report formulas that trigger a pull action, you need a Formatting Range if you have more than one row in your column definition. If you have only one row, and you don’t specify a formatting range, the formatting of the first row in the TargetDataRange will be copied to the output rows.

**Step 1:** Let’s start by making the cells in our Formatting Range white. Select **rows 6-8** and select **white** from the dropdown list of paint bucket colors.

![](/images/L-Dev-Report_from_Scratch/65.png)

When designing Formatting Ranges, we use contrived sample data to illustrate how the real data will look in the TargetDataRange when it’s pulled in.

**Step 2:** First, we’ll format how we want **CustomerID** to look in the output data. Enter the sample ID **GREAL** into cell **B6**.

![](/images/L-Dev-Report_from_Scratch/66.png)

For **CompanyName**, enter **Great Lakes Food Market** into cell **C6** and make it bold.

1. Enter **Great Lakes Food Market** into cell **C6**.
2. Select the text in the cell.
3. Apply bold to the selected text.

![](/images/L-Dev-Report_from_Scratch/67.png)

For **ContactName**, enter **Howard Snyder** into cell **E6**.

![](/images/L-Dev-Report_from_Scratch/68.png)

For **OrderID**:

1. Enter **11061** into cell **F6**.
2. Select the text and make it bold.
3. Click the center-align text button.
4. Click the paint bucket.
5. Select the lightest grey color.

![](/images/L-Dev-Report_from_Scratch/69.png)

For **OrderDate**:

1. Enter the sample date **4/30/98** in cell **G6**.
2. Enter **Date** in the format options for the cell.

![](/images/L-Dev-Report_from_Scratch/70.png)

For **OrderAmount**:

1. Enter the sample data **510** in cell **H6**.
2. Enter **Accounting** in the format options for the cell.

![](/images/L-Dev-Report_from_Scratch/71.png)

For **Freight**, enter **14.01** into cell **I6** and change the format for the cell to **Accounting**.

![](/images/L-Dev-Report_from_Scratch/72.png)

For **TotalAmount**, enter **524.01** into cell **J6**, make the text bold, and again change the format for the cell to **Accounting**.

![](/images/L-Dev-Report_from_Scratch/73.png)

We now only have to format the cells for **ShipVia** and **ShippedDate**. We will make titles for these fields in the row to the left of them and leave the values themselves without formatting.

Enter **Shipped Via:** in cell **B7** and **Ship Date:** in cell **D7**. Also expand column C by a bit.

![](/images/L-Dev-Report_from_Scratch/74.png)

Now, we want to add a border under row 7 (at the top of row 8) to demarcate the end of each record set.

Select cells **B8-J8**.

![](/images/L-Dev-Report_from_Scratch/75.png)

Click on the **Borders** dropdown menu and choose **Top Border**.

![](/images/L-Dev-Report_from_Scratch/76.png)

Lastly, we’ll reduce row 8 to provide a small padding under the border we just added.

![](/images/L-Dev-Report_from_Scratch/77.png)

### Testing ReportRange()

Now that we’ve set up ReportRange() with all of its arguments, let’s test it. It is always good practice to test each formula you add to the report you’re building once it’s done, before you move on to building the next part/formula on the report. This ensures that at the end when you’re ready to test the finished report, you know that all the constituent parts work by themselves.

**Step 1:** Enter **market** into the Company Name filter parameter in cell **C21**. This will filter the result set in our SQL query that we wrote, only selecting the records whose CompanyName column contains the string ”market.” Providing a filter is helpful for 2 reasons: 1) it reduces the amount of data you are requesting back from the database which reduces the execution speed of the data pull, and 2) it helps test your query to see if it selected all of the expected data records.

![](/images/L-Dev-Report_from_Scratch/78.png)

**Step 2:** Now run a data PULL on the report.

1. Press **CTRL + SHIFT + J** together on your keyboard or click the **PULL Data** button.
2. Press **Enter** or click **Pull Data**.

![](/images/L-Dev-Report_from_Scratch/79.png)

**Step 3:** Unfreeze the panes to view the data how the end user would.

1. Press **CTRL + SHIFT + T** together on your keyboard or click the **Quick Tools** button.
2. Press **Enter** or click **Freeze/Unfreeze Panes (current tab)** in the menu.

![](/images/L-Dev-Report_from_Scratch/80.png)

Your data should look like the following.

![](/images/L-Dev-Report_from_Scratch/81.png)

Now that we know which pieces of data we need in our report, **we can design the stored procedure**.

### Setting up ReportDefaults()

The ReportDefaults() function is used to capture values from one or a set of cells (or an independently specified value) and send the value/s to another cell or set of cells. Its execution is triggered based on an action or event (read the distinction between and INTERJECT action/event [here](https://docs.gointerject.com/wIndex/ReportDefaults.html#trigger-combination-list)) happening in the report (for example a save or clear action). ReportDefaults() is commonly used to clear values in the filter list after data has been pulled in and then cleared, which is how it will be used in our report. Read more about ReportDefaults() [here](https://docs.gointerject.com/wIndex/ReportDefaults.html#function-summary).

In this report, we will be using ReportDefaults to clear out the filter values in cells C21-C23 after a CLEAR is run on the report. CLEAR does not do this by default, because it’s scope of control over the report is limited to the *results* of the data pull (CLEAR is only allowed to modify the data that a PULL action brings in, because a CLEAR reverses a PULL).

**Step 1:** Start by typing **=ReportDefaults()** into cell **C11** and opening the function builder.

![](/images/L-Dev-Report_from_Scratch/82.png)

For the arguments OnPullSaveOrBoth and OnClearRunOrBoth, we want ”Pull” and ”Clear” respectively. This is because we want to execute “Trigger 2” explained in [the ReportDefaults() documentation](https://docs.gointerject.com/wIndex/ReportDefaults.html#trigger-combination-list). With these arguments, the defaults will trigger when the user performs a Pull-Clear event-action sequence.

Enter **”Pull”** into the **OnPullSaveOrBoth** field, and **”Clear”** into **OnClearRunOrBoth**.

![](/images/L-Dev-Report_from_Scratch/83.png)

The TransferPairs argument is what decides the default values to place in the selected cells. We want to clear out our filter parameter cells, so we will pair each of these cells (C21-C23) with a blank value to pass in when a Pull-Clear occurs.

Each of these pairs (a cell with a blank string (“”)) needs its own Pair() function inside the overarching PairGroup() function that ReportDefaults() takes as a parameter. Learn more about PairGroup() and Pair()  [here]().

Input **PairGroup()** into **TransferPairs**, then press **Ok**.

![](/images/L-Dev-Report_from_Scratch/84.png)

Now, to give arguments to the PairGroup() function, click inside the PairGroup() function within ReportDefaults() and open the function builder.

1. Click in cell C11.
2. In the **Formula Bar**, place the cursor somewhere in the function name text “PairGroup().”
3. Click the function builder icon.

![](/images/L-Dev-Report_from_Scratch/85.png)

The parameters to Pair() are **From, Target**, or, if you do not have a “From” cell, as in our case, they can be thought of as **SourceValue, Target**. Our “Source Value” is “” (an empty string) and our Targets are cells C21-C23

Type **Pair(””, C21)** into **Pair1**, **Pair(””, C22)** into **Pair2** and **Pair(””, C23)** into **Pair3**.

![](/images/L-Dev-Report_from_Scratch/86.png)

Your ReportDefaults() function should now look like the following:

![](/images/L-Dev-Report_from_Scratch/87.png)

### Testing ReportDefaults()

Now let’s test our ReportDefaults() function. It should clear any filter arguments from cells C21-C23 after a PULL-CLEAR.

**Step 1:** Enter some filters into the filter fields. Enter **market** into cell **C21**, and **b** into cell **C22**.

88 NAME CHANGE

**Step 2:** PULL in the data.

![](/images/L-Dev-Report_from_Scratch/89.png)

**Step 3:** CLEAR the data.

1. Press **CTRL + SHIFT + J** or click the **Quick Tools** menu button.
2. Press **Down Arrow** then **Enter** or click **Clear**.

![](/images/L-Dev-Report_from_Scratch/90.png)

Now you should see that, as well as the data in the Target Data Range being cleared, the filter values in cells C21-C23 will also clear out.

![](/images/L-Dev-Report_from_Scratch/91.png)

### Introducing the SalesOrder Report and Drilling Between Reports

We will create a second report so that we can demonstrate *drilling* between reports.

Drilling is a way to connect and pass values between separate worksheets or workbooks. A typical use case is when you have a general report that provides a summary of some data, you may want to allow the user to get more detail on certain data, and this can be accomplished by setting up a drill into a more detailed report. You can then pass some piece of data from the general/summary report into a cell in the detailed report so that the detailed report can automatically pull in and filter data based on the cell the user drilled on. You can read more about ReportDrill() [here]().

In our case, SalesOrderHistory, the report we’ve been working on so far, is the general/summary report. We will create a new report, SaleOrder, which will be the detailed report that we will drill into from SalesOrderHistory.

### Creating the SalesOrder Report

**Step 1:** First, let’s create another worksheet in our workbook and name it SalesOrder.

Click the plus sign to add another worksheet.

![](/images/L-Dev-Report_from_Scratch/92.png)

Right-click on the new worksheet and select **Rename**.

![](/images/L-Dev-Report_from_Scratch/93.png)

Enter **SalesOrder** in the input field.

![](/images/L-Dev-Report_from_Scratch/94.png)

**Step 2:** Let’s start by again turning off gridlines in this workbook.

Click into the **File** tab above the Excel ribbon.

![](/images/L-Dev-Report_from_Scratch/95.png)

Click **Options**

![](/images/L-Dev-Report_from_Scratch/96.png)

1. In the window that pops up, click the **Advanced** tab.
2. Scroll down until you see the header **Display options for this worksheet**.
3. **Uncheck** the **Show gridlines** box.

![](/images/L-Dev-Report_from_Scratch/97.png)










LATER:

**Step 1:** Let’s start by entering **=ReportDrill()** into cell **C12**.

##


