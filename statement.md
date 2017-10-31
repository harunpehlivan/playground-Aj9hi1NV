# Introduction
Are you looking for a way to Read/Write Excel file without Interop com? Want to Read both XLS and XLSX format? Then read this article - it will really help you Read or Write Excel files using OLEDB.

## Background
In earlier days when I was new to programming, I used to read/write Excel file using Interop object, but it is unmanaged and heavy entity and due to its 'HELL' **performance**, I desperately needed some good alternative to Interop. I have gone through OLEDB, it performs very well for reading and writing Excel files.

## Using Code
Before start Reading/Writing from/in Excel file, we need to connect to OLEDB using connection string, here OLEDB will act as Bridge between your program and EXCEL.

![Bridge between C# and EXCEL](https://www.codeproject.com/KB/aspnet/1088970/Bridge.jpg "Bridge between C# and EXCEL")

Rows and columns of Excel sheet can be directly imported to data-set using OLEDB, no need to open Excel file using INTROP EXCEL object.

Let's start with the code.

```javascript
// Connect EXCEL sheet with OLEDB using connection string
// if the File extension is .XLS using below connection string
//In following sample 'szFilePath' is the variable for filePath
 szConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;
                       "Data Source='" + szFilePath + 
                       "';Extended Properties=\"Excel 8.0;HDR=YES;\"";
 
 // if the File extension is .XLSX using below connection string
 szConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;
                      "Data Source='" + szFilePath + 
                      "';Extended Properties=\"Excel 12.0;HDR=YES;\"";
```

In the above connection string:

1. Provider is OLEDB provider for Excel file, e.g., Jet.OLEDB.4.0 is for XLS file and ACE.OLEDB.12.0 for XLSX file
2. Data Source is the file path of Excel file to be read
3. Connection string also contains 'Extended Properties' like Excel driver version, HDR Yes/No if source Excel file contains first row as header

After connection to EXCEL file, we need to fire Query to **retrieve** records from sheet1.

## Accessing Excel Tables

There are a couple of ways to refer to an Excel table:

1. Using sheet name: With the help of sheet name, you can refer to Excel data, you need to use '$' with sheet name, 

e.g. *Select * from [Sheet1$]*

2. Using Range: We can use Range to read Excel tables. It should have specific address to read, 

e.g. *Select * from [Sheet1$B1:D10]*

**Here $ indicates the EXCEL table/sheet already exists in workbook, if you want to create a New workbook/sheet, then do not use $, look at the sample below:

``` javascript
// Connect EXCEL sheet with OLEDB using connection string
 using (OleDbConnection conn = new OleDbConnection(connectionString))
    {
        conn.Open();
        OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter
        ("select * from [Sheet1$]", conn);
        DataSet excelDataSet = new DataSet();
        objDA.Fill(excelDataSet);
        dataGridView1.DataSource = excelDataSet.Tables[0];
    }
			
	//In above code '[Sheet1$]' is the first sheet name with '$' as default selector,
        // with the help of data adaptor we can load records in dataset		
	
	//write data in EXCEL sheet (Insert data)
 using (OleDbConnection conn = new OleDbConnection(connectionString))
    {
        try
        {
            conn.Open();
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = @"Insert into [Sheet1$] (month,mango,apple,orange) 
            VALUES ('DEC','40','60','80');";
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            //exception here
        }
        finally
        {
             conn.Close();
             conn.Dispose();
        }
    }
			
//update data in EXCEL sheet (update data)
using (OleDbConnection conn = new OleDbConnection(connectionString))
	{
        try
        {
            conn.Open();
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = "UPDATE [Sheet1$] SET month = 'DEC' WHERE apple = 74;";
            cmd.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            //exception here
        }
        finally
        {
            conn.Close();
            conn.Dispose();
        }
    }
```

*OLEDB does not support DELETE query.

![Media](https://www.codeproject.com/KB/aspnet/1088970/Media.jpg "Media")

### Exceptions, you might faced

1. The 'Microsoft.Jet.OLEDB.4.0' provider is not registered on the local machine.
**Cause:** The exception occurs when we run our code on 64Bit machine.

**How to Resolve:** If your application is Desktop based, compile your EXE with x86 CPU. If your application is web based, then Enable '32-Bit Applications' in application pool.

2. Deleting data in a linked table is not supported by this ISAM.
**Cause:** As we have already discussed, OLEDB does not support DELETE operation. If you try to Delete rows from EXCEL sheet, it gives you such exception.

## Advantage against INTEROP/COM object

We know EXCEL Interop application can also be used to complete this task, but there are several advantages against INTEROP/COM object, see the below points:

1. Interop objects are heavy and un-managed objects
2. Special permissions are needed to launch component services if you run this code as Web application in IIS
3. No Excel installation is needed when we need to Read/Write Excel using OLEDB. 4. OLEDB is faster in performance than Interop object, as No EXCEL object is created.

## Finally
There are always two sides of the coin. With OLEDB, you cannot format data that you inserted/updated in EXCEL sheet but Interop can do it efficiently. You cannot perform any mathematical operation or working on graphs using OLEDB, but it is really a good way to insert/update data in EXCEL where no Excel application is installed.

Comments and suggestions are always welcome

Thank you!