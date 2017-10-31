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



This C# template lets you get started quickly with a simple one-page playground.

```C# runnable
// { autofold
using System;

class Hello 
{
    static void Main() 
    {
// }

Console.WriteLine("Hello World!");

// { autofold
    }
}
// }
```

# Advanced usage

If you want a more complex example (external libraries, viewers...), use the [Advanced C# template](https://tech.io/select-repo/386)
