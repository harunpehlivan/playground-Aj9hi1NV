# Introduction
Are you looking for a way to Read/Write Excel file without Interop com? Want to Read both XLS and XLSX format? Then read this article - it will really help you Read or Write Excel files using OLEDB.

## Background
In earlier days when I was new to programming, I used to read/write Excel file using Interop object, but it is unmanaged and heavy entity and due to its 'HELL' **performance**, I desperately needed some good alternative to Interop. I have gone through OLEDB, it performs very well for reading and writing Excel files.

## Using Code
Before start Reading/Writing from/in Excel file, we need to connect to OLEDB using connection string, here OLEDB will act as Bridge between your program and EXCEL.

![Bridge between C# and EXCEL](https://www.codeproject.com/KB/aspnet/1088970/Bridge.jpg "Bridge between C# and EXCEL")




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
