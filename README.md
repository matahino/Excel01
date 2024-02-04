I want to create an application without using Visual Studio, and I also want to manipulate Excel. <br>
So today we will operate using only the command line.

First, please download Excel. You can download it from Office 365.<br>
![image](https://github.com/matahino/Excel01/assets/96413690/739ef333-c371-4409-a8dc-07794760010a)<br>


Once downloaded, we will look for the library we will be using. <br>
A library is a file that translates programming languages into a format that is more understandable for machines.<br>
The content is the same as the programming language. This time, the library we are looking for is Microsoft.Office.Interop.Excel.dll. <br> 
Now, let's search the C drive with the where command, which we always rely on. <br> 
The syntax is: where /r c:\ <br>
After c:\, you enter the name of the file you want to find. <br>
So it will be: where /r c:\ "Microsoft.Office.Interop.Excel.dll"<br>
![image](https://github.com/matahino/Excel01/assets/96413690/b1558351-15b0-4213-b8ac-790dfd442d26)<br>


Next: Compile this program and library together as found.
![image](https://github.com/matahino/Excel01/assets/96413690/facbb0dc-507c-4269-984f-525ab86caacd)<br>
Compile using the csc command
When adding library, /r:Microsoft.Office.Interop.Excel.dll
The resulting command
csc 3.cs /r:Microsoft.Office.Interop.Excel.dll
![image](https://github.com/matahino/Excel01/assets/96413690/f3334f60-fbd9-4c80-9ec4-7601c13fde32)
When you run the application.
![image](https://github.com/matahino/Excel01/assets/96413690/ae4d87f8-fc71-4bbb-8885-19ce0338c525)





