I want to create an application without using Visual Studio, and I also want to manipulate Excel. <br>
Therefore, I will be performing Excel operations today using only command line with the target set to winexe.

Create an app at the command prompt: https://github.com/matahino/Windowsforms01/releases/tag/C%231

First, please download Excel. You can download it from Office 365.<br>
![image](https://github.com/matahino/Excel01/assets/96413690/739ef333-c371-4409-a8dc-07794760010a)<br>


Once downloaded, we will look for the library we will be using. <br>
A library is a file that translates programming languages into a format that is more understandable for machines.<br>
The content is the same as the programming language. This time, the library we are looking for is Microsoft.Office.Interop.Excel.dll. <br> 
Now, let's search the C drive with the where command, which we always rely on. <br> 
The syntax is: where /r c:\ <br>
After c:\, you enter the name of the file you want to find. <br>
So it will be: where /r c:\ "Microsoft.Office.Interop.Excel.dll"<br>
![image](https://github.com/matahino/Excel01/assets/96413690/0e3f1713-2c5b-4fc2-9f55-1bc89445d260)



