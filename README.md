I want to create an application without using Visual Studio, and I also want to manipulate Excel. <br>
Therefore, I will be performing Excel operations today using only command line with the target set to winexe.

Create an app at the command prompt: https://github.com/matahino/Windowsforms01/releases/tag/C%231

まずは、Excelをダウンロードしてください。office360からダウンロードできます。<br>
![image](https://github.com/matahino/Excel01/assets/96413690/739ef333-c371-4409-a8dc-07794760010a)<br>
ダウンロードしたら今回使うlibraryを探します。libraryとは、プログラミング言語を機械に分かりやすく変換したファイル。内容は、プログラミング言語と同じです。
今回、探すlibraryは、Microsoft.Office.Interop.Excel.dllです。では、いつもお世話になっているコマンドのwhereコマンドでCドライブを探しましょう。
式: where /r c:\ 
これです。c:\ 後ろに今回探したいファイル名を記入します。
where /r c:\ Microsoft.Office.Interop.Excel.dll


