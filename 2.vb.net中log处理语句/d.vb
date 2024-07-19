Imports System.IO
一旦添加了这个引入语句，您就可以在 VB 文件中直接使用 File 类来写入 log 文件了。示例代码如下：

Imports System.IO

Module Program
    Sub Main(args As String())
        File.AppendAllText("C:\path\to\your\log\file.txt", "This is the value to log."& vbCrLf)
    End Sub
End Module


简单直接引入 Imports System.IO
然后直接使用下面语句即可输出log
File.AppendAllText("C:\path\to\your\log\file.txt", "This is the value to log." & vbCrLf )
