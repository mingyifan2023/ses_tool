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

如何用上面的log方法输出一个遍历的值还有这个遍历的类型



如果您想要将遍历的值和类型输出到日志文件中，您可以使用File.AppendAllText方法和字符串插值（String Interpolation）。以下是一个示例代码：

Imports System.IO

Module Program
    Sub Main(args As String())
        Dim logFilePath As String = "C:\path\to\your\log\file.txt"
        
        Dim values As Integer() = {1, 2, 3, 4, 5}
        For Each value In values
            File.AppendAllText(logFilePath, $"Value: {value}, Type: {value.GetType().Name}" & vbCrLf)
        Next
    End Sub
End Module
