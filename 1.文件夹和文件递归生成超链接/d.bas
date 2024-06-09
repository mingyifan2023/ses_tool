Sub test()
    On Error Resume Next
    Dim arr(1 To 10000) As String
    Dim f, i, k, f2, f3, x
    Dim arr1(1 To 100000, 1 To 7) As String, q As Integer
    Dim fso As Object, myfile As Object
    arr(1) = Application.InputBox("Please enter the path") & "/"
    i = 1: k = 1
    Do While i < UBound(arr)
        If arr(i) = "" Then Exit Do
        f = Dir(arr(i), vbDirectory)
        Do
            If InStr(f, ".") = 0 And f <> "" Then
                k = k + 1
                arr(k) = arr(i) & f & "\"
            End If
            f = Dir
        Loop Until f = ""
        i = i + 1
    Loop
    '******* Below is to extract files from each folder ***
    Set fso = CreateObject("Scripting.FileSystemObject")
    For x = 1 To UBound(arr)
        If arr(x) = "" Then Exit For
        f3 = Dir(arr(x) & "*.*")
        Do While f3 <> ""
            If InStr(f3, ".") > 0 Then
                q = q + 1
                arr1(q, 5) = arr(x) & f3
                Set myfile = fso.GetFile(arr1(q, 5))
                arr1(q, 1) = f3
                arr1(q, 2) = myfile.Size
                arr1(q, 3) = myfile.DateCreated
                arr1(q, 4) = myfile.DateLastModified
                arr1(q, 6) = myfile.DateLastAccessed
                ' Add hyperlink for the file
                Cells(q + 1, 7).Formula = "=HYPERLINK(""" & arr(x) & f3 & """, """ & f3 & """)" ' Actual file hyperlink
            End If
            f3 = Dir
        Loop
    Next x
    Sheets("Sheet2").Range("a2").Resize(1000, 7) = ""
    Sheets("Sheet2").Range("a2").Resize(q, 7) = arr1

    ' Add hyperlinks to the files
    Dim temp_Address As String
    Dim temp_anchor As Range ' Declare range variable
    For j = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        Set temp_anchor = Cells(j, 7) ' Anchor the position
        temp_Address = Cells(j, 5) ' Specific hyperlink address
        ActiveSheet.Hyperlinks.Add Anchor:=temp_anchor, Address:=temp_Address ' Correct to specify anchor and address when adding hyperlink
    Next j
End Sub

