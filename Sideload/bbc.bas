Attribute VB_Name = "bbc"
Declare Function DLLencode Lib "C:\Users\owner\Documents\Visual Studio 2013\Projects\SideloadDLL2\Debug\SideloadDLL2.dll" (ByVal filein As String, ByVal fileout As String) As Integer
Declare Function DLLdecode Lib "C:\Users\owner\Documents\BlackHat Tests\newdll\SideloadDLL2.dll" (ByVal filesin As String, ByVal fileout As String, ByVal md5 As String, ByVal filesize As Long) As Integer


Sub sldencodetest()
    Debug.Print "Start sldencodetest"
    Dim filein As String
    Dim fileout As String
    filein = "C:\Users\owner\Downloads\stuff.zip"
    fileout = "C:\Users\owner\Downloads\stuff.sld"
    Dim val As Integer
    val = DLLencode(filein, fileout)
    If val > 0 Then
        Debug.Print "Success"
    Else
        Debug.Print "Fail"
    End If
    
    Debug.Print "End"
End Sub

Sub slddecodetest()
    Debug.Print "Start decode"
    Dim filesin As String
    Dim fileout As String
    Dim filesize As Long
    Dim md5 As String
    ' ; delimited string of files, see dll code
    filesin = "C:\Users\owner\Documents\BlackHat Tests\pay140-1o3.png;C:\Users\owner\Documents\BlackHat Tests\pay140-2o3.png;C:\Users\owner\Documents\BlackHat Tests\pay140-3o3.png"
    fileout = "payload1.zip"
    filesize = 247088
    md5 = "a5a24c8a80c2a12e13dffdd91cbdf8cb"
    Dim val As Integer
    val = DLLdecode(filesin, fileout, md5, filesize)
    If val > 0 Then
        Debug.Print "Success"
    Else
        Debug.Print "Fail"
    End If
    
    Debug.Print "End"
End Sub

