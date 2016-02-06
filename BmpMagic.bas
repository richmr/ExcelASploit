Attribute VB_Name = "BmpMagic"
Function hexToDec(hex As String) As Integer
    hexToDec = WorksheetFunction.Hex2Dec(hex)
End Function

Function intToDWORD(number As Long)
    ' Will convert an integer number (signed) into a 4 byte array, little endian
    Dim numInHex As String
    numInHex = WorksheetFunction.Dec2Hex(number, 8)
    ' Debug.Print numInHex
    
    Dim dword(3) As Byte
    Dim numInHexIndex As Long
    numInHexIndex = Len(numInHex) - 1
    For Index = 0 To 3
        Dim currentByte As String
        currentByte = Mid(numInHex, numInHexIndex, 2)
        dword(Index) = hexToDec(currentByte)
        numInHexIndex = numInHexIndex - 2
    Next
        
    intToDWORD = dword
End Function

Function DWORDToInt(ByRef dataIn() As Byte, offset As Long, wordLength As Long)
    ' Converts little-endian data into a proper integer
    ' Allows pulling the data from the middle of a big data array, so I don't have
    ' to keep copying it
    ' WARNING: modifies offset! (which can be really useful)
    Dim hexdata As String
    For i = 1 To wordLength
        hexdata = WorksheetFunction.Dec2Hex(dataIn(offset), 2) + hexdata
        offset = offset + 1
    Next
    Dim result As Long
    result = WorksheetFunction.Floor(Val(WorksheetFunction.Hex2Dec(hexdata)), 1)
    DWORDToInt = result
End Function

Sub copyBArrtoBArr(ByRef outBytes() As Byte, startIndex As Long, ByRef inBytes() As Byte)
    For Index = 0 To UBound(inBytes)
        outBytes(startIndex) = inBytes(Index)
        startIndex = startIndex + 1
    Next
End Sub

Sub insertDWORDToBArr(ByRef barr() As Byte, startIndex As Long, dwordVal As Long)
    Dim dword() As Byte
    dword = intToDWORD(dwordVal)
    copyBArrtoBArr barr, startIndex, dword
End Sub

Function bmpheader(pxDataSize As Long, payloadSize As Long)
    ' Produces and returns an appropriate BMP header as bytes
    Dim header(13) As Byte
    
    ' Initial ID Field "BM"
    header(0) = Asc("B")
    header(1) = Asc("M")
    
    ' Code up the size
    Dim bmpFileSize As Long
    ' file size is pix matrix size + 54 for BITMAPINFOHEADER
    bmpFileSize = pxDataSize + 54
    insertDWORDToBArr header, 2, bmpFileSize
    
    ' Code up the payload size (the actual data in the image)
    ' In a normal BMP these bytes are "unused" or application specific
    insertDWORDToBArr header, 6, payloadSize
    
    ' Code up the Offset.  Always 54 with this header type
    insertDWORDToBArr header, 10, 54
            
    bmpheader = header
End Function

Function dibheader(pxDataSize As Long, width As Long, height As Long)
    ' Creates standard 40 byte DIB header (24 bit/px)
    Dim header(39) As Byte
    
    ' Bytes in DIB header
    insertDWORDToBArr header, 0, 40
    
    ' Width, Height
    insertDWORDToBArr header, 4, width
    insertDWORDToBArr header, 8, height
        
    ' Color planes
    header(12) = 1
    header(13) = 0
    
    ' Bits/pix
    header(14) = 24
    header(15) = 0
    
    ' pixel compression (none)
    insertDWORDToBArr header, 16, 0
    
    ' Size of raw data
    insertDWORDToBArr header, 20, pxDataSize
    
    ' 72 DPI
    insertDWORDToBArr header, 24, 2835
    insertDWORDToBArr header, 28, 2835
    
    ' No colors in palette
    insertDWORDToBArr header, 32, 0
    
    ' All colors important
    insertDWORDToBArr header, 36, 0
    
    dibheader = header
End Function



Function makeBMP(ByRef pxData() As Byte, width As Long, height As Long, payloadSize As Long)
    ' Returns a properly formatted byte array that, when saved, will be a bmp
    ' Expects pxData to be a properly formatted array of pixels in 24-bit mode
    ' Also expects the pxData to be properly padded on each row as necessary to
    ' meet the BMP 4-byte requirement
    
    Dim pxDataSize As Long
    pxDataSize = UBound(pxData) + 1
    
    Dim finalDataSize As Long
    finalDataSize = 54 + pxDataSize
    
    Dim finalData() As Byte
    ReDim finalData(finalDataSize - 1)
    Dim bmphead() As Byte
    Dim dibhead() As Byte
    Dim offset As Long
    
    bmphead = bmpheader(pxDataSize, payloadSize)
    dibhead = dibheader(pxDataSize, width, height)
    
    offset = 0
    copyBArrtoBArr finalData, offset, bmphead
    copyBArrtoBArr finalData, offset, dibhead
    copyBArrtoBArr finalData, offset, pxData
    
    makeBMP = finalData
End Function

Function turnDataIntoPxData(ByRef origData() As Byte, ByRef width As Long, ByRef height As Long)
    ' Will turn data into a byte array of px data to meet an approximately 5x7 ratio image
    ' This is an arbitrary choice.
    ' Returns the array and modifies arguments width and length to meet actual
    ' Calc width/height
    Dim origDataSize As Long
    origDataSize = UBound(origData) + 1
    
    Dim ratio As Double
    ratio = 5 / 7   ' Change to meet your needs
    
    width = WorksheetFunction.Floor(Sqr(ratio * origDataSize * 3), 1)
    
    Dim widthMod As Long
    ' Rows must be a multiple of least common denominator of 3 and 4
    widthMod = 12 - width Mod 12
    If (widthMod < 12) Then
        width = width + widthMod
    End If
    height = WorksheetFunction.Ceiling(origDataSize / width, 1)
    
    Dim pxDataSize As Double
    pxDataSize = width * height
    
    Dim padding As Long
    padding = pxDataSize - origDataSize
    
    ' Dim pxDataArraySize As Long
    ' pxDataArraySize = pxDataSize - 1
    
    Dim pxDataArray() As Byte
    ReDim pxDataArray(pxDataSize - 1)
    copyBArrtoBArr pxDataArray, 0, origData
    
    ' width was in bytes, now needs to be in px (div 3 per px)
    width = width / 3
    
    turnDataIntoPxData = pxDataArray
End Function

Function getFileintoByteArr(filename As String)
    ' Returns a byte array with all the data from the file in it
    ' clearly a possible memory problem.  Use at own risk
    Open filename For Binary Access Read As #1
    fileLength = FileLen(filename)
    
    Dim fileData() As Byte
    ReDim fileData(fileLength - 1)
    
    Get #1, , fileData
    Close #1
    
    getFileintoByteArr = fileData
End Function

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Sub writeByteArrToFile(filename As String, ByRef dataIn() As Byte)
    ' Will skrag existing files.  Use at own risk
    DeleteFile filename
    
    Open filename For Binary Access Write As #1
    For i = 0 To UBound(dataIn)
        Put #1, , dataIn(i)
    Next
    Close #1
End Sub

Sub xorEncodeBytes(ByRef dataIn() As Byte, ByVal key As Byte)
    ' Conducts a basic cascading XOR encoding of data in dataIn
    ' This is not adequate encryption to protect privacy
    ' newByte(0) = oldByte(0) Xor key
    ' newByte(i) = oldByte(i) Xor newByte(i-1) Xor key
    
    dataIn(0) = dataIn(0) Xor key
    
    For i = 1 To UBound(dataIn)
        dataIn(i) = (dataIn(i) Xor dataIn(i - 1)) Xor key
    Next
        
End Sub

Sub xorDecodeBytes(ByRef dataIn() As Byte, ByVal key As Byte)
    ' decodedByte(i) = encodedByte(i) Xor encodedByte(i-1) Xor key
    ' decodedByte(0) = encodedByte(0) Xor key
    
    For i = UBound(dataIn) To 1 Step -1
        dataIn(i) = (dataIn(i) Xor dataIn(i - 1)) Xor key
    Next
    
    dataIn(0) = dataIn(0) Xor key
End Sub

Sub convertFileToBMP(filein As String, fileout As String, Optional ByVal key As Byte)
    ' Get data
    Dim payloadData() As Byte
    payloadData = getFileintoByteArr(filein)
    
    Dim payloadSize As Long
    payloadSize = UBound(payloadData) + 1
    
    ' Scramble it?
    If key Then
        xorEncodeBytes payloadData, key
    End If
        
    ' Turn into a picture
    Dim width As Long
    Dim height As Long
    Dim pxData() As Byte
    pxData = turnDataIntoPxData(payloadData, width, height)
    
    Debug.Print (UBound(pxData) + 1), payloadSize, width, height
    
    Dim bmpData() As Byte
    bmpData = makeBMP(pxData, width, height, payloadSize)
    
    Debug.Print UBound(bmpData) + 1
    
    ' Save it
    writeByteArrToFile fileout, bmpData
    
End Sub

Sub recoverFileFromBMP(filein As String, fileout As String, Optional ByVal key As Byte)
    ' This is actually much easier.  It just recovers the payload length from the known offset
    ' recovers the data starting at the known offset, decodes as necessary and saves
    
    Dim allData() As Byte
    allData = getFileintoByteArr(filein)
    
    ' Payload size starts at 0x06
    Dim payloadSize As Long
    payloadSize = DWORDToInt(allData, 6, 4)
    
    ' Grab the data
    Dim dataOffset As Long
    dataOffset = DWORDToInt(allData, 10, 4)
    
    Dim payload() As Byte
    ReDim payload(payloadSize - 1)
    
    For i = 1 To payloadSize
        payload(i - 1) = allData(dataOffset)
        dataOffset = dataOffset + 1
    Next
    
    ' unscramble it?
    If key Then
        xorDecodeBytes payload, key
    End If
    
    ' Save it
    writeByteArrToFile fileout, payload
    
End Sub

