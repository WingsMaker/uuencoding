' ==============================================================================
' UUencode / UUdecode VBScript Implementation
' Usage:
'   cscript uu.vbs encode   (defaults: input="zz.zip", output="zz.txt")
'   cscript uu.vbs decode   (defaults: input="zz.txt", output="zz.dat")
' ==============================================================================

Option Explicit

Dim mode, inFile, outFile
Dim args

Set args = WScript.Arguments

If args.Count >= 1 Then
    mode = LCase(args(0))
Else
    ' Default to encode if no argument is passed
    mode = "encode"
End If

Select Case mode
    Case "encode"
        inFile = "zz.zip"
        outFile = "zz.txt"
        If args.Count >= 2 Then inFile = args(1)
        If args.Count >= 3 Then outFile = args(2)
        WScript.Echo "Encoding """ & inFile & """ -> """ & outFile & """..."
        UUEncodeFile inFile, outFile
        WScript.Echo "Encoding complete."

    Case "decode"
        inFile = "zz.txt"
        outFile = "zz.dat"
        If args.Count >= 2 Then inFile = args(1)
        If args.Count >= 3 Then outFile = args(2)
        WScript.Echo "Decoding """ & inFile & """ -> """ & outFile & """..."
        UUDecodeFile inFile, outFile
        WScript.Echo "Decoding complete."

    Case Else
        WScript.Echo "Usage: cscript uu.vbs [encode|decode] [input_file] [output_file]"
End Select


' ==============================================================================
' ENCODE FUNCTION
' ==============================================================================
Sub UUEncodeFile(srcPath, dstPath)
    Dim fso, streamIn, dataBytes, totalLen, i, j, chunkLen
    Dim b1, b2, b3, c1, c2, c3, c4
    Dim lineStr, textStream
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcPath) Then
        WScript.Echo "Error: Input file not found: " & srcPath
        WScript.Quit 1
    End If
    
    ' Read binary file using ADODB.Stream
    Set streamIn = CreateObject("ADODB.Stream")
    streamIn.Type = 1 ' adTypeBinary
    streamIn.Open
    streamIn.LoadFromFile srcPath
    dataBytes = streamIn.Read
    streamIn.Close
    
    totalLen = LenB(dataBytes)
    
    ' Write text file using FileSystemObject text stream
    Set textStream = fso.CreateTextFile(dstPath, True, False)
    
    ' Write standard UUencode header
    textStream.WriteLine "begin 644 " & fso.GetFileName(dstPath)
    
    i = 1
    Do While i <= totalLen
        ' Determine chunk length (up to 45 bytes per line)
        If i + 44 <= totalLen Then
            chunkLen = 45
        Else
            chunkLen = totalLen - i + 1
        End If
        
        ' Length character (encoded)
        lineStr = Chr((chunkLen Mod 64) + 32)
        If Asc(lineStr) = 32 Then lineStr = "`"
        
        ' Process bytes in groups of 3
        For j = 0 To chunkLen - 1 Step 3
            ' Read up to 3 bytes safely with padding zeros if needed
            b1 = AscB(MidB(dataBytes, i + j, 1))
            If j + 1 < chunkLen Then b2 = AscB(MidB(dataBytes, i + j + 1, 1)) Else b2 = 0
            If j + 2 < chunkLen Then b3 = AscB(MidB(dataBytes, i + j + 2, 1)) Else b3 = 0
            
            ' Split 24 bits into four 6-bit values
            c1 = (b1 \ 4) And &H3F
            c2 = ((b1 And &H3) * 16) Or (b2 \ 16)
            c3 = ((b2 And &HF) * 4) Or (b3 \ 64)
            c4 = b3 And &H3F
            
            lineStr = lineStr & EncodeChar(c1) & EncodeChar(c2) & EncodeChar(c3) & EncodeChar(c4)
        Next
        
        textStream.WriteLine lineStr
        i = i + chunkLen
    Loop
    
    ' Write UUencode footer (single backtick and "end")
    textStream.WriteLine "`"
    textStream.WriteLine "end"
    textStream.Close
End Sub

Function EncodeChar(val)
    If val = 0 Then
        EncodeChar = "`"
    Else
        EncodeChar = Chr(val + 32)
    End If
End Function


' ==============================================================================
' DECODE FUNCTION
' ==============================================================================
Sub UUDecodeFile(srcPath, dstPath)
    Dim fso, textStream, line, streamOut
    Dim chunkLen, i, j, c1, c2, c3, c4, b1, b2, b3
    Dim byteList
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcPath) Then
        WScript.Echo "Error: Input file not found: " & srcPath
        WScript.Quit 1
    End If
    
    Set textStream = fso.OpenTextFile(srcPath, 1, False)
    Set byteList = CreateObject("System.Collections.ArrayList")
    
    Do While Not textStream.AtEndOfStream
        line = textStream.ReadLine
        
        ' Skip header/footer markers
        If Left(line, 5) = "begin" Or line = "end" Or line = "`" Or Trim(line) = "" Then
            ' Skip
        Else
            ' First character represents byte length of the line
            If Len(line) > 0 Then
                chunkLen = DecodeChar(Mid(line, 1, 1))
                
                ' Process 4-character chunks representing 3 bytes
                i = 2
                Do While i + 3 <= Len(line) And chunkLen > 0
                    c1 = DecodeChar(Mid(line, i, 1))
                    c2 = DecodeChar(Mid(line, i+1, 1))
                    c3 = DecodeChar(Mid(line, i+2, 1))
                    c4 = DecodeChar(Mid(line, i+3, 1))
                    
                    b1 = ((c1 * 4) Or (c2 \ 16)) And &HFF
                    b2 = ((c2 * 16) Or (c3 \ 4)) And &HFF
                    b3 = ((c3 * 64) Or c4) And &HFF
                    
                    If chunkLen >= 1 Then 
                        byteList.Add CByte(b1)
                        chunkLen = chunkLen - 1
                    End If
                    If chunkLen >= 1 Then 
                        byteList.Add CByte(b2)
                        chunkLen = chunkLen - 1
                    End If
                    If chunkLen >= 1 Then 
                        byteList.Add CByte(b3)
                        chunkLen = chunkLen - 1
                    End If
                    
                    i = i + 4
                Loop
            End If
        End If
    Loop
    textStream.Close
    
    ' Convert ArrayList to binary stream and save to output file
    Dim streamArray()
    ReDim streamArray(byteList.Count - 1)
    For i = 0 To byteList.Count - 1
        streamArray(i) = byteList(i)
    Next
    
    Dim rawData
    If byteList.Count > 0 Then
        rawData = BinaryDataFromArray(streamArray)
    Else
        rawData = ChrB(0)
    End If
    
    Set streamOut = CreateObject("ADODB.Stream")
    streamOut.Type = 1 ' adTypeBinary
    streamOut.Open
    streamOut.Write rawData
    streamOut.SaveToFile dstPath, 2 ' adSaveCreateOverWrite
    streamOut.Close
End Sub

Function DecodeChar(charStr)
    Dim val
    If charStr = "`" Then
        DecodeChar = 0
    Else
        val = Asc(charStr) - 32
        If val < 0 Then val = val + 64
        DecodeChar = val And &H3F
    End If
End Function

Function BinaryDataFromArray(arr)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    Dim rs
    Set rs = CreateObject("ADODB.Recordset")
    rs.Fields.Append "bin", 205, UBound(arr) + 1 ' adLongVarBinary
    rs.Open
    rs.AddNew
    rs("bin").AppendChunk arr
    rs.Update
    stream.Write rs("bin")
    rs.Close
    Set BinaryDataFromArray = stream.Read
    stream.Close
End Function
