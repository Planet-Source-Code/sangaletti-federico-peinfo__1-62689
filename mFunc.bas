Attribute VB_Name = "mFunc"
'########################################################
'####                 SOME FUNCTIONS                 ####
'########################################################


Public Function FirstImageSection(ntheader As IMAGE_NT_HEADERS, start_address As Long) As Long

'THIS FUNCTION RETURNS A POINTER TO THE FIRST SECTION HEADER

    FirstImageSection = start_address + ntheader.FileHeader.SizeOfOptionalHeader + &H18
End Function

Public Function GetAlignedSize(RealSize As Long, Alignment As Long) As Long
    
'THIS FUNCTION RETURNS THE ALIGNED SIZE
    
    While GetAlignedSize < RealSize
        GetAlignedSize = GetAlignedSize + Alignment
    Wend
End Function

Public Function RVAToOffset(Section_Headers() As IMAGE_SECTION_HEADER, RVA As Long) As Long
    
'THIS FUNCTION RETURNS THE REAL OFFSET FROM A REALTIVE VIRTUAL ADDRESS (RVA)
    
    Dim Offset As Long, Limit As Long, i As Integer
    
    Offset = RVA
    If RVA < Section_Headers(0).PointerToRawData Then
        RVAToOffset = RVA
        Exit Function
    End If
    
    For i = 0 To UBound(Section_Headers)
        If Section_Headers(i).SizeOfRawData Then
            Limit = Section_Headers(i).SizeOfRawData
        Else
            Limit = Section_Headers(i).VirtualSize
        End If
        
        If RVA >= Section_Headers(i).VirtualAddress And RVA < (Section_Headers(i).VirtualAddress + Limit) Then
            If Section_Headers(i).PointerToRawData <> 0 Then
                Offset = Offset - Section_Headers(i).VirtualAddress
                Offset = Offset + Section_Headers(i).PointerToRawData
            End If
            
            RVAToOffset = Offset
            Exit Function
        End If
    Next i
    
    RVAToOffset = 0
End Function

