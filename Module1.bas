Attribute VB_Name = "mPEinfO"
'##########################################################
'##########################################################
'######## Title:. PEInfO [26/09/05]                ########
'######## Author: Sangaletti Federico              ########
'######## e-mail: sangaletti@aliceposta.it         ########
'########------------------------------------------########
'########   !IF YOU LIKE THIS CODE PLEASE VOTE!    ########
'########------------------------------------------########
'##########################################################
'##########################################################

Public Const ICON_DOS_HEADER = 1
Public Const ICON_NT_HEADERS = 2
Public Const ICON_SECTION_HEADER = 3
Public Const ICON_EXPORT_TABLE = 4
Public Const ICON_IMPORT_TABLE = 4

Public DOS_HEADER_INFO As String
Public NT_HEADERS_INFO As String
Public SECTION_TABLE As String
Public EXPORT_TABLE As String
Public IMPORT_TABLE As String

Public Type ModuleName
    Name As String * 32
End Type

Public Type FunctionsAddress
    Address As Long
End Type

Public Type FunctionNames
    Names As Long
End Type

Public Type FunctionNameOrds
    NameOrds As Long
End Type

Public Type FunctionName
    Name As String * 32
End Type

Public Type Thunk
    Addr As Long
End Type

Public SectionHeaders() As IMAGE_SECTION_HEADER

Public Function GetPEInfo(FileName As String) As Long
    On Error GoTo hError
    
    Dim hFile As Long, FileSize As Long, bRW As Long, nSections As Integer, curOffset As Long, cur2Offset As Long
    Dim ExportTableOffset As Long, ImportTableOffset As Long
    Dim DOSheader As IMAGE_DOS_HEADER
    Dim NTHeaders As IMAGE_NT_HEADERS
    'Dim SectionHeaders() As IMAGE_SECTION_HEADER
    Dim ExportDirectory As IMAGE_EXPORT_DIRECTORY
    Dim ImportDescriptor As IMAGE_IMPORT_DESCRIPTOR
    
    Dim mName As ModuleName, fName As FunctionName
    Dim fAddress() As FunctionsAddress
    Dim fNames() As FunctionNames
    Dim fNameOrds() As FunctionNameOrds
    Dim Thunks As Thunk
    Dim ImgName As IMAGE_IMPORT_BY_NAME
    
    
    DOS_HEADER_INFO = ""
    NT_HEADERS_INFO = ""
    SECTION_TABLE = ""
    EXPORT_TABLE = ""
    IMPORT_TABLE = ""
    
    
    'OPEN THE FILE
    hFile = CreateFile(FileName, ByVal (GENERIC_READ Or GENERIC_WRITE), FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)
    If hFile = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to open '" & FileName & "'", vbCritical
        AddPESection = 0
        Exit Function
    End If
    
    
'########################################################
'####                GET DOS HEADER                  ####
'########################################################
    
    'READ THE DOS HEADER
    ReadFile hFile, DOSheader, Len(DOSheader), bRW, ByVal 0&
    
    'CHECK IF IS A VALID DOS HEADER
    If DOSheader.e_magic <> IMAGE_DOS_SIGNATURE Then
        MsgBox "Invalid DOS header!", vbCritical
        AddPESection = 0
        GoTo hError
    End If
    
    'ADD DOS HEADER ICON AND PRINT INFORMATIONS
    fPEinfO.lstReport.ListItems.Add 1, "kDOSHeader", "DOS Header", ICON_DOS_HEADER
    
    AddFileInfo " [ IMAGE_DOS_HEADER ]", DOS_HEADER_INFO
    AddFileInfo " --------------------", DOS_HEADER_INFO
    AddFileInfo " Signature: ........ 0x" & Hex$(DOSheader.e_magic) & " (MZ)", DOS_HEADER_INFO
    AddFileInfo " NT Headers address: 0x" & Hex$(DOSheader.e_lfanew), DOS_HEADER_INFO
    AddFileInfo "", DOS_HEADER_INFO
    AddFileInfo " Total size: " & Len(DOSheader) & " bytes", DOS_HEADER_INFO
    
    
'########################################################
'####                 GET NT HEADERS                 ####
'########################################################
    
    'e_lfanew IS THE POINTER TO NT HEADERS
    SetFilePointer hFile, DOSheader.e_lfanew, 0, 0
    
    'READ THE NT HEADERS
    ReadFile hFile, NTHeaders, Len(NTHeaders), bRW, ByVal 0&
    
    'CHECK IF IS A VALID PE HEADER
    If NTHeaders.Signature <> IMAGE_NT_SIGNATURE Then
        MsgBox "Invalid PE file!", vbCritical
        AddPESection = 0
        GoTo hError
    End If
    
    'ADD NTHEADERS ICON AND PRINT INFORMATIONS
    fPEinfO.lstReport.ListItems.Add 2, "kNTHeaders", "NT Headers", ICON_NT_HEADERS
    
    AddFileInfo " [ IMAGE_NT_HEADERS ]", NT_HEADERS_INFO
    AddFileInfo " --------------------", NT_HEADERS_INFO
    AddFileInfo " Signature: ............. 0x" & Hex$(NTHeaders.Signature) & " (PE__)", NT_HEADERS_INFO
    AddFileInfo "", NT_HEADERS_INFO
    
    'PRINT FILE HEADER INFORMATION
    With NTHeaders.FileHeader
        AddFileInfo " [ IMAGE_FILE_HEADER ]", NT_HEADERS_INFO
        AddFileInfo " ---------------------", NT_HEADERS_INFO
        AddFileInfo " Machine (CPU): ......... 0x" & .Machine & " " & GetCPUName(.Machine), NT_HEADERS_INFO
        AddFileInfo " Number of sections: .... " & .NumberOfSections, NT_HEADERS_INFO
        AddFileInfo " Size of Optional Header: 0x" & Hex$(.SizeOfOptionalHeader), NT_HEADERS_INFO
        AddFileInfo " Characteristics: ....... 0x" & Hex$(.Characteristics), NT_HEADERS_INFO
        AddFileInfo "", NT_HEADERS_INFO
        AddFileInfo "", NT_HEADERS_INFO
    End With
    
    'PRINT OPTIONAL HEADER INFORMATION
    With NTHeaders.OptionalHeader
        AddFileInfo " [ IMAGE_OPTIONAL_HEADER ]", NT_HEADERS_INFO
        AddFileInfo " -------------------------", NT_HEADERS_INFO
        AddFileInfo " Magic word: ............ 0x" & Hex$(.Magic) & " " & GetMagicName(.Magic), NT_HEADERS_INFO
        AddFileInfo " Size of executable code: " & .SizeOfCode & " bytes", NT_HEADERS_INFO
        AddFileInfo " Address of Entrypoint: . 0x" & Hex$(.AddressOfEntryPoint), NT_HEADERS_INFO
        AddFileInfo " Base of code: .......... 0x" & Hex$(.BaseOfCode), NT_HEADERS_INFO
        AddFileInfo " Base of data: .......... 0x" & Hex$(.BaseOfData), NT_HEADERS_INFO
        AddFileInfo " ImageBase: ............. 0x" & Hex$(.ImageBase), NT_HEADERS_INFO
        AddFileInfo " Section alignment: ..... 0x" & Hex$(.SectionAlignment), NT_HEADERS_INFO
        AddFileInfo " File alignment: ........ 0x" & Hex$(.FileAlignment), NT_HEADERS_INFO
        AddFileInfo " Size of image in memory: 0x" & Hex$(.SizeOfImage), NT_HEADERS_INFO
        AddFileInfo " Size of headers: ....... 0x" & Hex$(.SizeOfHeaders), NT_HEADERS_INFO
        AddFileInfo " Subsystem: ............. 0x" & Hex$(.Subsystem) & " " & GetSubsystemName(.Subsystem), NT_HEADERS_INFO
        AddFileInfo " Checksum: .............. 0x" & Hex$(.CheckSum), NT_HEADERS_INFO
    End With
    
    AddFileInfo "", NT_HEADERS_INFO
    AddFileInfo " Size of File Header: ... " & Len(NTHeaders.FileHeader) & " bytes", NT_HEADERS_INFO
    AddFileInfo " Size of Optional Header: " & Len(NTHeaders.OptionalHeader) & " bytes", NT_HEADERS_INFO
    
    
'########################################################
'####              GET SECTION TABLE                 ####
'########################################################

    'GET THE NUMBER OF SECTIONS
    nSections = NTHeaders.FileHeader.NumberOfSections - 1
    ReDim SectionHeaders(nSections)
    
    'POINT TO FIRST SECTION HEADER
    SetFilePointer hFile, FirstImageSection(NTHeaders, DOSheader.e_lfanew), 0, 0

    'ADD SECTION HEADER ICON AND PRINT INFORMATIONS
    fPEinfO.lstReport.ListItems.Add 3, "kSectionHeader", "Section Table", ICON_SECTION_HEADER

    AddFileInfo " [ IMAGE_SECTION_HEADER ]", SECTION_TABLE
    AddFileInfo " ------------------------", SECTION_TABLE

    'FILL SectionHeaders ARRAY AND PRINT INFOS OF ALL SECTIONS
    For i = 0 To nSections
        ReadFile hFile, SectionHeaders(i), IMAGE_SIZEOF_SECTION_HEADER, bRW, ByVal 0&
        AddFileInfo "", SECTION_TABLE
        With SectionHeaders(i)
            AddFileInfo " Section name: ..... " & FullName(.nameSec), SECTION_TABLE
            AddFileInfo " Size: ............. 0x" & Hex$(.VirtualSize), SECTION_TABLE
            AddFileInfo " Aligned size: ..... 0x" & Hex$(.SizeOfRawData), SECTION_TABLE
            AddFileInfo " Virtual address: .. 0x" & Hex$(.VirtualAddress), SECTION_TABLE
            AddFileInfo " Phisical address: . 0x" & Hex$(.PhisicalAddress), SECTION_TABLE
            AddFileInfo " Characteristics: .. 0x" & Hex$(.Characteristics) & " " & GetCharacteristicsName(.Characteristics), SECTION_TABLE
        End With
    Next i
    
    AddFileInfo "", SECTION_TABLE
    AddFileInfo " Size of Section Header: " & IMAGE_SIZEOF_SECTION_HEADER & " bytes", SECTION_TABLE
    AddFileInfo " Total size of headers:  " & IMAGE_SIZEOF_SECTION_HEADER * (nSections + 1) & " bytes", SECTION_TABLE


'########################################################
'####               GET EXPORT TABLE                 ####
'########################################################
    
    'ADD EXPORT TABLE ICON
    fPEinfO.lstReport.ListItems.Add 4, "kExportTable", "Export Table", ICON_EXPORT_TABLE
    
    AddFileInfo " [ IMAGE_EXPORT_DIRECTORY ]", EXPORT_TABLE
    AddFileInfo " --------------------------", EXPORT_TABLE
    
    'CHECK IF EXISTS EXPORTS
    If NTHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY.IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress <> 0 Then
        ExportTableOffset = RVAToOffset(SectionHeaders, NTHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY.IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress)
        
        If ExportTableOffset = 0 Then
            AddFileInfo "", EXPORT_TABLE
            AddFileInfo " This PE doesn't contain a valid Export Table.", EXPORT_TABLE
        Else
            'READ EXPORT TABLE
            SetFilePointer hFile, ExportTableOffset, 0, 0
            ReadFile hFile, ExportDirectory, Len(ExportDirectory), bRW, ByVal 0&
            
            SetFilePointer hFile, RVAToOffset(SectionHeaders, ExportDirectory.Name), 0, 0
            ReadFile hFile, mName, Len(mName), bRW, ByVal 0&
            
            With ExportDirectory
                AddFileInfo " Name: ............... " & FormatString(mName.Name), EXPORT_TABLE
                AddFileInfo " Base: ............... " & .Base, EXPORT_TABLE
                AddFileInfo " Functions exported: . " & .NumberOfFunctions, EXPORT_TABLE
                AddFileInfo " Numbers of names: ... " & .NumberOfNames, EXPORT_TABLE
                AddFileInfo " Addr of functions: .. 0x" & Hex$(.AddressOfFunctions), EXPORT_TABLE
                AddFileInfo " Addr of names: ...... 0x" & Hex$(.AddressOfNames), EXPORT_TABLE
                AddFileInfo " Addr of name ords: .. 0x" & Hex$(.AddressOfNameOrdinals), EXPORT_TABLE
            End With
            
            'FILL ALL EXPORTS ARRAY
            SetFilePointer hFile, RVAToOffset(SectionHeaders, ExportDirectory.AddressOfFunctions), 0, 0
            ReDim fAddress(ExportDirectory.NumberOfFunctions - 1)
            For i = 0 To UBound(fAddress)
                ReadFile hFile, fAddress(i), Len(fAddress(i)), bRW, ByVal 0&
            Next i
            
            SetFilePointer hFile, RVAToOffset(SectionHeaders, ExportDirectory.AddressOfNames), 0, 0
            ReDim fNames(ExportDirectory.NumberOfNames - 1)
            For i = 0 To UBound(fNames)
                ReadFile hFile, fNames(i), Len(fNames(i)), bRW, ByVal 0&
            Next i
            
            SetFilePointer hFile, RVAToOffset(SectionHeaders, ExportDirectory.AddressOfNameOrdinals), 0, 0
            ReDim fNameOrds(ExportDirectory.NumberOfNames - 1)
            For i = 0 To UBound(fNames)
                ReadFile hFile, fNameOrds(i), Len(fNameOrds(i)), bRW, ByVal 0&
            Next i
            
            'READ ALL EXPORTS
            For x = 0 To UBound(fAddress)
                If fAddress(x).Address <> 0 Then
                
                    AddFileInfo "", EXPORT_TABLE
                    AddFileInfo " Ordinal: ........ " & (x + ExportDirectory.Base), EXPORT_TABLE
                    AddFileInfo " Function address: 0x" & Hex$(fAddress(x).Address), EXPORT_TABLE
                    
                    'IF EXPORTED BY NAME PRINT THEN NAME ELSE ONLY THE ORDINAL
                    For y = 0 To UBound(fNames)
                        If y = x Then
                            SetFilePointer hFile, RVAToOffset(SectionHeaders, fNames(y).Names), 0, 0
                            ReadFile hFile, fName, Len(fName), bRW, ByVal 0&
                            AddFileInfo " Function name: .. " & FormatString(fName.Name), EXPORT_TABLE
                        End If
                    Next y
                End If
            Next x
            
        End If
    Else
        AddFileInfo "", EXPORT_TABLE
        AddFileInfo " This PE doesn't contain an Export Table.", EXPORT_TABLE
    End If
    
    
'########################################################
'####               GET IMPORT TABLE                 ####
'########################################################
    
    'ADD IMPORT TABLE ICON
    fPEinfO.lstReport.ListItems.Add 5, "kImportTable", "Import Table", ICON_IMPORT_TABLE
    
    AddFileInfo " [ IMAGE_IMPORT_DESCRIPTOR ]", IMPORT_TABLE
    AddFileInfo " ---------------------------", IMPORT_TABLE
    
    'CHECK IF EXISTS A VALID IMPORT TABLE
    If NTHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY.IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress <> 0 Then
        ImportTableOffset = RVAToOffset(SectionHeaders, NTHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY.IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress)
        
        If ImportTableOffset = 0 Then
            AddFileInfo "", IMPORT_TABLE
            AddFileInfo " This PE doesn't contain an valid Import Table.", IMPORT_TABLE
        Else
            
            'READ IMPORT DESCRIPTOR
            SetFilePointer hFile, ImportTableOffset, 0, 0
            ReadFile hFile, ImportDescriptor, Len(ImportDescriptor), bRW, ByVal 0&
            
            'READ ALL IMPORTS
            While ImportDescriptor.FirstThunk <> 0
                curOffset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
                SetFilePointer hFile, RVAToOffset(SectionHeaders, ImportDescriptor.Name), 0, 0
                ReadFile hFile, mName, Len(mName), bRW, ByVal 0&
                
                AddFileInfo "", IMPORT_TABLE
                AddFileInfo " [ Module " & StrConv(FormatString(mName.Name), vbUpperCase) & " ]", IMPORT_TABLE
                
                
                
                SetFilePointer hFile, RVAToOffset(SectionHeaders, IIf(ImportDescriptor.OriginalFirstThunk <> 0, ImportDescriptor.OriginalFirstThunk, ImportDescriptor.FirstThunk)), 0, 0
                ReadFile hFile, Thunks, Len(Thunks), bRW, ByVal 0&
                
                'READ ALL THUNKS
                While Thunks.Addr <> 0
                    
                    'IF IMPORTED BY NAME PRINT ELSE PRINT ONLY THE ORDINAL
                    If Thunks.Addr And IMAGE_ORDINAL_FLAG Then
                        AddFileInfo " Function ordinal: 0x" & Hex$(Thunks.Addr - IMAGE_ORDINAL_FLAG), IMPORT_TABLE
                    Else
                        cur2Offset = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
                        SetFilePointer hFile, RVAToOffset(SectionHeaders, Thunks.Addr), 0, 0
                        ReadFile hFile, ImgName, Len(ImgName), bRW, ByVal 0&
                        AddFileInfo " Function name: " & FormatString(ImgName.Name), IMPORT_TABLE
                    End If
                    SetFilePointer hFile, cur2Offset, 0, 0
                    ReadFile hFile, Thunks, Len(Thunks), bRW, ByVal 0&
                Wend
                
                SetFilePointer hFile, curOffset, 0, 0
                ReadFile hFile, ImportDescriptor, Len(ImportDescriptor), bRW, ByVal 0&
            Wend
        End If
    Else
        AddFileInfo "", IMPORT_TABLE
        AddFileInfo " This PE doesn't contain an Import Table.", IMPORT_TABLE
    End If
    
    CloseHandle hFile
    Exit Function
    
hError:
    If hFile <> 0 Then CloseHandle hFile
    MsgBox "Error " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbCritical
End Function


Public Sub AddFileInfo(Info As String, Buffer As String)
    Buffer = Buffer & Info & vbCrLf
End Sub

Public Function FullName(Name As String) As String
    Dim i As Integer
    
    For i = 1 To 6
        If Mid(Name, i, 1) = Chr(&H0) Then Exit For
    Next i
    FullName = FullName & Left(Name, i - 1)
End Function

Public Function GetCPUName(Code As Integer) As String
    Select Case Code
        Case 0: GetCPUName = "MACHINE_UNKNOWN"
        Case &H14C: GetCPUName = "MACHINE_I386"
        Case &H162: GetCPUName = "MACHINE_R3000"
        Case &H166: GetCPUName = "MACHINE_R4000"
        Case &H168: GetCPUName = "MACHINE_R10000"
        Case &H169: GetCPUName = "MACHINE_WCEMIPSV2"
        Case &H184: GetCPUName = "MACHINE_ALPHA"
        Case &H1F0: GetCPUName = "MACHINE_POWERPC"
        Case &H1A2: GetCPUName = "MACHINE_SH3"
        Case &H1A4: GetCPUName = "MACHINE_SH3E"
        Case &H1A6: GetCPUName = "MACHINE_SH4"
        Case &H1C0: GetCPUName = "MACHINE_ARM"
        Case &H1C2: GetCPUName = "MACHINE_THUMB"
        Case &H200: GetCPUName = "MACHINE_IA64"
        Case &H266: GetCPUName = "MACHINE_MIPS16"
        Case &H366: GetCPUName = "MACHINE_MIPSFPU"
        Case &H466: GetCPUName = "MACHINE_MIPSFPU16"
        Case &H284: GetCPUName = "MACHINE_ALPHA6"
    End Select
End Function

Public Function GetMagicName(Code As Integer) As String
    Select Case Code
        Case &H10B: GetMagicName = "NT_OPTIONAL_HDR32_MAGIC"
        Case &H20B: GetMagicName = "NT_OPTIONAL_HDR64_MAGIC"
        Case &H107: GetMagicName = "ROM_OPTIONAL_HDR_MAGIC"
    End Select
End Function

Public Function GetSubsystemName(Code As Integer) As String
    Select Case Code
        Case 0: GetSubsystemName = "SUBSYSTEM_UNKNOWN"
        Case 1: GetSubsystemName = "SUBSYSTEM_NATIVE"
        Case 2: GetSubsystemName = "SUBSYSTEM_WINDOWS_GUI"
        Case 3: GetSubsystemName = "SUBSYSTEM_WINDOWS_CONSOLE"
        Case 5: GetSubsystemName = "SUBSYSTEM_OS2_CONSOLE"
        Case 7: GetSubsystemName = "SUBSYSTEM_POSIX_CONSOLE"
        Case 8: GetSubsystemName = "SUBSYSTEM_NATIVE_DRIVER_WINDOWS"
        Case 9: GetSubsystemName = "WINDOWS_CE_GUI"
    End Select
End Function

Public Function GetCharacteristicsName(Code As Long) As String
    
    If Code And &H20 Then GetCharacteristicsName = "SECTION_CODE"
    If Code And &H40 Then GetCharacteristicsName = "SECTION_INITIALIZED_DATA"
    If Code And &H80 Then GetCharacteristicsName = "SECTION_UNINITIALIZED_DATA"
    If Code And &H10000000 Then GetCharacteristicsName = GetCharacteristicsName & " + SHAREABLE"
    If Code And &H20000000 Then GetCharacteristicsName = GetCharacteristicsName & " + EXECUTABLE"
    If Code And &H40000000 Then GetCharacteristicsName = GetCharacteristicsName & " + READABLE"
    If Code And &H80000000 Then GetCharacteristicsName = GetCharacteristicsName & " + WRITABLE"

End Function


Public Function FormatString(str As String) As String
    Dim i As Integer
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) = Chr(&H0) Then Exit For
    Next i
    
    FormatString = Left(str, i - 1)
End Function
