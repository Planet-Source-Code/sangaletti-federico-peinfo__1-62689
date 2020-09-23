Attribute VB_Name = "mText"
'########################################################
'####              TEXTS FOR HEADER INFO             ####
'########################################################


Public Const TXT_IMAGE_DOS_HEADER = _
 "typedef struct _IMAGE_DOS_HEADER {      // DOS .EXE header" & vbCrLf & _
 "    WORD   e_magic;                     // Magic number" & vbCrLf & _
 "    WORD   e_cblp;                      // Bytes on last page of file" & vbCrLf & _
 "    WORD   e_cp;                        // Pages in file" & vbCrLf & _
 "    WORD   e_crlc;                      // Relocations" & vbCrLf & _
 "    WORD   e_cparhdr;                   // Size of header in paragraphs" & vbCrLf & _
 "    WORD   e_minalloc;                  // Minimum extra paragraphs needed" & vbCrLf & _
 "    WORD   e_maxalloc;                  // Maximum extra paragraphs needed" & vbCrLf & _
 "    WORD   e_ss;                        // Initial (relative) SS value" & vbCrLf & _
 "    WORD   e_sp;                        // Initial SP value" & vbCrLf & _
 "    WORD   e_csum;                      // Checksum" & vbCrLf & _
 "    WORD   e_ip;                        // Initial IP value" & vbCrLf & _
 "    WORD   e_cs;                        // Initial (relative) CS value" & vbCrLf & _
 "    WORD   e_lfarlc;                    // File address of relocation table" & vbCrLf & _
 "    WORD   e_ovno;                      // Overlay number" & vbCrLf & _
 "    WORD   e_res[4];                    // Reserved words" & vbCrLf & _
 "    WORD   e_oemid;                     // OEM identifier (for e_oeminfo)" & vbCrLf & _
 "    WORD   e_oeminfo;                   // OEM information; e_oemid specific" & vbCrLf & _
 "    WORD   e_res2[10];                  // Reserved words" & vbCrLf & _
 "    LONG   e_lfanew;                    // File address of new exe header" & vbCrLf & _
 "  } IMAGE_DOS_HEADER, *PIMAGE_DOS_HEADER;"

Public Const TXT_IMAGE_NT_HEADERS = _
"typedef struct _IMAGE_NT_HEADERS {" & vbCrLf & _
"    DWORD Signature;" & vbCrLf & _
"    IMAGE_FILE_HEADER FileHeader;" & vbCrLf & _
"    IMAGE_OPTIONAL_HEADER OptionalHeader;" & vbCrLf & _
"} IMAGE_NT_HEADERS, *PIMAGE_NT_HEADERS;"

Public Const TXT_IMAGE_FILE_HEADER = _
"typedef struct _IMAGE_FILE_HEADER {" & vbCrLf & _
"    WORD    Machine;" & vbCrLf & _
"    WORD    NumberOfSections;" & vbCrLf & _
"    DWORD   TimeDateStamp;" & vbCrLf & _
"    DWORD   PointerToSymbolTable;" & vbCrLf & _
"    DWORD   NumberOfSymbols;" & vbCrLf & _
"    WORD    SizeOfOptionalHeader;" & vbCrLf & _
"    WORD    Characteristics;" & vbCrLf & _
"} IMAGE_FILE_HEADER, *PIMAGE_FILE_HEADER;"


Public Const TXT_IMAGE_OPTIONAL_HEADER_1 = _
"typedef struct _IMAGE_OPTIONAL_HEADER {" & vbCrLf & _
"    WORD    Magic;" & vbCrLf & _
"    BYTE    MajorLinkerVersion;" & vbCrLf & _
"    BYTE    MinorLinkerVersion;" & vbCrLf & _
"    DWORD   SizeOfCode;" & vbCrLf & _
"    DWORD   SizeOfInitializedData;" & vbCrLf & _
"    DWORD   SizeOfUninitializedData;" & vbCrLf & _
"    DWORD   AddressOfEntryPoint;" & vbCrLf & _
"    DWORD   BaseOfCode;" & vbCrLf & _
"    DWORD   BaseOfData;" & vbCrLf & _
"    DWORD   ImageBase;" & vbCrLf & _
"    DWORD   SectionAlignment;" & vbCrLf & _
"    DWORD   FileAlignment;" & vbCrLf & _
"    WORD    MajorOperatingSystemVersion;" & vbCrLf & _
"    WORD    MinorOperatingSystemVersion;" & vbCrLf & _
"    WORD    MajorImageVersion;" & vbCrLf & _
"    WORD    MinorImageVersion;" & vbCrLf & _
"    WORD    MajorSubsystemVersion;" & vbCrLf & _
"    WORD    MinorSubsystemVersion;" & vbCrLf & _
"    DWORD   Win32VersionValue;" & vbCrLf & _
"    DWORD   SizeOfImage;" & vbCrLf & _
"    DWORD   SizeOfHeaders;" & vbCrLf & _
"    DWORD   CheckSum;" & vbCrLf & _
"    WORD    Subsystem;" & vbCrLf

Public Const TXT_IMAGE_OPTIONAL_HEADER = TXT_IMAGE_OPTIONAL_HEADER_1 & _
"    WORD    DllCharacteristics;" & vbCrLf & _
"    DWORD   SizeOfStackReserve;" & vbCrLf & _
"    DWORD   SizeOfStackCommit;" & vbCrLf & _
"    DWORD   SizeOfHeapReserve;" & vbCrLf & _
"    DWORD   SizeOfHeapCommit;" & vbCrLf & _
"    DWORD   LoaderFlags;" & vbCrLf & _
"    DWORD   NumberOfRvaAndSizes;" & vbCrLf & _
"    IMAGE_DATA_DIRECTORY DataDirectory[IMAGE_NUMBEROF_DIRECTORY_ENTRIES];" & vbCrLf & _
"} IMAGE_OPTIONAL_HEADER, *PIMAGE_OPTIONAL_HEADER;"

Public Const TXT_IMAGE_SECTION_HEADER = _
"typedef struct _IMAGE_SECTION_HEADER {" & vbCrLf & _
"    BYTE    Name[IMAGE_SIZEOF_SHORT_NAME];" & vbCrLf & _
"    union {" & vbCrLf & _
"            DWORD   PhysicalAddress;" & vbCrLf & _
"            DWORD   VirtualSize;" & vbCrLf & _
"    } Misc;" & vbCrLf & _
"    DWORD   VirtualAddress;" & vbCrLf & _
"    DWORD   SizeOfRawData;" & vbCrLf & _
"    DWORD   PointerToRawData;" & vbCrLf & _
"    DWORD   PointerToRelocations;" & vbCrLf & _
"    DWORD   PointerToLinenumbers;" & vbCrLf & _
"    WORD    NumberOfRelocations;" & vbCrLf & _
"    WORD    NumberOfLinenumbers;" & vbCrLf & _
"    DWORD   Characteristics;" & vbCrLf & _
"} IMAGE_SECTION_HEADER, *PIMAGE_SECTION_HEADER;"

Public Const TXT_IMAGE_EXPORT_DIRECTORY = _
"typedef struct _IMAGE_EXPORT_DIRECTORY {" & vbCrLf & _
"    DWORD   Characteristics;" & vbCrLf & _
"    DWORD   TimeDateStamp;" & vbCrLf & _
"    WORD    MajorVersion;" & vbCrLf & _
"    WORD    MinorVersion;" & vbCrLf & _
"    DWORD   Name;" & vbCrLf & _
"    DWORD   Base;" & vbCrLf & _
"    DWORD   NumberOfFunctions;" & vbCrLf & _
"    DWORD   NumberOfNames;" & vbCrLf & _
"    DWORD   AddressOfFunctions;     // RVA from base of image" & vbCrLf & _
"    DWORD   AddressOfNames;         // RVA from base of image" & vbCrLf & _
"    DWORD   AddressOfNameOrdinals;  // RVA from base of image" & vbCrLf & _
"} IMAGE_EXPORT_DIRECTORY, *PIMAGE_EXPORT_DIRECTORY;"

Public Const TXT_IMAGE_IMPORT_DESCRIPTOR = _
"typedef struct _IMAGE_IMPORT_DESCRIPTOR {" & vbCrLf & _
"    union {" & vbCrLf & _
"        DWORD   Characteristics;            // 0 for terminating null import descriptor" & vbCrLf & _
"        DWORD   OriginalFirstThunk;         // RVA to original unbound IAT (PIMAGE_THUNK_DATA)" & vbCrLf & _
"    };" & vbCrLf & _
"    DWORD   TimeDateStamp;                  // 0 if not bound," & vbCrLf & _
"                                            // -1 if bound, and real date\time stamp" & vbCrLf & _
"                                            //     in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)" & vbCrLf & _
"                                            // O.W. date/time stamp of DLL bound to (Old BIND)" & vbCrLf & _
"    DWORD   ForwarderChain;                 // -1 if no forwarders" & vbCrLf & _
"    DWORD   Name;" & vbCrLf & _
"    DWORD   FirstThunk;                     // RVA to IAT (if bound this IAT has actual addresses)" & vbCrLf & _
"} IMAGE_IMPORT_DESCRIPTOR, *PIMAGE_IMPORT_DESCRIPTOR;"


