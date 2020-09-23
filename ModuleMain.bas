Attribute VB_Name = "ModuleMain"
Option Explicit

' Overall .ZIP file format:

 '   [local file header 1]
 '   [file data 1]
 '   [data descriptor 1]
 '   .
 '   .
 '   .
 '   [local file header n]
 '   [file data n]
 '   [data descriptor n]
 '   [central directory]
 '   [zip64 end of central directory record]
 '   [zip64 end of central directory locator]
 '   [end of central directory record]

Public Type LocalFileHeader
  LocalFileHeaderSignature As Long ' (0x04034b50)
  VersionNeededToExtract As Integer
  GeneralPurposeBitFlag As Integer
  CompressionMethod As Integer
  LastModFileTime As Integer
  LastModFileDate As Integer
  CRC32 As Long
  CompressedSize As Long
  UncompressedSize As Long
  FileNameLength As Integer
  ExtraFieldLength As Integer
End Type
  
  ' file name (variable size)
  ' extra field (variable size)

Public Type DataDescriptor
  CRC32 As Long
  CompressedSize As Long
  UncompressedSize As Long
End Type

' Central directory structure:

 '     [file header 1]
 '     .
 '     .
 '     .
 '     [file header n]
 '     [digital signature]

Public Type FileHeader ' 47 bytes
  CentralFileHeaderSignature As Long ' (0x02014b50)
  VersionMadeBy As Integer
  VersionNeededToExtract As Integer
  GeneralPurposeBitFlag As Integer
  CompressionMethod As Integer
  LastModFileTime As Integer
  LastModFileDate As Integer
  CRC32 As Long
  CompressedSize As Long
  UncompressedSize As Long
  FileNameLength As Integer
  ExtraFieldLength As Integer
  FileCommentLength As Integer
  DiskNumberStart As Integer
  InternalFileAttributes As Integer
  ExternalFileAttributes As Long
  RelativeOffsetOfLocalHeader As Long
End Type

  ' file name (variable size)
  ' extra field (variable size)
  ' file comment (variable size)

Public Type DigitalSignature
  HeaderSignature As Long '(0x05054b50)
  SizeOfData As Integer
End Type

  ' signature data (variable size)
  
Public Type EndOfCentralDirectory
  EndCentralSignature As Long ' (0x06054b50)
  NumberThisDisk As Integer
  NumberTheDiskWithStart As Integer
  TotalNumberOfEntriesThisDisk As Integer
  TotalNumberOfEntries As Integer
  SizeOfCentralDirectory As Long
  OffsetOfStartCentralDirectory As Long
  ZIPFileCommentLength  As Integer
End Type
        
  ' ZIP file comment       (variable size)

Public Function ResolveMethod(Method As Integer) As String
Select Case Method
  Case 0: ResolveMethod = "[ No Compression ]"
  Case 1: ResolveMethod = "[ Shrunk ]"
  Case 2: ResolveMethod = "[ Factor 1 ]"
  Case 3: ResolveMethod = "[ Factor 2 ]"
  Case 4: ResolveMethod = "[ Factor 3 ]"
  Case 5: ResolveMethod = "[ Factor 4 ]"
  Case 6: ResolveMethod = "[ Imploded ]"
  Case 7: ResolveMethod = "[ Tokenized ]"
  Case 8: ResolveMethod = "[ Deflated ]"
  Case 9: ResolveMethod = "[ Enhanced Deflating ]"
  Case 10: ResolveMethod = "[ PKWARE Imploding ]"
End Select
End Function
