#Region " Shared Memory "

Public Class clsSharedMemory
    'APIs
    Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Integer, ByVal lpFileMappigAttributes As Integer, ByVal flProtect As Integer, ByVal dwMaximumSizeHigh As Integer, ByVal dwMaximumSizeLow As Integer, ByVal lpName As String) As Integer
    Private Declare Function MapViewOfFile Lib "kernel32" Alias "MapViewOfFile" (ByVal hFileMappingObject As Integer, ByVal dwDesiredAccess As Integer, ByVal dwFileOffsetHigh As Integer, ByVal dwFileOffsetLow As Integer, ByVal dwNumberOfBytesToMap As Integer) As IntPtr
    Private Declare Function UnmapViewOfFile Lib "kernel32" Alias "UnmapViewOfFile" (ByVal lpBaseAddress As IntPtr) As Integer
    Private Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Integer) As Integer

    'Constants
    Private Const FILE_MAP_ALL_ACCESS As Integer = &HF001F
    Private Const PAGE_READWRITE As Integer = &H4
    Private Const INVALID_HANDLE_VALUE As Integer = -1

    'Variables
    Private FileHandle As Integer
    Private SharePoint As IntPtr

#Region " Open and Close (Memory) Procedures "

    Public Function Open(ByVal MemoryName As String) As Boolean
        'Get a handle to an area of memory and name it the name passed in MemoryName.
        'Any application that maps an area of memory with that name gets the same 
        'address, so data can be shared.
        'Note: the INVALID_HANDLE_VALUE, which tells windows not to use a file but
        'just memory.
        FileHandle = CreateFileMapping(INVALID_HANDLE_VALUE, 0, PAGE_READWRITE, 0, 128, MemoryName)

        'Get a pointer to the area of memory we mapped.
        If Not FileHandle = 0 Then
            SharePoint = MapViewOfFile(FileHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub Close()
        'Close the memory handle.
        UnmapViewOfFile(SharePoint)
        CloseHandle(FileHandle)
    End Sub

    Protected Overrides Sub Finalize()
        'Close the memory handle
        Call Close()

        'Finalize Base Class
        MyBase.Finalize()
    End Sub

#End Region

    Public Function Peek() As String
        'Copy the data length to a variable.
        Dim myDataLength As Integer = Marshal.ReadInt32(SharePoint)

        'Create an array to hold the data in memory.
        Dim myBuffer(myDataLength - 1) As Byte

        'Copy the data in memory to the array. 'SharePoint.ToInt64
        Try
            Marshal.Copy(New IntPtr(SharePoint.ToInt32 + 4), myBuffer, 0, myDataLength)
        Catch ex As Exception
            Marshal.Copy(New IntPtr(SharePoint.ToInt64 + 4), myBuffer, 0, myDataLength)
        End Try

        'Return Output (Unicode)
        Return System.Text.Encoding.UTF8.GetString(myBuffer)
    End Function

    Public Sub Put(ByVal Data As String)
        'Create an array with one element for each character. (Unicode)
        Dim myBuffer As Byte() = System.Text.Encoding.UTF8.GetBytes(Data)

        'Copy the length of the string into the first four bytes of the memory location
        Marshal.WriteInt32(SharePoint, Data.Length)

        'Copy the string data to memory right after the length.
        Marshal.Copy(myBuffer, 0, New IntPtr(SharePoint.ToInt32 + 4), myBuffer.Length)
    End Sub

    Public Sub ResetMemory()
        'Reset Data Lenght (Set Data Length to 0 - Empty)
        Marshal.WriteInt32(SharePoint, 0I)
    End Sub

    Public ReadOnly Property DataExists() As Boolean
        Get
            'Copy the data length to a variable.
            Dim myDataLength As Integer
            myDataLength = Marshal.ReadInt32(SharePoint)
            If Not myDataLength = 0 Then Return True
            Return False
        End Get
    End Property

End Class

#End Region