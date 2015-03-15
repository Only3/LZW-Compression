
    Public Class LZW

        Private EOF As Integer = -1
        Private BITS As Integer = 14
        Private HASHING_SHIFT As Integer = 4
        Private TABLE_SIZE As Integer = 18041
        Private MAX_VALUE As Integer = (1 << BITS) - 1
        Private MAX_CODE As Integer = MAX_VALUE - 1
        Private AppendChar(TABLE_SIZE) As Byte
        Private CodeValue(TABLE_SIZE) As Integer
        Private PrefixCode(TABLE_SIZE) As Integer
        Public Input As IO.BinaryReader = Nothing
        Public Output As IO.BinaryWriter = Nothing

        Public Sub Compress(File As String)
            Input = New IO.BinaryReader(IO.File.Open(File, IO.FileMode.Open))
            Output = New IO.BinaryWriter(IO.File.Open(File.Substring(0, File.LastIndexOf(".")) & ".lzw", IO.FileMode.OpenOrCreate, IO.FileAccess.Write))
            Dim Index As Integer = 0, Character As Integer = 0
            Dim StringCode As Integer = 0, NextCode As Integer = 256
            For i = 0 To TABLE_SIZE - 1
                CodeValue(i) = -1
            Next
            StringCode = ReadByte()
            Character = ReadByte()
            While Character <> -1
                Index = Match(StringCode, Character)
                If CodeValue(Index) <> -1 Then
                    StringCode = CodeValue(Index)
                Else
                    If NextCode <= MAX_CODE Then
                        CodeValue(Index) = NextCode
                        NextCode += 1
                        PrefixCode(Index) = StringCode
                        AppendChar(Index) = CByte(Character)
                    End If
                    OutputCode(StringCode)
                    StringCode = Character
                End If
                Character = ReadByte()
            End While
            OutputCode(StringCode)
            OutputCode(MAX_VALUE)
            OutputCode(0)
            Output.Close()
        End Sub

        Public Sub Decompress(File As String)
            Input = New IO.BinaryReader(IO.File.Open(File, IO.FileMode.Open))
            Output = New IO.BinaryWriter(IO.File.Open(File.Substring(0, File.LastIndexOf(".")) & "New.exe", IO.FileMode.OpenOrCreate, IO.FileAccess.Write))
            Dim i As Integer
            Dim NewCode As Integer, OldCode As Integer
            Dim CurrCode As Integer, NextCode As Integer = 256
            Dim Character As Byte, DecodeStack(TABLE_SIZE) As Byte
            OldCode = InputCode()
            Character = CType(OldCode, Byte)
            Output.Write(CByte(OldCode))
            NewCode = InputCode()
            While NewCode <> MAX_VALUE
                If NewCode >= NextCode Then
                    DecodeStack(0) = Character
                    i = 1
                    CurrCode = OldCode
                Else
                    i = 0
                    CurrCode = NewCode
                End If
                While CurrCode > 255
                    DecodeStack(i) = AppendChar(CurrCode)
                    i = i + 1
                    If i >= MAX_CODE Then Throw New Exception("CurrCode Exception.")
                    CurrCode = PrefixCode(CurrCode)
                End While
                DecodeStack(i) = CType(CurrCode, Byte)
                Character = DecodeStack(i)
                While i >= 0
                    Output.Write(DecodeStack(i))
                    i = i - 1
                End While
                If NextCode <= MAX_CODE Then
                    PrefixCode(NextCode) = OldCode
                    AppendChar(NextCode) = Character
                    NextCode += 1
                End If
                OldCode = NewCode
                NewCode = InputCode()
            End While
            Input.Close()
        End Sub

        Private Sub OutputCode(Code As Integer)
            Static Buffer As Long = 0, Count As Integer = 0
            Buffer = Buffer Or (Code << (32 - BITS - Count))
            Count += BITS
            While Count >= 8
                Output.Write(CByte((Buffer >> 24) And 255))
                Buffer <<= 8
                Count -= 8
            End While
        End Sub

        Private Function InputCode() As Integer
            Dim Value As Long
            Static Buffer As Long = 0, Count As Integer = 0
            Static Mask32 As Long = CLng(2 ^ 32) - 1
            While Count <= 24
                Buffer = (Buffer Or ReadByte() << (24 - Count)) And Mask32
                Count += 8
            End While
            Value = (Buffer >> 32 - BITS) And Mask32
            Buffer = (Buffer << BITS) And Mask32
            Count -= BITS
            Return CInt(Value)
        End Function

        Private Function Match(Prefix As Integer, Character As Integer) As Integer
            Dim Index As Integer = 0, Offset As Integer = 0
            Index = CInt((Character << HASHING_SHIFT) Xor Prefix)
            If Index = 0 Then Offset = 1 Else Offset = TABLE_SIZE - Index
            While True
                If CodeValue(Index) = -1 Then Return Index
                If PrefixCode(Index) = Prefix And AppendChar(Index) = Character Then Return Index
                Index -= Offset
                If Index < 0 Then Index += TABLE_SIZE
            End While
            Return Nothing
        End Function

        Private Function ReadByte() As Integer
            Dim B(1) As Byte
            If Input.Read(B, 0, 1) = 0 Then Return -1
            Return B(0)
        End Function
    End Class
