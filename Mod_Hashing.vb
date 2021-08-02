Imports System.Security.Cryptography

Module Mod_Hashing
    Public Function SHA512(ByRef Data As String) As String
        Using SHA_512 As New SHA512Managed
            SHA_512.Initialize()
            SHA_512.ComputeHash(GetBytes(Data))
            Return GetHex(SHA_512.Hash)
        End Using
    End Function

    Public Function SHA256(ByVal Password As String) As Byte()
        Using SHA_256 As New SHA256Managed
            SHA_256.Initialize()
            SHA_256.ComputeHash(GetBytes(Password))
            Return SHA_256.Hash
        End Using
    End Function

    Private Function GetHex(ByRef Data() As Byte) As String
        Return BitConverter.ToString(Data).Replace("-", "")
    End Function

    Private Function GetBytes(ByRef Data As String) As Byte()
        Return System.Text.Encoding.Unicode.GetBytes(Data)
    End Function

    Private Function T2S(ByVal PlainText As String) As Byte()
        Return System.Text.Encoding.UTF8.GetBytes(PlainText)
    End Function

    Private Function S2T(ByVal Data() As Byte) As String
        Return System.Text.Encoding.UTF8.GetChars(Data)
    End Function
End Module