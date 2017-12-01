Imports System.Security.Cryptography
Public Class Crypto
    Private TripleDes As New TripleDESCryptoServiceProvider
    Private Function CreateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider
        ' Hash key
        Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)
        'Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function
    Sub New(ByVal key As String)
        ' Initialize the crypto provider
        TripleDes.Key = CreateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = CreateHash("", TripleDes.BlockSize \ 8)
    End Sub
    Public Function EncryptData(ByVal plaintext As String) As String
        ' Convert the plaintext string to a byte array
        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)
        Return EncryptData(plaintextBytes, True)
    End Function
    Public Function EncryptData(ByVal plaintextBytes() As Byte, ByVal returnString As Boolean) As Object
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
        ' use the crypto stream to write the byte array to the stream
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()
        ' convert the encrypted stream to a printable string
        If returnString Then
            Return Convert.ToBase64String(ms.ToArray)
        Else
            Return ms.ToArray
        End If
    End Function
    Public Function DecryptData(ByVal encryptedtext As String) As String
        Dim encryptedbytes() As Byte = Convert.FromBase64String(encryptedtext)
        Return DecryptData(encryptedbytes, True)
    End Function
    Public Function DecryptData(ByVal encryptedbytes As Byte(), ByVal returnString As Boolean) As Object
        ' Create stream
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream
        Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
        ' Use the crypto stream to write the byte array to the stream
        decStream.Write(encryptedbytes, 0, encryptedbytes.Length)
        decStream.FlushFinalBlock()

        'convert the plaintext stream to string
        If returnString Then
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        Else
            Return ms.ToArray
        End If
    End Function
End Class
