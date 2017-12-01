Imports System
Imports System.IO
Imports System.Collections
Module Module1
    Dim input As String = Nothing
    Sub Main()
        Console.ForegroundColor = ConsoleColor.Green
        log("*************************************************************************")
        log("*************************************************************************")
        log("******************         File encryption         **********************")
        log("*****************      By parfait mutshipayi    *************************")
        log("                                                                         ")
        log("*************************************************************************")
        log("*************************************************************************")
        log(" ")
        log(" ")
        Console.ForegroundColor = ConsoleColor.Red
        
        Console.ForegroundColor = ConsoleColor.Green

        While True
            log("1) Encrypt a directory")
            log("2) Decrypt a directory")
            log("3) Exit the program")
            log("")
            Console.Write("SELECT OPTION :")
            input = Console.ReadLine()
            If input = "1" Or input = "2" Or input = "3" Then
                Select Case input
                    Case 1
                        menu1()
                    Case 2
                        menu2()
                    Case 3
                        Environment.Exit("0")
                End Select
            Else
                Console.ForegroundColor = ConsoleColor.Red
                log("Invalid selection")
                Console.ForegroundColor = ConsoleColor.Green
            End If
        End While
        log("DONE!!!")
        Console.ReadKey()
    End Sub
    Sub menu1()
        Console.ForegroundColor = ConsoleColor.Red
        log("1) EXIT")
        Console.ForegroundColor = ConsoleColor.Green
        While True
            Console.Write("Enter the path :")
            input = Console.ReadLine()
            Select Case input
                Case 1
                    Exit While
                Case Else
                    If My.Computer.FileSystem.DirectoryExists(input) Then
                        ' Encrypt sub dir later
                        Dim dir As String = input
                        Console.Write("Enter the encryption password :")
                        input = Console.ReadLine()
                        Dim result = crypt_dir(dir, input)
                        ' result(0) number of file found
                        ' result(1) number of file crypted
                        log("Please wait...")
                        Console.ForegroundColor = ConsoleColor.Yellow
                        log(result(0) & " file(s) found in the directory")
                        log(result(1) & " file(s) encrypted")
                        Console.ForegroundColor = ConsoleColor.Green
                        log("Encryption Done!")
                    Else
                        ' The directory specified doesn't exist
                        Console.ForegroundColor = ConsoleColor.Red
                        log("The path " & input & " Doesn't exist")
                        Console.ForegroundColor = ConsoleColor.Green
                    End If
            End Select
        End While
    End Sub
    Sub menu2()
        Console.ForegroundColor = ConsoleColor.Red
        log("1) EXIT")
        Console.ForegroundColor = ConsoleColor.Green
        While True
            Console.Write("Enter the path :")
            input = Console.ReadLine()
            Select Case input
                Case 1
                    Exit While
                Case Else
                    If My.Computer.FileSystem.DirectoryExists(input) Then
                        Dim dir As String = input
                        Console.Write("Enter the decryption password : ")
                        input = Console.ReadLine()
                        Dim des = decrypt_dir(dir, input)
                        log("Please wait...")
                        Console.ForegroundColor = ConsoleColor.Yellow
                        log(des(0) & " File(s) found")
                        log(des(1) & " file(s) decrypted ")
                        Console.ForegroundColor = ConsoleColor.Green
                    Else
                        Console.ForegroundColor = ConsoleColor.Red
                        log(input & " Not found !")
                        Console.ForegroundColor = ConsoleColor.Green
                    End If
            End Select
        End While
    End Sub
    Sub log(ByVal msg As Object)
        Console.WriteLine(msg)
    End Sub
    Function crypt_dir(ByVal dir As String, ByVal pwd As String) As ArrayList
        Dim result As New ArrayList
        Dim _file = My.Computer.FileSystem.GetFiles(dir)
        Dim count_file As Integer = 0
        Dim count_crypted_file As Integer = 0
        For Each File In _file
            If isCryptableFile(File) Then
                Try
                    Dim encrypter As New Crypto(pwd)
                    Dim crypt_name_file As String = File.split("\")(File.split("\").length - 1)

                    If isPlainTextFile(File) Then
                        ' Encrypt plain text file
                        My.Computer.FileSystem.WriteAllText(File, encrypter.EncryptData(IO.File.ReadAllText(File)), False)
                        count_crypted_file += 1
                    Else
                        ' Dealing with a binary file
                        Dim all_bytes() As Byte = IO.File.ReadAllBytes(File)
                        My.Computer.FileSystem.WriteAllBytes(File, encrypter.EncryptData(all_bytes, False), False)
                        count_crypted_file += 1
                    End If
                Catch ex As Exception
                    log("ERROR ==> " & ex.Message)
                End Try
            End If
            count_file += 1
        Next
        result.Add(count_file)
        result.Add(count_crypted_file)
        Return result
    End Function
    Function decrypt_dir(ByVal dir As String, ByVal pwd As String) As ArrayList
        Dim result As New ArrayList
        Dim count_encrypted_file As Integer = 0
        Dim count_des_file As Integer = 0
        Dim _file = My.Computer.FileSystem.GetFiles(dir)
        For Each File In _file
            count_encrypted_file += 1
            If isCryptableFile(File) Then
                Dim decrypter As New Crypto(pwd)
                Dim crypt_name_file As String = File.Split("\")(File.Split("\").Length - 1)
                Try
                    If isPlainTextFile(File) Then
                        Dim des As Object = decrypter.DecryptData(IO.File.ReadAllText(File))
                        Dim writer As New StreamWriter(File)
                        writer.Write(des)
                        writer.Close()
                        count_des_file += 1
                    Else
                        ' Decrypt a binary file
                        My.Computer.FileSystem.WriteAllBytes(File, decrypter.DecryptData(IO.File.ReadAllBytes(File), False), False)
                        ' rename the encrypted file
                        'My.Computer.FileSystem.RenameFile(File, crypt_name_file.Replace(".encrypted", ""))
                        count_des_file += 1
                    End If
                Catch err As System.Security.Cryptography.CryptographicException
                    log("INVALID PASSWORD")
                Catch ex As Exception
                    log("***************************")
                    log("COULDN'T DECRYPT " & File)
                    log(ex.Message)
                    log("****************************")
                End Try
            End If
        Next
        result.Add(count_encrypted_file)
        result.Add(count_des_file)
        Return result
    End Function
    Function isCryptableFile(ByVal filename As String) As Boolean
        ' Encrypt only following file's format only.
        ' Images 
        '      png, jpeg, jpg, bwm, gif, ico
        ' Audio
        '     mp3, mp4, acc
        ' video
        '     .avi, mkw
        ' Others
        '   txt doc, docx, ppt etc
        ' Program 
        '    .exe, .msi
        ' Archive
        '   .rar, .zip
        Dim cryptable_file_format() As String = IO.File.ReadAllLines("cryptable_file_format.txt")
        For Each file_format In cryptable_file_format
            If filename.ToLower().EndsWith("." & file_format.ToLower().Trim()) Then
                Return True
            End If
        Next
        Return False
    End Function
    Function isPlainTextFile(ByVal filename As String) As Boolean
        ' Check if a file is a plain or it's a binary such as .exe, png file etc..
        Dim plaintext_format() As String = IO.File.ReadAllLines("plain_text_format.txt")
        For Each Format As String In plaintext_format
            If Format <> Nothing Then
                If filename.ToLower().EndsWith("." & Format.ToLower()) Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function
End Module
