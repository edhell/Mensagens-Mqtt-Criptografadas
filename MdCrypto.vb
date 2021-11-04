
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
Imports System.Configuration

Module MdCrypto


#Region "Cifra Cesar"

    Public Function cifraCesarCodificar(mensagem As String, deslocamento As Integer)
        Dim auxiliar As String
        Dim resultado As String

        Try 'Tratamento de erro
            For i As Integer = 0 To mensagem.Length - 1 'Loop que percorrará a string inteira do texto
                auxiliar = Asc(mensagem(i)) 'Pega o número Asc do Caracter

                If (auxiliar >= 65 And auxiliar <= 90) Then 'Se o caracter for MAISCULO
                    auxiliar = auxiliar + deslocamento ' Soma o Asc com a Chave digitada

                    If (auxiliar > 90) Then ' Se a conversão passar do Z
                        auxiliar = auxiliar - 26 'Volta para o Inicio do alfabeto
                    End If

                    resultado = resultado + Chr(auxiliar) 'Salva o caracter convertido na nova string

                ElseIf (auxiliar >= 97 And auxiliar <= 122) Then ' Se o caracter for MINUSCULO
                    auxiliar = auxiliar + deslocamento ' Soma o Asc com a chave digitada

                    If (auxiliar > 122) Then ' Se a conversão passar de z
                        auxiliar = auxiliar - 26 'Volta ao ínicio do alfabeto
                    End If

                    resultado = resultado + Chr(auxiliar) 'Salva o caracter convertido na nova string

                Else 'Caso o caracter não foi uma letra
                    resultado = resultado + Chr(auxiliar) 'Salva o caracter na nova string
                End If

            Next 'Fim do Loop

            Return resultado 'O tbResultado mostra a string resultado.
        Catch ex As Exception
            MsgBox("Erro: " & ex.Message)
            'MsgBox("Preencha os campos necessários", , "Cifra de César") 'Campo da chave não pode estar vazio
        End Try
    End Function

    Public Function cifraCesarDecodificar(mensagem As String, deslocamento As Integer)
        Dim auxiliar As String
        Dim resultado As String

        Try 'Tratamento de erro
            For i As Integer = 0 To mensagem.Length - 1 'Loop que percorrará a string inteira do texto
                auxiliar = Asc(mensagem(i)) 'Pega o número Asc do Caracter

                If (auxiliar >= 65 And auxiliar <= 90) Then 'Se o caracter for MAISCULO
                    auxiliar = auxiliar - deslocamento ' Soma o Asc com a Chave digitada

                    If (auxiliar < 65) Then ' Se a conversão passar do Z
                        auxiliar = auxiliar + 26 'Vai para o fim do alfabeto
                    End If

                    resultado = resultado + Chr(auxiliar) 'Salva o caracter convertido na nova string

                ElseIf (auxiliar >= 97 And auxiliar <= 122) Then ' Se o caracter for MINUSCULO
                    auxiliar = auxiliar - deslocamento ' Soma o Asc com a chave digitada

                    If (auxiliar < 97) Then ' Se a conversão passar de z
                        auxiliar = auxiliar + 26 'Vai para o fim do alfabeto
                    End If

                    resultado = resultado + Chr(auxiliar) 'Salva o caracter convertido na nova string

                Else 'Caso o caracter não foi uma letra
                    resultado = resultado + Chr(auxiliar) 'Salva o caracter na nova string
                End If

            Next 'Fim do Loop

            Return resultado 'O tbResultado mostra a string resultado.
        Catch ex As Exception
            MsgBox("Erro: " & ex.Message)
            'MsgBox("Preencha os campos necessários", , "Cifra de César") 'Campo da chave não pode estar vazio
        End Try
    End Function

#End Region

#Region "Cifra Vigenere"

    Public Function cifraVigenereCodificar(mensagem As String, chave As String)
        Return cifraVigenere(mensagem, chave, True)
    End Function

    Public Function cifraVigenereDecodificar(mensagem As String, chave As String)
        Return cifraVigenere(mensagem, chave, False)
    End Function

    Private Function [ModCV](a As Integer, b As Integer) As Integer
        Return (a Mod b + b) Mod b
    End Function

    Private Function cifraVigenere(input As String, key As String, encipher As Boolean) As String
        For i As Integer = 0 To key.Length - 1
            If Not Char.IsLetter(key(i)) Then
                Return Nothing ' Error
            End If
        Next

        Dim output As String = String.Empty
        Dim nonAlphaCharCount As Integer = 0

        For i As Integer = 0 To input.Length - 1
            If Char.IsLetter(input(i)) Then
                Dim cIsUpper As Boolean = Char.IsUpper(input(i))
                Dim offset As Integer = Convert.ToInt32(If(cIsUpper, "A"c, "a"c))
                Dim keyIndex As Integer = (i - nonAlphaCharCount) Mod key.Length
                Dim k As Integer = Convert.ToInt32(If(cIsUpper, Char.ToUpper(key(keyIndex)), Char.ToLower(key(keyIndex)))) - offset
                k = If(encipher, k, -k)
                Dim ch As Char = ChrW(([ModCV](((Convert.ToInt32(input(i)) + k) - offset), 26)) + offset)
                output += ch
            Else
                output += input(i)
                nonAlphaCharCount += 1
            End If
        Next

        Return output
    End Function

#End Region

#Region "Cifra OTP"

    Public Function cifraOTPCodificar(mensagem As String, chave As String)
        If mensagem.Length() <> chave.Length() Then
            MsgBox("O tamanho da mensagem e da chave devem ser iguais.")
            Return ""
        End If
        'Dim im As Integer() = charArrayToInt(mensagem.ToCharArray())
        'Dim ik As Integer() = charArrayToInt(chave.ToCharArray())
        Dim data As Integer() = New Integer(mensagem.Length() - 1) {}

        For i As Integer = 0 To mensagem.Length() - 1
            data(i) = Convert.ToInt32(mensagem(i)) + Convert.ToInt32(chave(i))
        Next

        Dim retorno As String
        For Each dt In data
            retorno += ChrW(dt.ToString)
        Next

        Return retorno
    End Function

    Public Function cifraOTPDecodificar(mensagem As String, chave As String)
        If mensagem.Length() <> chave.Length() Then
            MsgBox("O tamanho da mensagem e da chave devem ser iguais.")
            Return ""
        End If
        'Dim im As Integer() = charArrayToInt(mensagem.ToCharArray())
        'Dim ik As Integer() = charArrayToInt(chave.ToCharArray())
        Dim data As Integer() = New Integer(mensagem.Length() - 1) {}

        For i As Integer = 0 To mensagem.Length() - 1
            data(i) = Convert.ToInt32(mensagem(i)) - Convert.ToInt32(chave(i))
        Next

        Dim retorno As String
        For Each dt In data
            retorno += ChrW(dt.ToString)
        Next

        Return retorno
    End Function

    Public Function cifraOTPGenKey(ByVal tamanho As Integer) As String
        Dim randomico As Random = New Random()
        Dim key As Char() = New Char(tamanho - 1) {}

        For i As Integer = 0 To tamanho - 1
            key(i) = ChrW(randomico.Next(132))
            If Convert.ToInt32(key(i)) < 97 Then key(i) = ChrW((Convert.ToInt32(key(i)) + 72))
            If Convert.ToInt32(key(i)) > 122 Then key(i) = ChrW((Convert.ToInt32(key(i)) - 72))
        Next

        Return New String(key)
    End Function

#End Region

#Region "Cifra Playfair"

    Private Function [ModCP](a As Integer, b As Integer) As Integer
        Return (a Mod b + b) Mod b
    End Function

    Private Function FindAllOccurrences(str As String, value As Char) As List(Of Integer)
        Dim indexes As New List(Of Integer)()

        Dim index As Integer = 0
        While index <> -1
            index = str.IndexOf(value, index)
            If index <> -1 Then
                indexes.Add(index)
                index += 1
            End If
        End While

        Return indexes
    End Function

    Private Function RemoveAllDuplicates(str As String, indexes As List(Of Integer)) As String
        Dim retVal As String = str

        For i As Integer = indexes.Count - 1 To 1 Step -1
            retVal = retVal.Remove(indexes(i), 1)
        Next

        Return retVal
    End Function

    Private Function GenerateKeySquare(key As String) As Char(,)
        Dim keySquare As Char(,) = New Char(4, 4) {}
        Dim defaultKeySquare As String = "ABCDEFGHIKLMNOPQRSTUVWXYZ"
        Dim tempKey As String = If(String.IsNullOrEmpty(key), "CIPHER", key.ToUpper())

        tempKey = tempKey.Replace("J", "")
        tempKey += defaultKeySquare

        For i As Integer = 0 To 24
            Dim indexes As List(Of Integer) = FindAllOccurrences(tempKey, defaultKeySquare(i))
            tempKey = RemoveAllDuplicates(tempKey, indexes)
        Next

        tempKey = tempKey.Substring(0, 25)

        For i As Integer = 0 To 24
            keySquare((i \ 5), (i Mod 5)) = tempKey(i)
        Next

        Return keySquare
    End Function

    Private Sub GetPosition(ByRef keySquare As Char(,), ch As Char, ByRef row As Integer, ByRef col As Integer)
        If ch = "J"c Then
            GetPosition(keySquare, "I"c, row, col)
        End If

        For i As Integer = 0 To 4
            For j As Integer = 0 To 4
                If keySquare(i, j) = ch Then
                    row = i
                    col = j
                End If
            Next
        Next
    End Sub

    Private Function SameRow(ByRef keySquare As Char(,), row As Integer, col1 As Integer, col2 As Integer, encipher As Integer) As Char()
        Return New Char() {keySquare(row, [ModCP]((col1 + encipher), 5)), keySquare(row, [ModCP]((col2 + encipher), 5))}
    End Function

    Private Function SameColumn(ByRef keySquare As Char(,), col As Integer, row1 As Integer, row2 As Integer, encipher As Integer) As Char()
        Return New Char() {keySquare([ModCP]((row1 + encipher), 5), col), keySquare([ModCP]((row2 + encipher), 5), col)}
    End Function

    Private Function SameRowColumn(ByRef keySquare As Char(,), row As Integer, col As Integer, encipher As Integer) As Char()
        Return New Char() {keySquare([ModCP]((row + encipher), 5), [ModCP]((col + encipher), 5)), keySquare([ModCP]((row + encipher), 5), [ModCP]((col + encipher), 5))}
    End Function

    Private Function DifferentRowColumn(ByRef keySquare As Char(,), row1 As Integer, col1 As Integer, row2 As Integer, col2 As Integer) As Char()
        Return New Char() {keySquare(row1, col2), keySquare(row2, col1)}
    End Function

    Private Function RemoveOtherChars(input As String) As String
        Dim output As String = input

        Dim i As Integer = 0
        While i < output.Length
            If Not Char.IsLetter(output(i)) Then
                output = output.Remove(i, 1)
            End If
            i += 1
        End While

        Return output
    End Function

    Private Function AdjustOutput(input As String, output As String) As String
        Dim retVal As New StringBuilder(output)

        For i As Integer = 0 To input.Length - 1
            If Not Char.IsLetter(input(i)) Then
                retVal = retVal.Insert(i, input(i).ToString())
            End If

            If Char.IsLower(input(i)) Then
                retVal(i) = Char.ToLower(retVal(i))
            End If
        Next

        Return retVal.ToString()
    End Function

    Private Function Cipher(input As String, key As String, encipher As Boolean) As String
        Dim retVal As String = String.Empty
        Dim keySquare As Char(,) = GenerateKeySquare(key)
        Dim tempInput As String = RemoveOtherChars(input)
        Dim e As Integer = If(encipher, 1, -1)

        If (tempInput.Length Mod 2) <> 0 Then
            tempInput += "X"
            'tempInput += " "
        End If

        For i As Integer = 0 To tempInput.Length - 1 Step 2
            Dim row1 As Integer = 0
            Dim col1 As Integer = 0
            Dim row2 As Integer = 0
            Dim col2 As Integer = 0

            GetPosition(keySquare, Char.ToUpper(tempInput(i)), row1, col1)
            GetPosition(keySquare, Char.ToUpper(tempInput(i + 1)), row2, col2)

            If row1 = row2 AndAlso col1 = col2 Then
                retVal += New String(SameRowColumn(keySquare, row1, col1, e))
            ElseIf row1 = row2 Then
                retVal += New String(SameRow(keySquare, row1, col1, col2, e))
            ElseIf col1 = col2 Then
                retVal += New String(SameColumn(keySquare, col1, row1, row2, e))
            Else
                retVal += New String(DifferentRowColumn(keySquare, row1, col1, row2, col2))
            End If
        Next

        retVal = AdjustOutput(input, retVal)

        Return retVal
    End Function

    Public Function cifraPlayfairCodificar(input As String, key As String) As String
        Return Cipher(input, key, True)
    End Function

    Public Function cifraPlayfairDecodificar(input As String, key As String) As String
        Return Cipher(input, key, False)
    End Function
#End Region

#Region "Cifra de Hill"

    Private c1, c2 As Integer
    Private c As Char
    Private cryptMatriceGlobal As Integer()

    Public Function cifraHillCodificar(mensagem As String, ByVal cryptMatrice As Integer())
        Dim finalMessage As String = ""

        For i As Integer = 0 To mensagem.Count() - 1 Step 2

            If i + 2 > mensagem.Count() Then
                c1 = (cryptMatrice(0) * (Convert.ToInt32(mensagem(i)) - 96) + cryptMatrice(1) * (Convert.ToInt32("a") - 96)) Mod 26
                c2 = (cryptMatrice(2) * (Convert.ToInt32(mensagem(i)) - 96) + cryptMatrice(3) * (Convert.ToInt32("a") - 96)) Mod 26
            Else
                c1 = (cryptMatrice(0) * (Convert.ToInt32(mensagem(i)) - 96) + cryptMatrice(1) * (Convert.ToInt32(mensagem(i + 1)) - 96)) Mod 26
                c2 = (cryptMatrice(2) * (Convert.ToInt32(mensagem(i)) - 96) + cryptMatrice(3) * (Convert.ToInt32(mensagem(i + 1)) - 96)) Mod 26
            End If

            c = Convert.ToChar(c1 + 96)
            finalMessage += c
            c = Convert.ToChar(c2 + 96)
            finalMessage += c
        Next

        Return finalMessage
    End Function

    Public Function cifraHillDecodificar(mensagem As String, ByVal cryptMatrice As Integer())
        cryptMatriceGlobal = cryptMatrice
        Dim decryptMatrice As Integer() = DecryptMatriceFunc()
        Dim finalMessage As String = ""

        For i As Integer = 0 To mensagem.Count() - 1 - 1 Step 2
            c1 = (decryptMatrice(0) * Convert.ToInt32(mensagem(i)) - 96) + decryptMatrice(1) * ((Convert.ToInt32(mensagem(i + 1))) - 96) Mod 26
            c2 = (decryptMatrice(2) * Convert.ToInt32(mensagem(i)) - 96) + decryptMatrice(3) * ((Convert.ToInt32(mensagem(i + 1)) + 1) - 96) Mod 26
            c = Convert.ToChar(c1 + 96)
            finalMessage += c
            c = Convert.ToChar(c2 + 96)
            finalMessage += c
        Next

        Return finalMessage
    End Function

    Private Function DetM() As Integer
        Return cryptMatriceGlobal(0) * cryptMatriceGlobal(3) - cryptMatriceGlobal(1) * cryptMatriceGlobal(2)
    End Function

    Private Function InverseModuloKK() As Integer
        Dim primeNumbers As Integer() = {1, 3, 5, 7, 9, 11, 15, 17, 19, 21, 23, 25}
        Dim determinant As Integer = DetM()
        Dim result As Integer = 0
        Dim i As Integer = 0

        While result <> 1
            i += 1
            result = determinant * primeNumbers(i) Mod 26
        End While

        Return primeNumbers(i)
    End Function

    Private Function InverseMatriceC() As Integer()
        Return New Integer() {cryptMatriceGlobal(3), -cryptMatriceGlobal(1), -cryptMatriceGlobal(2), cryptMatriceGlobal(0)}
    End Function

    Private Function DecryptMatriceFunc() As Integer()
        Dim inverseMatrice As Integer() = InverseMatriceC()
        Dim inverseModuloK As Integer = InverseModuloKK()
        Dim dMatrice As Integer() = {inverseMatrice(0) * inverseModuloK Mod 26, inverseMatrice(1) * inverseModuloK Mod 26, inverseMatrice(2) * inverseModuloK Mod 26, inverseMatrice(3) * inverseModuloK Mod 26}

        For i As Integer = 0 To dMatrice.Length - 1
            If dMatrice(i) < 0 Then dMatrice(i) = dMatrice(i) + 26
        Next

        Return dMatrice
    End Function


#End Region









#Region "TripleDES"

    ' <summary>
    ''' Encriptação simples usando TripleDES
    ''' (Triple Data Encryption Standard)
    ''' </summary>
    Public Class CryptoTripleDES

        Private Shared TripleDES As New TripleDESCryptoServiceProvider
        Private Shared MD5 As New MD5CryptoServiceProvider

        ' Definição da chave de encriptação/desencriptação
        'Private Const key As String = "CHAVE12345teste"

        ''' <summary>
        ''' Calcula o MD5 Hash 
        ''' </summary>
        ''' <param name="value">Chave</param>
        Public Shared Function MD5Hash(ByVal value As String) As Byte()

            ' Converte a chave para um array de bytes 
            Dim byteArray() As Byte = ASCIIEncoding.ASCII.GetBytes(value)
            Return MD5.ComputeHash(byteArray)

        End Function

        ''' <summary>
        ''' Encripta uma string com base em uma chave
        ''' </summary>
        ''' <param name="stringToEncrypt">String a encriptar</param>
        ''' <param name="chave">Chave</param>
        Public Shared Function Encrypt(ByVal stringToEncrypt As String, ByVal chave As String) As String
            Try
                ' Definição da chave e da cifra (que neste caso é Electronic
                ' Codebook, ou seja, encriptação individual para cada bloco)
                TripleDES.Key = CryptoTripleDES.MD5Hash(chave)
                TripleDES.Mode = CipherMode.ECB

                ' Converte a string para bytes e encripta
                Dim Buffer As Byte() = ASCIIEncoding.ASCII.GetBytes(stringToEncrypt)
                Return Convert.ToBase64String(TripleDES.CreateEncryptor().TransformFinalBlock(Buffer, 0, Buffer.Length))

            Catch ex As Exception
                MessageBox.Show(ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

                Return String.Empty

            End Try

        End Function

        ''' <summary>
        ''' Desencripta uma string com base em uma chave
        ''' </summary>
        ''' <param name="encryptedString">String a decriptar</param>
        ''' <param name="chave">Chave</param>
        Public Shared Function Decrypt(ByVal encryptedString As String, ByVal chave As String) As String
            Try
                ' Definição da chave e da cifra
                TripleDES.Key = CryptoTripleDES.MD5Hash(chave)
                TripleDES.Mode = CipherMode.ECB

                ' Converte a string encriptada para bytes e decripta
                Dim Buffer As Byte() = Convert.FromBase64String(encryptedString)
                Return ASCIIEncoding.ASCII.GetString(TripleDES.CreateDecryptor().TransformFinalBlock(Buffer, 0, Buffer.Length))

            Catch ex As Exception
                MessageBox.Show(ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)

                Return String.Empty

            End Try

        End Function


    End Class

#End Region


    'simetrica 'https://israelaece.files.wordpress.com/2016/06/chapter08.pdf



    ' q tipo é esse
    Public Function Crypt(Text As String) As String

        Dim strTempChar As String

        For i = 1 To Len(Text)

            If Asc(Mid$(Text, i, 1)) < 128 Then
                strTempChar = Asc(Mid$(Text, i, 1)) + 128
            ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
                strTempChar = Asc(Mid$(Text, i, 1)) - 128
            End If

            Mid$(Text, i, 1) = Chr(strTempChar)

        Next i

        Crypt = Text

    End Function

    Public Function DES()
        Using csp As New DESCryptoServiceProvider()
            Dim buffer() As Byte = Nothing
            Using ms As New MemoryStream()
                Using stream As New CryptoStream(ms, csp.CreateEncryptor, CryptoStreamMode.Write)
                    Using sw As New StreamWriter(stream)
                        sw.Write("Trabalhando com Criptografia")
                    End Using
                End Using
                buffer = ms.ToArray()
                Console.WriteLine(Encoding.Default.GetString(buffer))
            End Using
            Using ms As New MemoryStream(buffer)
                Using stream As New CryptoStream(ms, csp.CreateDecryptor(), CryptoStreamMode.Read)
                    Using sr As New StreamReader(stream)
                        Console.WriteLine(sr.ReadLine())
                    End Using
                End Using
            End Using
        End Using
    End Function

    Public Function asimetrica() 'chave publica
        Dim hash As Byte() = {22, 45, 78, 53, 1, 2, 205, 98, 75, 123, 45, 76, 143, 189, 205, 65, 12, 193, 211, 255}
        Dim algoritmo As String = "SHA1"
        Using csp As New DSACryptoServiceProvider
            Dim formatter As New DSASignatureFormatter(csp)
            formatter.SetHashAlgorithm(algoritmo)
            Dim signedHash As Byte() = formatter.CreateSignature(hash)
            Dim deformatter As New DSASignatureDeformatter(csp)
            deformatter.SetHashAlgorithm(algoritmo)
            If deformatter.VerifySignature(hash, signedHash) Then
                Console.WriteLine("Assinatura válida.")
            Else
                Console.WriteLine("Assinatura inválida.")
            End If
        End Using

    End Function


    Public Function rca()
        Using csp As New RSACryptoServiceProvider
            Dim msg As Byte() = Encoding.Default.GetBytes("Trabalhando com criptografia")
            Dim msgEncriptada As Byte() = csp.Encrypt(msg, True)
            Console.WriteLine(Encoding.Default.GetString(msgEncriptada))
            Dim msgDescriptada As Byte() = csp.Decrypt(msgEncriptada, True)
            Console.WriteLine(Encoding.Default.GetString(msgDescriptada))
        End Using

    End Function

    Public Function sha1Teste()
        Dim csp As SHA1CryptoServiceProvider = DirectCast(CryptoConfig.CreateFromName("SHA"), SHA1CryptoServiceProvider)

    End Function

    Public Function md5()
        Using csp As New MD5CryptoServiceProvider
            Dim msg As Byte() = Encoding.Default.GetBytes("MinhaSenha")
            Dim hash As Byte() = csp.ComputeHash(msg)
            For i As Integer = 0 To hash.Length - 1
                Console.Write(hash(i).ToString("x2"))
            Next
        End Using
    End Function


#Region "3DES"
    Public Function TripleDESEncrypt(ByVal toEncrypt As String, ByVal useHashing As Boolean) As String
        Dim keyArray As Byte()
        Dim toEncryptArray As Byte() = UTF8Encoding.UTF8.GetBytes(toEncrypt)
        Dim settingsReader As System.Configuration.AppSettingsReader = New AppSettingsReader()
        'Dim key As String = CStr(settingsReader.GetValue("SecurityKey", GetType(String)))
        Dim key As String = "TESTE123"

        If useHashing Then
            Dim hashmd5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
            keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key))
            hashmd5.Clear()
        Else
            keyArray = UTF8Encoding.UTF8.GetBytes(key)
        End If

        Dim tdes As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tdes.Key = keyArray
        tdes.Mode = CipherMode.ECB
        tdes.Padding = PaddingMode.PKCS7
        Dim cTransform As ICryptoTransform = tdes.CreateEncryptor()
        Dim resultArray As Byte() = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length)
        tdes.Clear()
        Return Convert.ToBase64String(resultArray, 0, resultArray.Length)
    End Function

    Public Function TripleDESDecrypt(ByVal cipherString As String, ByVal useHashing As Boolean) As String
        Dim keyArray As Byte()
        Dim toDecryptArray As Byte() = Convert.FromBase64String(cipherString)
        Dim settingsReader As System.Configuration.AppSettingsReader = New AppSettingsReader()
        'Dim key As String = CStr(settingsReader.GetValue("SecurityKey", GetType(String)))
        Dim key As String = "TESTE123"

        If useHashing Then
            Dim hashmd5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
            keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key))
            hashmd5.Clear()
        Else
            keyArray = UTF8Encoding.UTF8.GetBytes(key)
        End If

        Dim tdes As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tdes.Key = keyArray
        tdes.Mode = CipherMode.ECB
        tdes.Padding = PaddingMode.PKCS7
        Dim cTransform As ICryptoTransform = tdes.CreateDecryptor()
        Dim resultArray As Byte() = cTransform.TransformFinalBlock(toDecryptArray, 0, toDecryptArray.Length)
        tdes.Clear()
        Return UTF8Encoding.UTF8.GetString(resultArray)
    End Function

#End Region


#Region "AES"
    'https://www.ti-enxame.com/pt/vb.net/aes-encriptar-cadeia-em-vb.net/973197818/
    'AES NORMAL
    Public Function AES_Encrypt(ByVal input As String, ByVal pass As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim encrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = Security.Cryptography.CipherMode.ECB
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
            encrypted = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return encrypted
        Catch ex As Exception
        End Try
    End Function
    'AES NORMAL
    Public Function AES_Decrypt(ByVal input As String, ByVal pass As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim decrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = Security.Cryptography.CipherMode.ECB
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = Convert.FromBase64String(input)
            decrypted = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return decrypted
        Catch ex As Exception
        End Try
    End Function


    ''AES-CBC
    'Private Function AESE(ByVal plaintext As String, ByVal key As String) As String
    '    Dim AES As New System.Security.Cryptography.RijndaelManaged
    '    Dim SHA256 As New System.Security.Cryptography.SHA256Cng
    '    Dim ciphertext As String = ""
    '    Try
    '        AES.GenerateIV()
    '        AES.Key = SHA256.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(key))

    '        AES.Mode = Security.Cryptography.CipherMode.CBC
    '        Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
    '        Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(plaintext)
    '        ciphertext = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))

    '        Return Convert.ToBase64String(AES.IV) & Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))

    '    Catch ex As Exception
    '        Return ex.Message
    '    End Try
    'End Function
    ''AES-CBC
    'Private Function AESD(ByVal ciphertext As String, ByVal key As String) As String
    '    Dim AES As New System.Security.Cryptography.RijndaelManaged
    '    Dim SHA256 As New System.Security.Cryptography.SHA256Cng
    '    Dim plaintext As String = ""
    '    Dim iv As String = ""
    '    Try
    '        Dim ivct = ciphertext.Split({"=="}, StringSplitOptions.None)
    '        iv = ivct(0) & "=="
    '        ciphertext = If(ivct.Length = 3, ivct(1) & "==", ivct(1))

    '        AES.Key = SHA256.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(key))
    '        AES.IV = Convert.FromBase64String(iv)
    '        AES.Mode = Security.Cryptography.CipherMode.CBC
    '        Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
    '        Dim Buffer As Byte() = Convert.FromBase64String(ciphertext)
    '        plaintext = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
    '        Return plaintext
    '    Catch ex As Exception
    '        Return ex.Message
    '    End Try
    'End Function



    ''AES-ECB
    'Private Function AESE(ByVal input As Byte(), ByVal key As String) As Byte()
    '    Dim AES As New System.Security.Cryptography.RijndaelManaged
    '    Dim SHA256 As New System.Security.Cryptography.SHA256Cng
    '    Dim ciphertext As String = ""
    '    Try
    '        AES.Key = SHA256.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(key))
    '        AES.Mode = Security.Cryptography.CipherMode.ECB
    '        Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
    '        Dim Buffer As Byte() = input
    '        Return DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length)
    '    Catch ex As Exception
    '    End Try
    'End Function
    ''AES-ECB
    'Private Function AESD(ByVal input As Byte(), ByVal key As String) As Byte()
    '    Dim AES As New System.Security.Cryptography.RijndaelManaged
    '    Dim SHA256 As New System.Security.Cryptography.SHA256Cng
    '    Try
    '        AES.Key = SHA256.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(key))
    '        AES.Mode = Security.Cryptography.CipherMode.ECB
    '        Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
    '        Dim Buffer As Byte() = input
    '        Return DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length)
    '    Catch ex As Exception
    '    End Try
    'End Function

#End Region

#Region "DES"
    Public Function DESEncrypt(ByVal message As String, ByVal password As String) As String
        Dim messageBytes As Byte() = ASCIIEncoding.ASCII.GetBytes(message)
        Dim passwordBytes As Byte() = ASCIIEncoding.ASCII.GetBytes(password)
        Dim provider As DESCryptoServiceProvider = New DESCryptoServiceProvider()
        Dim transform As ICryptoTransform = provider.CreateEncryptor(passwordBytes, passwordBytes)
        Dim mode As CryptoStreamMode = CryptoStreamMode.Write
        Dim memStream As MemoryStream = New MemoryStream()
        Dim cryptoStream As CryptoStream = New CryptoStream(memStream, transform, mode)
        cryptoStream.Write(messageBytes, 0, messageBytes.Length)
        cryptoStream.FlushFinalBlock()
        Dim encryptedMessageBytes As Byte() = New Byte(memStream.Length - 1) {}
        memStream.Position = 0
        memStream.Read(encryptedMessageBytes, 0, encryptedMessageBytes.Length)
        Dim encryptedMessage As String = Convert.ToBase64String(encryptedMessageBytes)
        Return encryptedMessage
    End Function

    Public Function DESDecrypt(ByVal encryptedMessage As String, ByVal password As String) As String
        Dim encryptedMessageBytes As Byte() = Convert.FromBase64String(encryptedMessage)
        Dim passwordBytes As Byte() = ASCIIEncoding.ASCII.GetBytes(password)
        Dim provider As DESCryptoServiceProvider = New DESCryptoServiceProvider()
        Dim transform As ICryptoTransform = provider.CreateDecryptor(passwordBytes, passwordBytes)
        Dim mode As CryptoStreamMode = CryptoStreamMode.Write
        Dim memStream As MemoryStream = New MemoryStream()
        Dim cryptoStream As CryptoStream = New CryptoStream(memStream, transform, mode)
        cryptoStream.Write(encryptedMessageBytes, 0, encryptedMessageBytes.Length)
        cryptoStream.FlushFinalBlock()
        Dim decryptedMessageBytes As Byte() = New Byte(memStream.Length - 1) {}
        memStream.Position = 0
        memStream.Read(decryptedMessageBytes, 0, decryptedMessageBytes.Length)
        'Dim message As String = Convert.ToBase64String(decryptedMessageBytes)
        Dim message As String = ASCIIEncoding.ASCII.GetString(decryptedMessageBytes)
        Return message
    End Function

#End Region

End Module
