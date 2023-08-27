Imports Microsoft.VisualBasic
Imports System
Imports System.Security.Cryptography
Imports System.Text

Public Class fn_Cryptografia
    ' Encripta una cadena de texto usando el algoritmo de encriptacion de hash MD5.
    ' el "Message Digest" es una encriptacion de 128-bit y es usado comunmente para
    ' verificar datos chequeando el "Checksum MD5", mas informacion se puede
    ' encontrar en: http://www.faqs.org/rfcs/rfc1321.html

    ' cadena conteniendo el string a hashear a MD5.
    ' Una cadena de texto conteniendo en forma encriptada la cadena ingresada.
    Public Function MD5Hash(ByVal Data As String) As String
        Dim md5 As MD5 = New MD5CryptoServiceProvider()
        Dim hash As Byte() = md5.ComputeHash(Encoding.UTF8.GetBytes(Data))

        Dim stringBuilder As New StringBuilder()

        For Each b As Byte In hash
            stringBuilder.AppendFormat("{0:x2}", b)
        Next
        Return stringBuilder.ToString()
    End Function

    ' Encripta una cadena utilizando el algoritmo SHA256 (Secure Hash Algorithm)
    ' Detalles: http://www.itl.nist.gov/fipspubs/fip180-1.htm
    ' Esto trabaja de misma manera que el MD5, solo que utilizando una
    ' encriptacion en 256 bits.

    ' Un string conteniendo los datos a encriptar.
    ' Un string conteniendo al string de entrada, encriptado con el algoritmo SHA256.
    Public Shared Function SHA256Hash(ByVal Data As String) As String
        Dim sha As SHA256 = New SHA256Managed()
        Dim hash As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(Data))

        Dim stringBuilder As New StringBuilder()
        For Each b As Byte In hash
            stringBuilder.AppendFormat("{0:x2}", b)
        Next
        Return stringBuilder.ToString()
    End Function

    ' Encripta una cadena utilizando el algoritmo SHA256 (Secure Hash Algorithm)
    ' Detalles: http://www.itl.nist.gov/fipspubs/fip180-1.htm
    ' Esto trabaja de misma manera que el MD5, solo que utilizando una
    ' encriptacion en 256bits.

    ' Un string conteniendo los datos a encriptar.
    ' Un string conteniendo al string de entrada, encriptado con el algoritmo SHA384.
    Public Shared Function SHA384Hash(ByVal Data As String) As String
        Dim sha As SHA384 = New SHA384Managed()
        Dim hash As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(Data))

        Dim stringBuilder As New StringBuilder()
        For Each b As Byte In hash
            stringBuilder.AppendFormat("{0:x2}", b)
        Next
        Return stringBuilder.ToString()
    End Function

    ' Encripta una cadena utilizando el algoritmo SHA256 (Secure Hash Algorithm)
    ' Detalles: http://www.itl.nist.gov/fipspubs/fip180-1.htm
    ' Esto trabaja de misma manera que el MD5, solo que utilizando una
    ' encriptacion en 512 bits.

    ' Un string conteniendo los datos a encriptar.
    ' Un string conteniendo al string de entrada, encriptado con el algoritmo SHA512.
    Public Shared Function SHA512Hash(ByVal Data As String) As String
        Dim sha As SHA512 = New SHA512Managed()
        Dim hash As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(Data))

        Dim stringBuilder As New StringBuilder()
        For Each b As Byte In hash
            stringBuilder.AppendFormat("{0:x2}", b)
        Next
        Return stringBuilder.ToString()
    End Function

    ''' <summary>
    ''' Desencripta un texto.
    ''' </summary>
    ''' <param name="cryptedString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Decrypt(ByVal cryptedString As String, ByVal strKey As String) As String
        Dim bytes As Byte() = ASCIIEncoding.ASCII.GetBytes(strKey)
        If [String].IsNullOrEmpty(cryptedString) Then
            Throw New ArgumentNullException("The string which needs to be decrypted can not be null.")
        End If
        Dim cryptoProvider As New DESCryptoServiceProvider()
        Dim memoryStream As New IO.MemoryStream(Convert.FromBase64String(cryptedString))
        Dim cryptoStream As New CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read)
        Dim reader As New IO.StreamReader(cryptoStream)

        Return reader.ReadToEnd()
    End Function

    ''' <summary>
    ''' Encripta un texto dado
    ''' </summary>
    ''' <param name="originalString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Encrypt(ByVal originalString As String, ByVal strKey As String) As String
        Dim bytes As Byte() = ASCIIEncoding.ASCII.GetBytes(strKey)
        If [String].IsNullOrEmpty(originalString) Then
            Throw New ArgumentNullException("The string which needs to be encrypted can not be null.")
        End If
        Dim cryptoProvider As New DESCryptoServiceProvider()
        Dim memoryStream As New IO.MemoryStream()
        Dim cryptoStream As New CryptoStream(memoryStream, cryptoProvider.CreateEncryptor(bytes, bytes), CryptoStreamMode.Write)
        Dim writer As New IO.StreamWriter(cryptoStream)
        writer.Write(originalString)
        writer.Flush()
        cryptoStream.FlushFinalBlock()
        writer.Flush()

        Return Convert.ToBase64String(memoryStream.GetBuffer(), 0, CInt(memoryStream.Length))
    End Function
End Class