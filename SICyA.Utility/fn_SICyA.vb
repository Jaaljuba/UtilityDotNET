Imports Microsoft.VisualBasic
Imports System
Imports System.Configuration
Imports System.Configuration.ConfigurationManager
Imports System.DirectoryServices
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Net
Imports System.Net.Mail
Imports System.IO
Imports System.Exception
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Web.HttpServerUtility
Imports System.Globalization.CultureInfo

''' <summary>
''' Funciones desarrolladas por Javier A. Junca Barreto.
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Jaaljuba]	00/00/0000	Created
''' </history>
Public Class fn_SICyA
    Public sqlConnection As SqlConnection
    Public sqlCommand As SqlCommand

    ''' <summary>
    ''' Devuelve el ConnectionString del web.config según el nombre de dicha conexión parametrizada en una llave del mismo archvio.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function GetConnString() As String
        Dim NameConn As String
        Dim strConn As String

        NameConn = GetAppSettings("strNameConn")
        strConn = ConnectionStrings(NameConn).ConnectionString

        Return strConn
    End Function

    ''' <summary>
    ''' Devuelve el valor de una llave del web.config
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function GetAppSettings(ByVal strKey As String) As String
        Dim appSettings As System.Collections.Specialized.NameValueCollection
        Dim strValue As String

        Try
            appSettings = ConfigurationManager.AppSettings
            strValue = appSettings(strKey)
        Catch ex As Exception
            strValue = ""
        End Try

        Return strValue
    End Function

    Public Function GetUserSession(ByVal pUserMachine As String) As String
        Dim User As String
        Dim pInit As Integer
        Dim nCharsTotal As Integer
        Dim nCharsUser As Integer

        nCharsTotal = Len(pUserMachine)
        pInit = InStr(pUserMachine, "\")
        nCharsUser = nCharsTotal - pInit
        pInit += 1
        User = Mid(pUserMachine, pInit, nCharsUser)

        Return User
    End Function

    Private Function GetDirectoryEntry(Optional ByVal pUser As String = Nothing, Optional ByVal pPassword As String = Nothing) As DirectoryEntry
        Dim de As New DirectoryEntry()

        de.Path = GetAppSettings("PathLDAP")
        If (pUser <> Nothing) Then
            de.Username = pUser
            If (pPassword <> Nothing) Then
                de.Password = pPassword
            End If
        End If
        de.AuthenticationType = AuthenticationTypes.Secure

        Return de
    End Function

    Public Function getUserNameLDAP(ByVal pIdUser As String) As String
        Dim entry As DirectoryEntry = GetDirectoryEntry()

        Try
            Dim search As DirectorySearcher = New DirectorySearcher(entry)

            search.Filter = "(SAMAccountName=" + pIdUser + ")"
            search.PropertiesToLoad.Add("displayName")

            Dim result As SearchResult = search.FindOne()

            If (Not IsNothing(result)) Then
                Return result.Properties("displayname")(0).ToString()
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getUserNameV2(ByVal user As String, ByVal pwd As String) As String
        If user = "" Then Return ""

        Dim entry As DirectoryEntry = GetDirectoryEntry(user, pwd)

        Try
            Dim var As String = "displayname"
            Dim search As DirectorySearcher = New DirectorySearcher(entry)
            search.Filter = "(SAMAccountName=" + user.Trim + ")"
            search.PropertiesToLoad.Add(var)

            Dim result As SearchResult = search.FindOne()
            If (Not IsNothing(result)) Then
                Return result.Properties(var)(0).ToString()
            Else
                Return ""
            End If
        Catch ex As Exception
            SaveErrorLog(ex, False)

            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Devuelve un DataSet con la informacion cargada desde el archivo Excel.
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function LoadFileExcel(ByVal FileExcel As String, ByVal Version As String, ByVal strQuery As String) As DataSet
        Dim ConnOleDB As OleDbConnection
        Dim CmdOleDB As New OleDbCommand
        Dim Consulta As New OleDbDataAdapter
        Dim dsInfo As New DataSet

        Dim strConnExcel As String = ""

        If Version = "2003" Then
            strConnExcel = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & FileExcel & ";" & _
                           "Extended Properties=""Excel 5.0;HDR=Yes;IMEX=1"";"
        ElseIf Version = "2007" Then
            strConnExcel = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                           "Data Source=" & FileExcel & ";" & _
                           "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=2"";"
        End If


        ConnOleDB = New OleDbConnection(strConnExcel)

        ConnOleDB.Open()
        CmdOleDB.CommandText = strQuery
        CmdOleDB.CommandType = CommandType.Text
        CmdOleDB.Connection = ConnOleDB
        Consulta.SelectCommand = CmdOleDB
        Consulta.Fill(dsInfo, "Excel")
        ConnOleDB.Close()

        Return dsInfo
    End Function

    ''' <summary>
    ''' Procedimiento que graba el contenido de un DataTable en un archivo de tipo Excel.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Sub SaveDataTableExcel(ByVal FileName As String, ByVal dtData As DataTable)
        Web.HttpContext.Current.Response.Clear()
        Web.HttpContext.Current.Response.AddHeader("Content-Disposition", String.Format("attachment; filename={0}", FileName))
        Web.HttpContext.Current.Response.ContentType = "application/vnd.ms-excel"

        Dim sb As StringBuilder = New StringBuilder()
        Dim sw As StringWriter = New System.IO.StringWriter(sb)
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        Dim Page As Page = New Page()
        Dim form As HtmlControls.HtmlForm = New HtmlControls.HtmlForm()

        Dim Fila As DataRow
        Dim Arreglo As Object
        Dim i As Integer
        Dim iColumna As Integer

        '--- Se generan las columnas del archivo excel ---'
        For iColumna = 0 To dtData.Columns.Count - 1
            Web.HttpContext.Current.Response.Write(dtData.Columns(iColumna).ToString & vbTab)
        Next
        Web.HttpContext.Current.Response.Write(vbCrLf)

        '--- Se generan las filas del archivo excel ---'
        For Each Fila In dtData.Rows
            Arreglo = Fila.ItemArray
            For i = 0 To UBound(Arreglo)
                Web.HttpContext.Current.Response.Write(Arreglo(i).ToString & vbTab)
            Next
            Web.HttpContext.Current.Response.Write(vbCrLf)
        Next

        Web.HttpContext.Current.Response.Buffer = True
        Web.HttpContext.Current.Response.Charset = "UTF-8"
        Web.HttpContext.Current.Response.ContentEncoding = Encoding.UTF8
        Web.HttpContext.Current.Response.Write(sw.ToString)
        Web.HttpContext.Current.Response.End()
    End Sub

    ''' <summary>
    ''' Procedimiento que graba las excepciones o los errores ocurridos en un archivo Log y envia la notificacion por correo (Opcional).
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Sub SaveErrorLog(ByVal pEx As Exception, Optional ByVal pSendMail As Boolean = False)
        Dim FileLog As StreamWriter
        Dim strError As String = ""
        Dim FileNameLog As String

        Dim NameFrom As String
        Dim MailFrom As String
        Dim MailTo As String
        Dim Subject As String
        Dim User As String
        Dim Password As String
        Dim Port As Short
        Dim Host As String
        Dim ActiveSSL As Boolean

        FileNameLog = GetAppSettings("FilePathErrors")
        If FileNameLog <> "" Then
            Try
                FileLog = New StreamWriter(FileNameLog, True)
                strError = "Fecha y Hora: " & DateTime.Now.ToString & vbCrLf & _
                           "Usuario: " & My.User.Name.ToString & vbCrLf & _
                           "Message: " & pEx.Message.ToString & vbCrLf & _
                           "Source: " & pEx.Source.ToString & vbCrLf & _
                           "Trace: " & pEx.StackTrace.ToString & vbCrLf & vbCrLf
                FileLog.Write(strError)
                FileLog.Close()
            Catch

            End Try
        End If

        If pSendMail = True Then
            NameFrom = GetAppSettings("sendeMail_From")
            MailFrom = GetAppSettings("sendeMail_MailFrom")
            MailTo = GetAppSettings("sendeMail_MailTo")
            Subject = "Error en Aplicación [eViajes]"
            User = GetAppSettings("sendeMail_User")
            Password = GetAppSettings("sendeMail_Password")
            Port = CType(GetAppSettings("sendeMail_Port"), Short)
            Host = GetAppSettings("sendeMail_Host")
            ActiveSSL = IIf(GetAppSettings("sendeMail_SSL") = "T", True, False)

            senderMail(NameFrom, MailFrom, MailTo, Subject, strError, User, Password, Port, Host, ActiveSSL)
        End If
    End Sub

    ''' <summary>
    ''' Procedimiento que graba las excepciones o los errores ocurridos en un archivo Log y envia la notificacion por correo (Opcional).
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Sub SaveErrorLog(ByVal pEx As String, Optional ByVal pSendMail As Boolean = False)
        Dim FileLog As StreamWriter
        Dim strError As String = ""
        Dim FileNameLog As String

        Dim NameFrom As String
        Dim MailFrom As String
        Dim MailTo As String
        Dim Subject As String
        Dim User As String
        Dim Password As String
        Dim Port As Short
        Dim Host As String
        Dim ActiveSSL As Boolean

        FileNameLog = GetAppSettings("FilePathErrors")
        If FileNameLog <> "" Then
            Try
                FileLog = New StreamWriter(FileNameLog, True)
                strError = "Fecha y Hora: " & DateTime.Now.ToString & vbCrLf & _
                           "Usuario: " & My.User.Name.ToString & vbCrLf & _
                           "Message: " & pEx.ToString
                FileLog.Write(strError)
                FileLog.Close()
            Catch
            End Try
        End If
        If pSendMail = True Then
            NameFrom = GetAppSettings("sendeMail_From")
            MailFrom = GetAppSettings("sendeMail_MailFrom")
            MailTo = GetAppSettings("sendeMail_MailTo")
            Subject = "Error en Aplicación [eViajes]"
            User = GetAppSettings("sendeMail_User")
            Password = GetAppSettings("sendeMail_Password")
            Port = CType(GetAppSettings("sendeMail_Port"), Short)
            Host = GetAppSettings("sendeMail_Host")
            ActiveSSL = IIf(GetAppSettings("sendeMail_SSL") = "T", True, False)

            senderMail(NameFrom, MailFrom, MailTo, Subject, strError, User, Password, Port, Host, ActiveSSL)
        End If
    End Sub

    ''' <summary>
    ''' Funcion que envia Mail y devuelve el estado del envio.
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function senderMail(ByVal strNameFrom As String, ByVal strFrom As String, ByVal strTo As String, ByVal strSubject As String, ByVal strBody As String, ByVal strUser As String, ByVal strPass As String, ByVal strPort As Short, ByVal strHost As String, Optional ByVal SSL As Boolean = False, Optional ByVal HTML As Boolean = False, Optional ByVal strToCC As String = Nothing, Optional ByVal strAttachment As String = Nothing) As Boolean
        Dim clientMail As New SmtpClient
        Dim msgMail As New MailMessage
        Dim attachmentMail As Attachment
        Dim ErrorFound As Boolean

        msgMail.To.Add(strTo)
        If strToCC IsNot Nothing Then
            msgMail.CC.Add(strToCC)
        End If
        msgMail.From = New MailAddress(strFrom, strNameFrom, Encoding.UTF8)
        msgMail.Subject = strSubject
        msgMail.SubjectEncoding = Encoding.UTF8
        msgMail.Body = strBody
        msgMail.BodyEncoding = Encoding.UTF8
        msgMail.IsBodyHtml = HTML
        msgMail.Priority = MailPriority.Normal

        If strAttachment IsNot Nothing Then
            attachmentMail = New Attachment(strAttachment)
            msgMail.Attachments.Add(attachmentMail)
        End If

        If SSL = True Then
            clientMail.Credentials = New NetworkCredential(strUser, strPass)
        End If

        clientMail.Port = strPort
        clientMail.Host = strHost
        clientMail.EnableSsl = SSL

        Try
            clientMail.Send(msgMail)
            ErrorFound = False
        Catch ex As Exception
            SaveErrorLog(ex, False)
            ErrorFound = True
        End Try

        Return ErrorFound
    End Function

    ''' <summary>
    ''' Devuelve el ConnectionString del web.config según el nombre de dicha conexión parametrizada en una llave del mismo archvio.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function CurrencySeparator(ByVal Tipo As String) As String
        Dim Separator As String = ""

        Select Case UCase(Tipo)
            Case "M"
                Separator = CurrentCulture.NumberFormat.CurrencyGroupSeparator.ToString
            Case "D"
                Separator = CurrentCulture.NumberFormat.CurrencyDecimalSeparator.ToString
        End Select

        Return Separator
    End Function

    Public Sub CreateXML(ByVal NameXML As String, ByVal NameNode As String)
        Dim xmlDoc As New XmlDocument()
        Dim xmlDeclaracion As XmlDeclaration
        Dim General As XmlNode

        xmlDeclaracion = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
        General = xmlDoc.CreateElement(NameNode)
        xmlDoc.AppendChild(xmlDeclaracion)
        xmlDoc.AppendChild(General)

        Try
            xmlDoc.Save(NameXML)
        Catch ex As Exception
            SaveErrorLog(ex)
        End Try
    End Sub

    Public Sub AddNodeAfterXML(ByVal FileNameXML As String, ByVal Pointer As Integer, ByVal ItemLevel As String, ByVal ItemField As String, ByVal ItemFieldNew As String)
        Dim xmlDoc As New XmlDocument()
        Dim Node As XmlNode
        Dim BeforeNode As XmlElement

        If File.Exists(FileNameXML) Then
            Try
                xmlDoc.Load(FileNameXML)
                Node = xmlDoc.DocumentElement.ChildNodes.Item(Pointer)

                If ItemField = Nothing Then
                    BeforeNode = Node(ItemLevel)
                    Node.InsertAfter(xmlDoc.CreateElement(ItemFieldNew), BeforeNode)
                Else
                    BeforeNode = Node(ItemLevel)(ItemField)
                    Node(ItemLevel).InsertAfter(xmlDoc.CreateElement(ItemFieldNew), BeforeNode)
                End If

                xmlDoc.Save(FileNameXML)
            Catch ex As Exception
                SaveErrorLog(ex)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Devuelve el ConnectionString del web.config según el nombre de dicha conexión parametrizada en una llave del mismo archvio.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function ParameterValueXML(ByVal Pointer As Integer, ByVal FileNameXML As String, ByVal ItemLevel As String, ByVal ItemField As String) As String
        Dim xmlDoc As New XmlDocument
        Dim Node As XmlNode
        Dim strReturn As String = String.Empty

        If File.Exists(FileNameXML) Then
            Try
                xmlDoc.Load(FileNameXML)

                Node = xmlDoc.DocumentElement.ChildNodes.Item(Pointer)
                If ItemField = Nothing Then
                    strReturn = Node(ItemLevel).InnerText.ToString
                Else
                    strReturn = Node(ItemLevel)(ItemField).InnerText.ToString
                End If
            Catch ex As Exception
                SaveErrorLog(ex)
            End Try
        End If

        Return strReturn
    End Function

    ''' <summary>
    ''' Devuelve el ConnectionString del web.config según el nombre de dicha conexión parametrizada en una llave del mismo archvio.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function EncodedASCII(ByVal strOriginal As String) As String
        Dim dstEncoding As Encoding
        Dim encodedBytes As Byte()
        Dim decodedString As String

        dstEncoding = System.Text.Encoding.ASCII
        encodedBytes = dstEncoding.GetBytes(strOriginal)
        decodedString = dstEncoding.GetString(encodedBytes)

        Return decodedString
    End Function

    ''' <summary>
    ''' Devuelve el ConnectionString del web.config según el nombre de dicha conexión parametrizada en una llave del mismo archvio.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function ReplaceAcentos(ByVal strOriginal As String) As String
        Dim strNew As String

        strNew = strOriginal
        strNew = Replace(strNew, "á", "a")
        strNew = Replace(strNew, "é", "e")
        strNew = Replace(strNew, "í", "i")
        strNew = Replace(strNew, "ó", "o")
        strNew = Replace(strNew, "ú", "u")
        strNew = Replace(strNew, "ñ", "n")
        strNew = Replace(strNew, "Ñ", "N")
        strNew = Replace(strNew, vbCrLf, " ")

        Return strNew
    End Function

    ''' <summary>
    ''' Devuelve el ConnectionString del web.config según el nombre de dicha conexión parametrizada en una llave del mismo archvio.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' [Jaaljuba]	00/00/0000	Created
    ''' [Jaaljuba]  14/11/2013  Modified
    ''' [Jaaljuba]  09/09/2015  Modified
    '''   
    ''' </history>
    Public Function GUID(Optional ByVal pPrefix As String = Nothing) As String
        Dim NumberGenerator As Random = New Random()
        Dim strID As String = Nothing
        Dim Residuo As Short

        If pPrefix = Nothing Then
            strID = "AUT"
        Else
            strID = pPrefix
        End If

        'For X As Short = Len(pPrefix) To 3
        '    strID += Chr(NumberGenerator.Next(65, 90))
        'Next

        For X = 4 To 20
            Residuo = X Mod 3

            Select Case Residuo
                Case 0
                    '-- Genera letras minusculas.
                    strID += Chr(NumberGenerator.Next(97, 122))
                Case 1
                    '-- Genera numeros.
                    strID += Chr(NumberGenerator.Next(48, 57))
                Case 2
                    '-- Genera letras mayusculas.
                    strID += Chr(NumberGenerator.Next(65, 90))
            End Select
        Next

        Return strID
    End Function

    ''' <summary>
    ''' Devuelve el ConnectionString del web.config según el nombre de dicha conexión parametrizada en una llave del mismo archvio.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Jaaljuba]	00/00/0000	Created
    ''' </history>
    Public Function LastDayMonth(ByVal pMonth As String, ByVal pYear As String) As String
        Dim NextMonth As Date
        Dim LastDay As String

        NextMonth = DateAdd(DateInterval.Month, 1, DateSerial(pYear, pMonth, "01"))
        LastDay = DateAdd(DateInterval.Day, -1, NextMonth)

        Return LastDay.ToString
    End Function

    ''' <sumary>
    ''' Reemplaza los parametros de un template y los devuelve en una cadana de caracteres.
    ''' </sumary>
    ''' <param name="pTemplateName"></param>
    ''' <param name="pParameters"></param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Jaaljuba] 13/01/2014 - Created
    ''' [Jaaljuba] 17/01/2013 - Se cambia el nombre de las variables 
    ''' </history>
    Public Function ReplaceTemplateString(ByVal pTemplateName As String, ByVal ParamArray pParameters() As String) As String
        Dim OldText As StreamReader
        Dim NewText As New StringBuilder
        Dim i As Integer
        Dim rText As String

        If pTemplateName IsNot Nothing Then
            If File.Exists(pTemplateName) Then
                OldText = New StreamReader(pTemplateName, Text.Encoding.Default)
                NewText.Append(OldText.ReadToEnd)
                OldText.Close()
            End If
            For i = 0 To pParameters.Length - 1 Step 2
                'If IsNothing(pParameters(i)) OrElse IsNothing(pParameters(i + 1) OrElse i + 1 >= pParameters.Length) Then Continue For
                If IsNothing(pParameters(i)) OrElse IsNothing(pParameters(i + 1)) OrElse i + 1 >= pParameters.Length Then Continue For
                NewText.Replace(pParameters(i), pParameters(i + 1))
            Next
        End If
        rText = NewText.ToString

        Return rText
    End Function

    ''' <summary>
    ''' Remplaza las palabras y las devuelve con un class de CSS de resaltado.
    ''' </summary>
    ''' <param name="pStrMatch"></param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Jaaljuba] 27/01/2014 - Created
    ''' </history>
    Public Function ReplaceWords(ByVal pStrMatch As Match) As String
        Return "<span class='highlight'>" + pStrMatch.ToString + "</span>"
    End Function

    ' ''' <summary>
    ' ''' Carga y devuelve una tabla a partir de la ejecucion de un store procedure.
    ' ''' El Store Procedure NO debe requerir parametros
    ' ''' </summary>
    ' ''' <param name="pStoreProcedure"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    ' ''' <history>
    ' ''' [Jaaljuba] 13/03/2015 - Created
    ' ''' </history>
    'Public Function CargaTablaStoreProcedure(ByVal pStoreProcedure As String) As DataTable
    '    Dim daData As New SqlDataAdapter
    '    Dim dtData As New DataTable
    '    Dim Parameter As New SqlParameter

    '    sqlConnection = New SqlConnection
    '    sqlCommand = New SqlCommand

    '    sqlConnection.ConnectionString = GetConnString()
    '    Try
    '        sqlConnection.Open()

    '        sqlCommand.CommandText = pStoreProcedure
    '        sqlCommand.CommandType = CommandType.StoredProcedure
    '        sqlCommand.Connection = sqlConnection

    '        daData.SelectCommand = sqlCommand
    '        daData.Fill(dtData)
    '    Catch ex As Exception
    '        SaveErrorLog(ex)
    '        dtData = Nothing
    '    End Try

    '    Return dtData
    'End Function

    ''' <summary>
    ''' Carga y devuelve una tabla a partir de la ejecucion de un store procedure.
    ''' El store procedure recibe parametros en un ParamArray
    ''' </summary>
    ''' <param name="pStoreProcedure"></param>
    ''' <param name="pParameters"></param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Jaaljuba] 13/03/2015 - Created
    ''' </history>
    Public Function LoadTableFromStoreProcedure(ByVal pStoreProcedure As String, ByVal ParamArray pParameters() As Object) As DataTable
        Dim daData As New SqlDataAdapter
        Dim dtData As New DataTable

        sqlConnection = New SqlConnection
        sqlCommand = New SqlCommand

        sqlConnection.ConnectionString = GetConnString()
        Try
            sqlConnection.Open()

            sqlCommand.CommandText = pStoreProcedure
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Connection = sqlConnection

            sqlCommand.Parameters.Clear()
            If IsArray(pParameters) Then
                sqlCommand.Parameters.AddRange(pParameters)
            End If

            daData.SelectCommand = sqlCommand
            daData.Fill(dtData)
        Catch ex As Exception
            SaveErrorLog(ex)
            dtData = Nothing
        Finally
            If sqlConnection.State = ConnectionState.Open Then
                sqlConnection.Close()
            End If
        End Try

        Return dtData
    End Function

    ''' <summary>
    ''' Carga y devuelve una tabla a partir de la ejecucion de un store procedure.
    ''' El store procedure recibe parametros en un List(Of SqlParameter)
    ''' </summary>
    ''' <param name="pStoreProcedure"></param>
    ''' <param name="pParameters"></param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Jaaljuba] 13/03/2015 - Created
    ''' </history>
    Public Function LoadTableFromStoreProcedure(ByVal pStoreProcedure As String, ByVal pParameters As List(Of SqlParameter)) As DataTable
        Dim daData As New SqlDataAdapter
        Dim dtData As New DataTable
        Dim Parameter As New SqlParameter

        sqlConnection = New SqlConnection
        sqlCommand = New SqlCommand

        sqlConnection.ConnectionString = GetConnString()
        Try
            sqlConnection.Open()

            sqlCommand.CommandText = pStoreProcedure
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Connection = sqlConnection

            sqlCommand.Parameters.Clear()
            For Each Parameter In pParameters
                sqlCommand.Parameters.Add(Parameter)
            Next

            daData.SelectCommand = sqlCommand
            daData.Fill(dtData)
        Catch ex As Exception
            SaveErrorLog(ex)
            dtData = Nothing
        Finally
            If sqlConnection.State = ConnectionState.Open Then
                sqlConnection.Close()
            End If
        End Try

        Return dtData
    End Function

    Public Function LoadGenericListFromStoreProcedure(ByVal pStoreProcedure As String, Optional ByVal pParameters As List(Of SqlParameter) = Nothing) As List(Of GenericList)
        Dim drDatos As SqlDataReader
        Dim Records As New List(Of GenericList)
        Dim Record As GenericList
        Dim _id, _value As String

        sqlConnection = New SqlConnection
        sqlCommand = New SqlCommand

        sqlConnection.ConnectionString = GetConnString()
        Try
            sqlConnection.Open()

            sqlCommand.CommandText = pStoreProcedure
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Connection = sqlConnection

            sqlCommand.Parameters.Clear()
            For Each Parameter In pParameters
                sqlCommand.Parameters.Add(Parameter)
            Next

            drDatos = sqlCommand.ExecuteReader()
            If drDatos.HasRows = True Then
                While drDatos.Read
                    _id = drDatos(0)
                    _value = drDatos(1).ToString
                    Record = New GenericList(_id, _value)
                    Records.Add(Record)
                End While
            End If
        Catch ex As Exception
            SaveErrorLog(ex)
        Finally
            If sqlConnection.State = ConnectionState.Open Then
                sqlConnection.Close()
            End If
        End Try

        Return Records
    End Function

End Class
