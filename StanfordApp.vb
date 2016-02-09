Imports system
Imports system.web
Imports system.web.ui
Imports system.web.ui.webcontrols
Imports system.web.ui.htmlcontrols
Imports system.configuration
Imports system.data
Imports system.data.sqlclient
Imports system.drawing
Imports system.collections.specialized
Imports system.text.regularexpressions
Imports system.collections
Imports system.io
Imports system.net.mail
Imports powerviewPDF
Imports FlatbridgeEncryption
Imports System.Net
Imports System.Security.Cryptography


Public Class ListData
    Private idString As String
    Private valString As String
    Public Sub New(ByVal iId As String, ByVal sName As String)
        idString = iId
        valString = sName
    End Sub
    Public Property Id() As String
        Get
            Return idString
        End Get
        Set(ByVal Value As String)
            idString = Value
        End Set
    End Property
    Public Property Name() As String
        Get
            Return valString
        End Get
        Set(ByVal Value As String)
            valString = Value
        End Set
    End Property
End Class



Public Class StanfordApp
    Inherits System.web.ui.page

    Dim TestMode As Integer = 0
    Dim ItemsPerPage As Integer = 25
    Dim MyConnection As System.Data.SqlClient.SqlConnection, MyConnection2 As System.Data.SqlClient.SqlConnection, MyConnection3 As System.Data.SqlClient.SqlConnection, MyConnection4 As System.Data.SqlClient.SqlConnection


    Function getMerchantID() As String
        Dim sql, rs
        Dim retVal = ""
        sql = "select b.merchantId from course a, application b where a.application_id=b.id and a.id=" & CInt(Request.QueryString("id"))
        rs = BQ(sql)
        If rs.Read() Then
            retVal = rs("merchantId")
        End If
        EQ(rs)
        Return retVal
    End Function

    Function getPublicKey() As String
        Dim sql, rs
        Dim retVal = ""
        sql = "select b.publicKey from course a, application b where a.application_id=b.id and a.id=" & CInt(Request.QueryString("id"))
        rs = BQ(sql)
        If rs.Read() Then
            retVal = rs("publicKey")
        End If
        EQ(rs)
        Return retVal
    End Function

    Function getPrivateKey() As String
        Dim sql, rs
        Dim retVal = ""
        sql = "select b.privateKey from course a, application b where a.application_id=b.id and a.id=" & CInt(Request.QueryString("id"))
        rs = BQ(sql)
        If rs.Read() Then
            retVal = rs("privateKey")
        End If
        EQ(rs)
        Return retVal
    End Function

    Function getSerialNumber() As String
        Dim sql, rs
        Dim retVal = ""
        sql = "select b.serialNumber from course a, application b where a.application_id=b.id and a.id=" & CInt(Request.QueryString("id"))
        rs = BQ(sql)
        If rs.Read() Then
            retVal = rs("serialNumber")
        End If
        EQ(rs)
        Return retVal
    End Function

    'function deprecated
    Function insertSignature(ByVal amount As String, ByVal currency As String) As String
        insertSignature3(amount, currency, "authorization")
        Return "1"
    End Function

    Sub LogError(msg As String)
        Dim sql, rs
        sql = "insert into System_Errors values('" & db(msg) & "',getDate(),'StanfordApp')"
        rs = BQ(sql)
        EQ(rs)

    End Sub

    Function insertSignature4(ByVal amount As String, ByVal currency As String, ByVal orderPage_transactionType As String) As String
        Try
            Dim timeSpanTime As TimeSpan = DateTime.UtcNow - New DateTime(1970, 1, 1)
            Dim arrayTime() As String = timeSpanTime.TotalMilliseconds.ToString().Split(".")
            Dim time As String = arrayTime(0)
            Dim merchantID As String = getMerchantID()
            If merchantID.Equals("") Then
                Response.Write("<b>Error:</b> <br>The current security script (HOP.vb) doesn't contain your merchant information. Please login to the <a href='https://ebc.cybersource.com/ebc/hop/HOPSecurityLoad.do'>CyberSource Business Center</a> and generate one before proceeding further. Be sure to replace the existing HOP.vb with the newly generated HOP.vb.<br><br>")
            End If

            Dim data As String = merchantID + amount + currency + time + orderPage_transactionType
            Dim pub As String = getPublicKey()
            Dim serialNumber As String = getSerialNumber()
            Dim byteData() As Byte = System.Text.Encoding.UTF8.GetBytes(data)
            Dim byteKey() As Byte = System.Text.Encoding.UTF8.GetBytes(pub)
            Dim hmac As HMACSHA1 = New HMACSHA1(byteKey)
            Dim publicDigest As String = Convert.ToBase64String(hmac.ComputeHash(byteData))
            publicDigest = publicDigest.Replace("\n", "")
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
            sb.Append("<input type='hidden' name='amount' value='")
            sb.Append(amount)
            sb.Append("'><input type='hidden' name='currency' value='")
            sb.Append(currency)
            sb.Append("'><input type='hidden' name='orderPage_timestamp' value='")
            sb.Append(time)
            sb.Append("'><input type='hidden' name='merchantID' value='")
            sb.Append(merchantID)
            sb.Append("'><input type='hidden' name='orderPage_transactionType' value='")
            sb.Append(orderPage_transactionType)
            sb.Append("'><input type='hidden' name='orderPage_signaturePublic' value='")
            sb.Append(publicDigest)
            sb.Append("'><input type='hidden' name='orderPage_version' value='4'>")
            sb.Append("<input type='hidden' name='orderPage_serialNumber' value='")
            sb.Append(serialNumber)
            sb.Append("'>")
            Response.Write(sb.ToString())
        Catch e As Exception
            ' Response.Write(e.StackTrace.ToString())
        End Try
        Return "1"
    End Function


    Function insertSignature3(ByVal amount As String, ByVal currency As String, ByVal orderPage_transactionType As String) As String
        Try
            Dim timeSpanTime As TimeSpan = DateTime.UtcNow - New DateTime(1970, 1, 1)
            Dim arrayTime() As String = timeSpanTime.TotalMilliseconds.ToString().Split(".")
            Dim time As String = arrayTime(0)
            Dim merchantID As String = getMerchantID()
            If merchantID.Equals("") Then
                Response.Write("<b>Error:</b> <br>The current security script (HOP.vb) doesn't contain your merchant information. Please login to the <a href='https://ebc.cybersource.com/ebc/hop/HOPSecurityLoad.do'>CyberSource Business Center</a> and generate one before proceeding further. Be sure to replace the existing HOP.vb with the newly generated HOP.vb.<br><br>")
            End If

            Dim data As String = merchantID + amount + currency + time + orderPage_transactionType
            Dim pub As String = getPublicKey()
            Dim serialNumber As String = getSerialNumber()
            Dim byteData() As Byte = System.Text.Encoding.UTF8.GetBytes(data)
            Dim byteKey() As Byte = System.Text.Encoding.UTF8.GetBytes(pub)
            Dim hmac As HMACSHA1 = New HMACSHA1(byteKey)
            Dim publicDigest As String = Convert.ToBase64String(hmac.ComputeHash(byteData))
            publicDigest = publicDigest.Replace("\n", "")
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
            sb.Append("<input type='hidden' name='amount' value='")
            sb.Append(amount)
            sb.Append("'><input type='hidden' name='currency' value='")
            sb.Append(currency)
            sb.Append("'><input type='hidden' name='orderPage_timestamp' value='")
            sb.Append(time)
            sb.Append("'><input type='hidden' name='merchantID' value='")
            sb.Append(merchantID)
            sb.Append("'><input type='hidden' name='orderPage_transactionType' value='")
            sb.Append(orderPage_transactionType)
            sb.Append("'><input type='hidden' name='orderPage_signaturePublic' value='")
            sb.Append(publicDigest)
            sb.Append("'><input type='hidden' name='orderPage_version' value='4'>")
            sb.Append("<input type='hidden' name='orderPage_serialNumber' value='")
            sb.Append(serialNumber)
            sb.Append("'>")
            Response.Write(sb.ToString())
        Catch e As Exception
            ' Response.Write(e.StackTrace.ToString())
        End Try
        Return "1"
    End Function

    Function insertSubscriptionSignature(ByVal subscriptionAmount As String, ByVal subscriptionStartDate As String, ByVal subscriptionFrequency As String, ByVal subscriptionNumberOfPayments As String, ByVal subscriptionAutomaticRenew As String) As String
        Try
            Dim data As String = subscriptionAmount + subscriptionStartDate + subscriptionFrequency + subscriptionNumberOfPayments + subscriptionAutomaticRenew
            Dim pub As String = getPublicKey()
            Dim byteData() As Byte = System.Text.Encoding.UTF8.GetBytes(data)
            Dim byteKey() As Byte = System.Text.Encoding.UTF8.GetBytes(pub)
            Dim HMAC As HMACSHA1 = New HMACSHA1(byteKey)
            Dim publicDigest As String = Convert.ToBase64String(HMAC.ComputeHash(byteData))
            publicDigest = publicDigest.Replace("\n", "")
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
            sb.Append("<input type='hidden' name='recurringSubscriptionInfo_amount' value='")
            sb.Append(subscriptionAmount)
            sb.Append("'><input type='hidden' name='recurringSubscriptionInfo_numberOfPayments' value='")
            sb.Append(subscriptionNumberOfPayments)
            sb.Append("'><input type='hidden' name='recurringSubscriptionInfo_frequency' value='")
            sb.Append(subscriptionFrequency)
            sb.Append("'><input type='hidden' name='recurringSubscriptionInfo_automaticRenew' value='")
            sb.Append(subscriptionAutomaticRenew)
            sb.Append("'><input type='hidden' name='recurringSubscriptionInfo_startDate' value='")
            sb.Append(subscriptionStartDate)
            sb.Append("'><input type='hidden' name='recurringSubscriptionInfo_signaturePublic' value='")
            sb.Append(publicDigest)
            sb.Append("'>")
            Response.Write(sb.ToString())
        Catch e As Exception
            ' Response.Write(e.StackTrace.ToString())
        End Try
        Return "1"
    End Function

    Function insertSubscriptionIDSignature(ByVal subscriptionID As String) As String
        Try
            Dim pub As String = getPublicKey()
            Dim data As String = subscriptionID
            Dim byteData() As Byte = System.Text.Encoding.UTF8.GetBytes(data)
            Dim byteKey() As Byte = System.Text.Encoding.UTF8.GetBytes(pub)
            Dim HMAC As HMACSHA1 = New HMACSHA1(byteKey)
            Dim publicDigest As String = Convert.ToBase64String(HMAC.ComputeHash(byteData))
            publicDigest = publicDigest.Replace("\n", "")
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
            sb.Append("<input type='hidden' name='paySubscriptionCreateReply_subscriptionID' value='")
            sb.Append(subscriptionID)
            sb.Append("'><input type='hidden' name='paySubscriptionCreateReply_subscriptionIDPublicSignature' value='")
            sb.Append(publicDigest)
            sb.Append("'>")
            Response.Write(sb.ToString())
        Catch e As Exception
            ' Response.Write(e.StackTrace.ToString())
        End Try
        Return "1"
    End Function

    Function verifyTransactionSignature(ByVal map As System.Web.HttpRequest) As Boolean
        Dim transactionSignature As String
        Dim signedFieldsArr As String()
        Try
            transactionSignature = map.Form.Get("transactionSignature")
            signedFieldsArr = map.Form.Get("signedFields").Split(",")
        Catch e As Exception
            Return False
        End Try
        Dim data As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim i As Integer = 0
        While i < signedFieldsArr.Length
            data.Append(map.Form.Get(signedFieldsArr(i)))
            i = i + 1
        End While
        Dim byteData As Byte() = System.Text.Encoding.UTF8.GetBytes(data.ToString())
        Dim byteKey As Byte() = System.Text.Encoding.UTF8.GetBytes(getPublicKey())
        Dim HMAC As HMACSHA1 = New HMACSHA1(byteKey)
        Dim publicDigest As String = Convert.ToBase64String(HMAC.ComputeHash(byteData))
        If (transactionSignature.Equals(publicDigest)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Sub AppCapture()
        Try
            Dim sql, rs
            Dim body

            Dim loop1, loop2 As Integer
            Dim arr1(), arr2() As String
            Dim coll As NameValueCollection

            ' Load Form variables into NameValueCollection variable.
            coll = Request.Form

            ' Get names of all forms into a string array.
            arr1 = coll.AllKeys
            For loop1 = 0 To arr1.GetUpperBound(0)
                If arr1(loop1) <> "Submit" Then
                    body = body & arr1(loop1) & ":" & Request.Form(arr1(loop1)) & vbCrLf
                End If
            Next loop1

            coll = Request.ServerVariables
            ' Get names of all keys into a string array.
            arr1 = coll.AllKeys
            For loop1 = 0 To arr1.GetUpperBound(0)
                body = body & vbCrLf & "Key: " & arr1(loop1) & vbCrLf
                arr2 = coll.GetValues(loop1) ' Get all values under this key.
                For loop2 = 0 To arr2.GetUpperBound(0)
                    body = body & "Value " & CStr(loop2) & ": " & arr2(loop2) & vbCrLf
                Next loop2
            Next loop1

            sql = "insert into Online_App_Submit values('" & db(body) & "',getDate())"
            rs = BQ(sql)
            EQ(rs)
        Catch ex As Exception

        End Try

    End Sub



    Sub WriteGeneralManagementFooter()
        Response.Write("<table  class=""cb"">")
        Response.Write("  <tr>")
        Response.Write("	<td>")
        Response.Write("<a href=""http://www.gsb.stanford.edu/policy/terms.html"">TERMS OF USE</a><b> :: </b><a href=""http://www.gsb.stanford.edu/policy/privacy.html"">PRIVACY POLICY</a><strong> :: </strong><a href=""http://www.gsb.stanford.edu/help/"">HELP</a></td>")
        Response.Write("	<td align=""right""> &copy; " & Now().Year & " Stanford University Graduate School of Business</td>")
        Response.Write("  </tr>")
        Response.Write("</table>")
    End Sub



    Sub RenderExcel(gView As GridView, file As String)
        Dim htmlWrite As HtmlTextWriter
        Dim stringWrite As System.IO.StringWriter
        gView.allowPaging = False
        Response.Clear()
        Response.AddHeader("content-disposition", "attachment;filename=" & file)
        Response.Charset = ""
        Response.ContentType = "application/vnd.xls"
        stringWrite = New System.IO.StringWriter()
        htmlWrite = New HtmlTextWriter(stringWrite)
        gView.GridLines = GridLines.None
        gView.RenderControl(htmlWrite)
        Response.Write(stringWrite.ToString())
        Response.End()
    End Sub


    Sub GenMenuLogon(page As String, course_id As Integer)
        Response.Write("</div><div class=""left2"">")
    End Sub

    Sub GenMenu(pageId As Integer, course_id As Integer)
        Dim sql As String
        Dim rs As System.Data.SqlClient.SqlDataReader
        sql = "select a.application_page_id, c.page from Application_Page_Assigned a, Course b, Application_Page c where a.application_id=b.application_id and a.active=1 and a.application_page_id=c.id and b.id=" & course_id & " order by a.display_order"
        rs = BQ(sql)
        If rs.Read() Then

            Do

                If pageId <> rs("application_page_id") Then
                    Response.Write("<p><a href=""" & "/apply.aspx?id=" & CInt(course_id) & "&p=" & rs("application_page_id") & """>" & rs("page") & "</a></p>")
                Else
                    Response.Write("<p>" & rs("page") & "</p>")
                End If
            Loop Until rs.Read() = 0

            If CheckCourse(course_id) Then
                If pageId = -1 Then
                    Response.Write("<p>Review/Print Your Application</p>")
                Else
                    Response.Write("<p><a href=""" & "/review.aspx?id=" & CInt(course_id) & """>Review/Print Your Application</a></p>")
                End If
            End If


            If pageId = 0 Then
                Response.Write("<p>Submit Your Application</p>")
            Else
                Response.Write("<p><a href=""" & "../app/submit.aspx?id=" & CInt(course_id) & """>Submit Your Application</a></p>")
            End If
            Response.Write("</div><div class=""left2"">")

        Else
            Response.Write("</div><div class=""left2"">")

        End If
        EQ(rs)

    End Sub



    Sub SetDataViewProperties(gView As GridView)
        gView.HeaderStyle.BackColor = Color.FromArgb(221, 221, 221)
        gView.HeaderStyle.ForeColor = Color.FromArgb(32, 32, 32)
        gView.AlternatingRowStyle.BackColor = Color.FromArgb(230, 230, 230)
        gView.cellPadding = 3
        gView.allowPaging = True
        gView.pageSize = ItemsPerPage
        gView.PagerSettings.Position = PagerPosition.TopAndBottom
        gView.BorderStyle = BorderStyle.Solid
        gView.BorderColor = System.Drawing.Color.FromArgb(160, 160, 160)
        gView.PagerSettings.Mode = PagerButtons.NumericFirstLast
    End Sub




    Function db(stringIn As String) As String
        db = Replace(stringIn, "'", "''")
    End Function

    Function BQ(sql As String) As System.Data.SqlClient.SqlDataReader
        Dim strConnect As String = ConfigurationManager.connectionStrings("pubsConnectionString").ConnectionString
        MyConnection = New System.Data.SqlClient.SqlConnection(strConnect)
        Dim myCommand As New System.Data.SqlClient.SqlCommand(sql, MyConnection)
        myCommand.CommandTimeout = 900
        MyConnection.Open()
        Dim rsEvent As System.Data.SqlClient.SqlDataReader
        rsEvent = myCommand.ExecuteReader()
        BQ = rsEvent
    End Function

    Function BQ2(sql As String) As System.Data.SqlClient.SqlDataReader
        Dim strConnect As String = ConfigurationManager.connectionStrings("pubsConnectionString").ConnectionString
        MyConnection2 = New System.Data.SqlClient.SqlConnection(strConnect)
        Dim myCommand As New System.Data.SqlClient.SqlCommand(sql, MyConnection2)
        myCommand.CommandTimeout = 900
        MyConnection2.Open()
        Dim rsEvent As System.Data.SqlClient.SqlDataReader
        rsEvent = myCommand.ExecuteReader()
        BQ2 = rsEvent
    End Function

    Function BQ3(sql As String) As System.Data.SqlClient.SqlDataReader
        Dim strConnect As String = ConfigurationManager.connectionStrings("pubsConnectionString").ConnectionString
        MyConnection3 = New System.Data.SqlClient.SqlConnection(strConnect)
        Dim myCommand As New System.Data.SqlClient.SqlCommand(sql, MyConnection3)
        myCommand.CommandTimeout = 900
        MyConnection3.Open()
        Dim rsEvent As System.Data.SqlClient.SqlDataReader
        rsEvent = myCommand.ExecuteReader()
        BQ3 = rsEvent
    End Function


    Function BQ4(sql As String) As System.Data.SqlClient.SqlDataReader
        Dim strConnect As String = ConfigurationManager.connectionStrings("pubsConnectionString").ConnectionString
        MyConnection4 = New System.Data.SqlClient.SqlConnection(strConnect)
        Dim myCommand As New System.Data.SqlClient.SqlCommand(sql, MyConnection4)
        myCommand.CommandTimeout = 900
        MyConnection4.Open()
        Dim rsEvent As System.Data.SqlClient.SqlDataReader
        rsEvent = myCommand.ExecuteReader()
        BQ4 = rsEvent
    End Function

    Sub EQ(rsEvent As System.Data.SqlClient.SqlDataReader)
        rsEvent.Close()
        MyConnection.Close()
    End Sub

    Sub EQ2(rsEvent As System.Data.SqlClient.SqlDataReader)
        rsEvent.Close()
        MyConnection2.Close()
    End Sub

    Sub EQ3(rsEvent As System.Data.SqlClient.SqlDataReader)
        rsEvent.Close()
        MyConnection3.Close()
    End Sub

    Sub EQ4(rsEvent As System.Data.SqlClient.SqlDataReader)
        rsEvent.Close()
        MyConnection4.Close()
    End Sub

    Function CheckCourse(course_id As Integer) As Boolean
        Dim sql, rs
        Dim result = True
        sql = "select IsNull(webex_meeting_id,'') as id from course_webex where course_id=" & CInt(course_id)
        rs = BQ(sql)
        If rs.Read() Then
            If rs("id") <> "" Then result = False
        End If
        EQ(rs)

        sql = "select IsNull(is_event,0) as event from course_webex where course_id=" & CInt(course_id)
        rs = BQ(sql)
        If rs.Read() Then
            If rs("event") = True Then result = False
        End If
        EQ(rs)

        Return result
    End Function

    Sub WriteHeaderLogon(course_id As Integer)

        course_id = CInt(course_id)

        Response.Write("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
        Response.Write("<html xmlns=""http://www.w3.org/1999/xhtml""><!-- InstanceBegin template=""/Templates/index.dwt"" codeOutsideHTMLIsLocked=""false"" -->")
        Response.Write("<head>")
        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />")
        Response.Write("<link rel=""schema.DC"" href=""http://purl.org/dc/elements/1.1/"" />")
        Response.Write("<link rel=""schema.DCTERMS"" href=""http://purl.org/dc/terms/"" />")
        Response.Write("<meta name=""DC.language"" scheme=""DCTERMS.ISO639-2""  content=""eng"" />")
        Response.Write("<meta name=""DC.type"" scheme=""DCMIType"" content=""Text"" />")
        Response.Write("<meta name=""DC.format"" content=""text/html;  charset=ISO-8859-1"" />")
        Response.Write("<meta name=""DC.identifier"" scheme=""DCTERMS.URI"" content=""META_URLhere"" />")
        Response.Write("<meta name=""DC.date.created"" scheme=""DCTERMS.W3CDTF"" content=""2007"" />")
        Response.Write("<meta name=""DC.date.modified"" scheme=""DCTERMS.W3CDTF"" content=""Aug 15 2007"" />")
        Response.Write("<meta name=""DC.subject"" xml:lang=""eng"" content="""" />")
        Response.Write("<meta name=""DC.publisher"" content=""The Board of Trustees of the Leland Stanford Junior University"" />")
        Response.Write("<meta name=""DC.contributor"" content=""Stanford Graduate School of Business"" />")
        Response.Write("<meta name=""DC.creator"" content=""GSB Web Editor"" />")
        Response.Write("<meta name=""DC.rights"" content=""&copy; Stanford Graduate School of Business"" />")
        Response.Write("<meta name=""verify-v1"" content=""rZ6sThU0QWxA3r0nPMXRS05s5Fl4dffmLOZpLpkRF3g="" />")
        Response.Write("<meta name=""msvalidate.01"" content=""229F801874BD0EF7F900E2DA99E7CD25"" />")
        Response.Write("<meta name=""verify-v1"" content=""5T23lTF21l9jdOKCgWNmr/wEwHmq66lFJ5jIpIBLJ7U="" />")
        Response.Write("<script type=""text/JavaScript"">")
        Response.Write("<!--")
        Response.Write("function MM_openBrWindow(theURL,winName,features) { //v2.0")
        Response.Write("  window.open(theURL,winName,features);")
        Response.Write("}")
        Response.Write("//-->")
        Response.Write("</script>")
        Response.Write("<!-- InstanceBeginEditable name=""head"" -->")
        Response.Write("<meta name=""google-site-verification"" content=""U9GIoGo7g27MqccvGRI0EC2uLg7vu4KF1HwzOJ350eM"" />")
        Response.Write("<meta name=""Keywords"" content=""MBA, Sloan Master's, executive education, PhD"" lang=""en-us"" xml:lang=""en-us"" />")
        Response.Write("<meta name=""Description"" content=""Stanford Graduate School of Business offers full-time MBA, PhD, and MS degree programs, executive programs for experienced managers, and lifelong learning opportunities for alumni."" />")
        Response.Write("<!-- InstanceEndEditable -->")
        Response.Write("<!-- InstanceParam name=""META_URL"" type=""text"" value=""http://www.gsb.stanford.edu/"" --><!-- InstanceParam name=""META_Date_Created"" type=""text"" value=""May 30, 2007"" --><!-- InstanceParam name=""META_Date_Modified"" type=""text"" value=""October 4, 2007"" --><!-- InstanceParam name=""META_Title"" type=""text"" value=""Stanford Graduate School of Business"" --><!-- InstanceParam name=""META_Description"" type=""text"" value=""Stanford Graduate School of Business homepage"" --><!-- InstanceParam name=""META_Publisher"" type=""text"" value=""The Board of Trustees of the Leland Stanford Junior University"" --><!-- InstanceParam name=""META_Contributor"" type=""text"" value=""Stanford Graduate School of Business"" --><!-- InstanceParam name=""META_Creator"" type=""text"" value=""GSB Web Editor"" --><!-- InstanceParam name=""META_Rights"" type=""text"" value=""© Stanford Graduate School of Business"" -->")
        Response.Write("<title>Stanford Graduate School of Business</title>")
        Response.Write("<link href=""/css/gsbhome.css"" rel=""stylesheet"" type=""text/css"" media=""all"" />")
        Response.Write("<link rel=""stylesheet"" type=""text/css"" media=""print"" href=""/css/print.css"" />")
        Response.Write("<style type=""text/css"">#noprint {display: block;}")
        Response.Write("<!--")
        Response.Write(".style1 {")
        Response.Write("	font-size: 14px;")
        Response.Write("	font-weight: bold;")
        Response.Write("	color: #C98818;")
        Response.Write("}")
        Response.Write(".style2 {font-size: 12px}")
        Response.Write("-->")
        Response.Write("</style>")
        Response.Write("</head>")
        Response.Write("<body>")
        Response.Write("<!-- InstanceBeginEditable name=""body"" -->")
        Response.Write("<!-- Test -->")
        Response.Write("<!-- ClickTale Top part -->")
        Response.Write("<script type=""text/javascript"">")
        Response.Write("var WRInitTime=(new Date()).getTime();")
        Response.Write("</script>")
        Response.Write("<!-- ClickTale end of Top part -->")
        Response.Write("<div id=""wrap1"">")
        Response.Write("<div id=""wrapper"" class=""mobilehidden"">")
        Response.Write("  <div id=""masthead"">")
        Response.Write("  <div id=""noprint"">")

        Response.Write("    <h1 id=""logo""><a href=""http://www.gsb.stanford.edu/""><img src=""https://gsb.stanford.edu/images/gsbhomelogo.gif"" alt=""Stanford Graduate School of Business"" width=""346"" height=""65"" /></a></h1>   		")

        If CheckCourse(course_id) Then
            Response.Write("   <ul id=""globalnav2"">")
            Response.Write("      <li><a href=""/application_logon.aspx?id=" & course_id & """>Logon</a> |</li>")
            Response.Write("      <li><a href=""/application_register.aspx?id=" & course_id & """>Register as New User</a> |</li>")
            Response.Write("      <li><a href=""/application_logon_password.aspx?id=" & course_id & """>Forgot Password</a> </li>")
            Response.Write("    </ul>")
        End If
        Response.Write("    <ul id=""globalnav1"">")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/az-index.html"">A-Z Index</a> | </li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/contact.html"">Find People</a> | </li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/about/gsbvisitors.html"">Visit</a></li>")
        Response.Write("    </ul>   	")
        Response.Write("    <form action=""http://gsbsearch.stanford.edu/query.html"" method=""get"" id=""search"">")
        Response.Write("      <input name=""qt"" alt=""search term"" title=""search term"" type=""text"" size=""25"" class=""anInput"" value=""Search the GSB"" onfocus=""this.value=''""/>")
        Response.Write("      <input src=""https://gsb.stanford.edu/images/gobutton.gif"" alt=""Go"" type=""image"" class=""go"" style=""width:34; height:18;"" />")
        Response.Write("      <input type=""hidden"" name=""qp"" value="""" />")
        Response.Write("      <input type=""hidden"" name=""style"" value=""public"" />")
        Response.Write("      <input type=""hidden"" name=""col"" value=""public"" />")
        Response.Write("      <input type=""hidden"" name=""qs"" value="""" />")
        Response.Write("      <input type=""hidden"" name=""qc"" value="""" />")
        Response.Write("    </form>")
        Response.Write("  </div>")
        Response.Write("  </div>")
        Response.Write("  <div id=""content"">   ")
        Response.Write("    <div id=""belowFold"">   ")
        Response.Write("      <div class=""left1"">")
        Dim sql, rs
        Dim oExcep As Exception
        Try
            sql = "select name from course where id=" & course_id
            rs = BQ(sql)
            If rs.Read() Then
                Response.Write("<h4>" & rs("name") & "</h4>")
            End If
            EQ(rs)
        Catch
            oExcep = New Exception("An attempt has been made to render this page with an invalid program id.")
            Throw oExcep

        End Try
    End Sub


    Sub WriteHeader(course_id As Integer)
        course_id = CInt(course_id)

        Response.Write("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
        Response.Write("<html xmlns=""http://www.w3.org/1999/xhtml""><!-- InstanceBegin template=""/Templates/index.dwt"" codeOutsideHTMLIsLocked=""false"" -->")
        Response.Write("<head>")
        Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />")
        Response.Write("<link rel=""schema.DC"" href=""http://purl.org/dc/elements/1.1/"" />")
        Response.Write("<link rel=""schema.DCTERMS"" href=""http://purl.org/dc/terms/"" />")
        Response.Write("<meta name=""DC.language"" scheme=""DCTERMS.ISO639-2""  content=""eng"" />")
        Response.Write("<meta name=""DC.type"" scheme=""DCMIType"" content=""Text"" />")
        Response.Write("<meta name=""DC.format"" content=""text/html;  charset=ISO-8859-1"" />")
        Response.Write("<meta name=""DC.identifier"" scheme=""DCTERMS.URI"" content=""META_URLhere"" />")
        Response.Write("<meta name=""DC.date.created"" scheme=""DCTERMS.W3CDTF"" content=""2007"" />")
        Response.Write("<meta name=""DC.date.modified"" scheme=""DCTERMS.W3CDTF"" content=""Aug 15 2007"" />")
        Response.Write("<meta name=""DC.subject"" xml:lang=""eng"" content="""" />")
        Response.Write("<meta name=""DC.publisher"" content=""The Board of Trustees of the Leland Stanford Junior University"" />")
        Response.Write("<meta name=""DC.contributor"" content=""Stanford Graduate School of Business"" />")
        Response.Write("<meta name=""DC.creator"" content=""GSB Web Editor"" />")
        Response.Write("<meta name=""DC.rights"" content=""&copy; Stanford Graduate School of Business"" />")
        Response.Write("<meta name=""verify-v1"" content=""rZ6sThU0QWxA3r0nPMXRS05s5Fl4dffmLOZpLpkRF3g="" />")
        Response.Write("<meta name=""msvalidate.01"" content=""229F801874BD0EF7F900E2DA99E7CD25"" />")
        Response.Write("<meta name=""verify-v1"" content=""5T23lTF21l9jdOKCgWNmr/wEwHmq66lFJ5jIpIBLJ7U="" />")
        Response.Write("<script type=""text/JavaScript"">")
        Response.Write("<!--")
        Response.Write("function MM_openBrWindow(theURL,winName,features) { //v2.0")
        Response.Write("  window.open(theURL,winName,features);")
        Response.Write("}")
        Response.Write("//-->")
        Response.Write("</script>")
        Response.Write("<!-- InstanceBeginEditable name=""head"" -->")
        Response.Write("<meta name=""google-site-verification"" content=""U9GIoGo7g27MqccvGRI0EC2uLg7vu4KF1HwzOJ350eM"" />")
        Response.Write("<meta name=""Keywords"" content=""MBA, Sloan Master's, executive education, PhD"" lang=""en-us"" xml:lang=""en-us"" />")
        Response.Write("<meta name=""Description"" content=""Stanford Graduate School of Business offers full-time MBA, PhD, and MS degree programs, executive programs for experienced managers, and lifelong learning opportunities for alumni."" />")
        Response.Write("<!-- InstanceEndEditable -->")
        Response.Write("<!-- InstanceParam name=""META_URL"" type=""text"" value=""http://www.gsb.stanford.edu/"" --><!-- InstanceParam name=""META_Date_Created"" type=""text"" value=""May 30, 2007"" --><!-- InstanceParam name=""META_Date_Modified"" type=""text"" value=""October 4, 2007"" --><!-- InstanceParam name=""META_Title"" type=""text"" value=""Stanford Graduate School of Business"" --><!-- InstanceParam name=""META_Description"" type=""text"" value=""Stanford Graduate School of Business homepage"" --><!-- InstanceParam name=""META_Publisher"" type=""text"" value=""The Board of Trustees of the Leland Stanford Junior University"" --><!-- InstanceParam name=""META_Contributor"" type=""text"" value=""Stanford Graduate School of Business"" --><!-- InstanceParam name=""META_Creator"" type=""text"" value=""GSB Web Editor"" --><!-- InstanceParam name=""META_Rights"" type=""text"" value=""© Stanford Graduate School of Business"" -->")
        Response.Write("<title>Stanford Graduate School of Business</title>")
        Response.Write("<link href=""/css/gsbhome.css"" rel=""stylesheet"" type=""text/css"" media=""all"" />")
        'Response.Write("<link rel=""stylesheet"" type=""text/css"" media=""print"" href=""http://gsb.stanford.edu/css/gsbhomeprint.css"" />")
        Response.Write("<style type=""text/css"">")
        Response.Write("<!--")
        Response.Write(".style1 {")
        Response.Write("	font-size: 14px;")
        Response.Write("	font-weight: bold;")
        Response.Write("	color: #C98818;")
        Response.Write("}")
        Response.Write(".style2 {font-size: 12px}")
        Response.Write("-->")
        Response.Write("</style>")
        Response.Write("</head>")
        Response.Write("<body>")
        Response.Write("<!-- InstanceBeginEditable name=""body"" -->")
        Response.Write("<!-- Test -->")
        Response.Write("<!-- ClickTale Top part -->")
        Response.Write("<script type=""text/javascript"">")
        Response.Write("var WRInitTime=(new Date()).getTime();")
        Response.Write("</script>")
        Response.Write("<!-- ClickTale end of Top part -->")
        Response.Write("<div id=""wrap1"">")
        Response.Write("<div id=""wrapper"" class=""mobilehidden"">")
        Response.Write("  <div id=""masthead"">")
        Response.Write("    <h1 id=""logo""><a href=""http://www.gsb.stanford.edu/""><img src=""https://gsb.stanford.edu/images/gsbhomelogo.gif"" alt=""Stanford Graduate School of Business"" width=""346"" height=""65"" /></a></h1>   		")

        Dim sql, rs

        If CheckCourse(course_id) Then

            Response.Write("   <ul id=""globalnav2"">")
            sql = "select a.application_page_id, c.page from Application_Page_Assigned a, Course b, Application_Page c where a.application_id=b.application_id and a.active=1 and a.application_page_id=c.id and b.id=" & course_id & " order by a.display_order"
            rs = BQ(sql)
            If rs.Read() Then
                Response.Write("      <li><a href=""/apply.aspx?id=" & course_id & "&amp;p=" & rs("application_page_id") & """>Home</a> |</li>")
            End If
            EQ(rs)


            Response.Write("      <li><a href=""/faq.aspx?id=" & course_id & """>Frequently Asked Questions</a> | </li>")
            Response.Write("      <li><a href=""/update_profile.aspx?id=" & course_id & """>Update Profile</a> | </li>")
            Response.Write("       <li><a href=""/submit.aspx?id=" & course_id & """>Submit</a>  | </li>")
            Response.Write("      <li><a href=""/logout.aspx?id=" & course_id & """>Logout</a></li>")
            Response.Write("    </ul>")
        End If
        Response.Write("    <ul id=""globalnav1"">")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/az-index.html"">A-Z Index</a> | </li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/contact.html"">Find People</a> | </li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/about/gsbvisitors.html"">Visit</a></li>")
        Response.Write("    </ul>   	")
        Response.Write("    <form action=""http://gsbsearch.stanford.edu/query.html"" method=""get"" id=""search"">")
        Response.Write("      <input name=""qt"" alt=""search term"" title=""search term"" type=""text"" size=""25"" class=""anInput"" value=""Search the GSB"" onfocus=""this.value=''""/>")
        Response.Write("      <input src=""https://gsb.stanford.edu/images/gobutton.gif"" alt=""Go"" type=""image"" class=""go"" style=""width:34; height:18;"" />")
        Response.Write("      <input type=""hidden"" name=""qp"" value="""" />")
        Response.Write("      <input type=""hidden"" name=""style"" value=""public"" />")
        Response.Write("      <input type=""hidden"" name=""col"" value=""public"" />")
        Response.Write("      <input type=""hidden"" name=""qs"" value="""" />")
        Response.Write("      <input type=""hidden"" name=""qc"" value="""" />")
        Response.Write("    </form>")
        Response.Write("  </div>")
        Response.Write("  <div id=""content"">   ")
        Response.Write("    <div id=""belowFold"">   ")
        Response.Write("      <div class=""left1"">")

        Dim oExcep As Exception
        Try
            sql = "select name from course where id=" & CInt(Request.QueryString("id"))
            rs = BQ(sql)
            If rs.Read() Then
                Response.Write("<h4>" & rs("name") & "</h4>")
            End If
            EQ(rs)
        Catch
            oExcep = New Exception("An attempt has been made to render this page with an invalid program id.")
            Throw oExcep

        End Try
    End Sub

    Sub WriteStartPage()
        Response.Write("<div id=""wrapper"">")
    End Sub



    Sub WriteEndPage(pageId As Integer, course_id As Integer)
        course_id = CInt(course_id)
        Response.Write("</div>")
        Response.Write("      ")
        Dim application_id = 0
        Dim sql, rs
        sql = "select IsNull(application_id,0) as application_id from Course where id=" & course_id
        rs = BQ(sql)
        If rs.Read() Then
            application_id = rs("application_id")
        End If
        EQ(rs)
        sql = "select b.id, b.page from application_page_instructions_assigned a, application_page b where b.id=a.application_page_instruction_id and a.application_id=" & application_id & " and a.application_page_id=" & pageId
        rs = BQ(sql)
        If rs.Read() Then
            Response.Write("      <div class=""left3"">")
            Response.Write("        <h2 class=""instructionLinks""></h2>")
            Response.Write("        <ul>")
            Do
                Response.Write("          <li><a href=""/apply.aspx?id=" & CInt(Request.QueryString("id")) & "&p=" & rs("id") & """>" & rs("page") & "</a></li>")
            Loop Until rs.Read() = 0

            Response.Write("        </ul><br/>")
            Response.Write("      </div>")
        End If
        EQ(rs)
        Response.Write("    </div>")
        Response.Write("  </div>")
        Response.Write("  <div id=""content_bot"">&nbsp;</div>")
        Response.Write("  <div id=""footer"">")
        Response.Write("    <ul>")
        Response.Write("      <li class=""first""><a href=""http://www.gsb.stanford.edu/admissions/""> Admissions</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/research/""> Faculty &amp; Research</a></li>")
        Response.Write("      <li><a href=""http://alumni.gsb.stanford.edu/"">Alumni</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/corporate/"">Recruiters &amp; Companies</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/giving/"">Giving</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/exed/"">Executive Education</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/news/"">News</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/about/"">About the GSB</a></li>")
        Response.Write("    </ul>")
        Response.Write("    <p>Copyright &copy;" & Now().Year & " Stanford Graduate School of Business | <a href=""http://www.gsb.stanford.edu/help.html"">Site")
        Response.Write("      Help</a> | <a href=""http://www.gsb.stanford.edu/policy/terms.html"">Terms of Use</a> | <a href=""http://www.stanford.edu/"">Stanford University</a></p>")
        Response.Write("  </div>")
        Response.Write("</div>")
        'Response.Write("<img src=""/images/logo.gif"" width=""239"" height=""51"" alt=""Stanford Graduate School of Business"" class=""mobile"" />")
        'Response.Write("<div class=""mobile"" id=""right"">")
        'Response.Write("		<p><strong>Stanford Graduate School of Business</strong><br />")
        'Response.Write("				518 Memorial Way<br />")
        'Response.Write("				Stanford, CS 94305-5015<br />")
        'Response.Write("				Phone: 650-723-2146<br />")
        'Response.Write("		</p>")
        'Response.Write("		<form action=""/includes/post"" method=""get"" id=""search_mobile"">")
        'Response.Write("				<label class=""mobile"">Search the GSB:</label>")
        'Response.Write("				<select name=""Search_Options"" title=""search options"" size=""1"" class=""dropdown"" >")
        'Response.Write("						<option>Search the GSB</option>")
        'Response.Write("						<option>Search GSB news</option>")
        'Response.Write("				</select>")
        'Response.Write("				<input name=""Search_Term"" alt=""search term"" title=""search term"" type=""text"" size=""25"" class=""anInput"" />")
        'Response.Write("				<input src=""/images/gobutton.gif"" alt=""Go"" type=""image"" class=""go"" style=""width:34; height:18;"" />")
        'Response.Write("		</form>")
        'Response.Write("		<div class=""mobileheadline""><strong>How to find us</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""/about/directions.html"">Directions</a></li>")
        'Response.Write("				<li><a href=""http://www.stanford.edu/home/visitors/maps.html"">Stanford Campus Maps</a></li>")
        'Response.Write("				<li> <a href=""http://transportation.stanford.edu/parking_info/VisitorParking.shtml"">Visitor Parking</a></li>")
        'Response.Write("				<li> <a href=""http://transportation.stanford.edu/alt_transportation/AlternateTransportation.shtml"">Transportation   			Alternatives</a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>Directory</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""/contact.html"">GSB Office Directory</a></li>")
        'Response.Write("				<li><a href=""http://stanfordwho.stanford.edu/lookup"">Stanford public directory</a> </li>")
        'Response.Write("		</ul>")
        'Response.Write("		<p><a href=""#""><b>Programs Info/Applications</b></a><br />")
        'Response.Write("				Links to information and application to the MBA Program, PhD Program, Sloan Master's program, Executive Education, Summer Institutes</p>")
        'Response.Write("		<div class=""mobileheadline""><strong>Program Sites</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""/mba/""><strong>MBA</strong></a></li>")
        'Response.Write("				<li><a href=""/sloan/""><strong>Sloan</strong></a></li>")
        'Response.Write("				<li><a href=""/phd/""><strong>Ph.D.</strong></a></li>")
        'Response.Write("				<li><a href=""/exed/""><strong>Executive Education</strong></a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<p><a href=""#""><b>About the GSB</b></a><br />")
        'Response.Write("				Our mission and more information on the New Curriculum, New Collaborations, and New Campus.</p>")
        'Response.Write("		<div class=""mobileheadline""><strong>GSB Students</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""#"">My GSB</a> (login required)</li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>Alumni</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""#"">Public Directory</a></li>")
        'Response.Write("				<li><a href=""#"">Visitor Info</a></li>")
        'Response.Write("				<li><a href=""#"">GSB Office Directory</a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>News</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""http://www.gsb.stanford.edu/news/"">GSB News</a></li>")
        'Response.Write("				<li><a href=""#"">Knowledgebase</a></li>")
        'Response.Write("				<li><a href=""#"">@GSB Today</a></li>")
        'Response.Write("				<li><a href=""#"">Stanford Business Magazine</a></li>")
        'Response.Write("				<li>Press Contact: <a href=""#"">Helen K. Chang</a>, 650-723-3358</li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>Date Book</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""#"">Academic Calendar</a></li>")
        'Response.Write("				<li><a href=""#"">Events at Stanford</a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<p>")
        'Response.Write("			<a href=""http://www.gsb.stanford.edu/corporate/""><b>Recruiters &amp; Companies</b></a><br />")
        'Response.Write("			<a href=""http://www.gsb.stanford.edu/giving/""><b>Giving to the GSB</b></a></p>")
        'Response.Write("		<p class=""mobilecopy"" style=""margin-left:0px;"">Copyright &copy;" & Now().Year & " Stanford Graduate School of Business</p>")
        'Response.Write("</div>")
        Response.Write("</div>")
        Response.Write("<!-- InstanceEndEditable -->")
        Response.Write("</body>")
        Response.Write("</html>")
    End Sub


    Sub WriteEndPageLogon(course_id As Integer)
        course_id = CInt(course_id)
        Response.Write("</div>")
        Response.Write("      ")
        Response.Write("    </div>")
        Response.Write("  </div>")
        Response.Write("  <div id=""content_bot"">&nbsp;</div>")
        Response.Write("  <div id=""footer"">")
        Response.Write("    <ul>")
        Response.Write("      <li class=""first""><a href=""http://www.gsb.stanford.edu/admissions/""> Admissions</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/research/""> Faculty &amp; Research</a></li>")
        Response.Write("      <li><a href=""http://alumni.gsb.stanford.edu/"">Alumni</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/corporate/"">Recruiters &amp; Companies</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/giving/"">Giving</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/exed/"">Executive Education</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/news/"">News</a></li>")
        Response.Write("      <li><a href=""http://www.gsb.stanford.edu/about/"">About the GSB</a></li>")
        Response.Write("    </ul>")
        Response.Write("    <p>Copyright &copy;" & Now().Year & " Stanford Graduate School of Business | <a href=""http://www.gsb.stanford.edu/help.html"">Site")
        Response.Write("      Help</a> | <a href=""http://www.gsb.stanford.edu/policy/terms.html"">Terms of Use</a> | <a href=""http://www.stanford.edu/"">Stanford University</a></p>")
        Response.Write("  </div>")
        Response.Write("</div>")
        'Response.Write("<img src=""/images/logo.gif"" width=""239"" height=""51"" alt=""Stanford Graduate School of Business"" class=""mobile"" />")
        'Response.Write("<div class=""mobile"" id=""right"">")
        'Response.Write("		<p><strong>Stanford Graduate School of Business</strong><br />")
        'Response.Write("				518 Memorial Way<br />")
        'Response.Write("				Stanford, CS 94305-5015<br />")
        'Response.Write("				Phone: 650-723-2146<br />")
        'Response.Write("		</p>")
        'Response.Write("		<form action=""/includes/post"" method=""get"" id=""search_mobile"">")
        'Response.Write("				<label class=""mobile"">Search the GSB:</label>")
        'Response.Write("				<select name=""Search_Options"" title=""search options"" size=""1"" class=""dropdown"" >")
        'Response.Write("						<option>Search the GSB</option>")
        'Response.Write("						<option>Search GSB news</option>")
        'Response.Write("				</select>")
        'Response.Write("				<input name=""Search_Term"" alt=""search term"" title=""search term"" type=""text"" size=""25"" class=""anInput"" />")
        'Response.Write("				<input src=""/images/gobutton.gif"" alt=""Go"" type=""image"" class=""go"" style=""width:34; height:18;"" />")
        'Response.Write("		</form>")
        'Response.Write("		<div class=""mobileheadline""><strong>How to find us</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""/about/directions.html"">Directions</a></li>")
        'Response.Write("				<li><a href=""http://www.stanford.edu/home/visitors/maps.html"">Stanford Campus Maps</a></li>")
        'Response.Write("				<li> <a href=""http://transportation.stanford.edu/parking_info/VisitorParking.shtml"">Visitor Parking</a></li>")
        'Response.Write("				<li> <a href=""http://transportation.stanford.edu/alt_transportation/AlternateTransportation.shtml"">Transportation   			Alternatives</a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>Directory</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""/contact.html"">GSB Office Directory</a></li>")
        'Response.Write("				<li><a href=""http://stanfordwho.stanford.edu/lookup"">Stanford public directory</a> </li>")
        'Response.Write("		</ul>")
        'Response.Write("		<p><a href=""#""><b>Programs Info/Applications</b></a><br />")
        'Response.Write("				Links to information and application to the MBA Program, PhD Program, Sloan Master's program, Executive Education, Summer Institutes</p>")
        'Response.Write("		<div class=""mobileheadline""><strong>Program Sites</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""/mba/""><strong>MBA</strong></a></li>")
        'Response.Write("				<li><a href=""/sloan/""><strong>Sloan</strong></a></li>")
        'Response.Write("				<li><a href=""/phd/""><strong>Ph.D.</strong></a></li>")
        'Response.Write("				<li><a href=""/exed/""><strong>Executive Education</strong></a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<p><a href=""#""><b>About the GSB</b></a><br />")
        'Response.Write("				Our mission and more information on the New Curriculum, New Collaborations, and New Campus.</p>")
        'Response.Write("		<div class=""mobileheadline""><strong>GSB Students</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""#"">My GSB</a> (login required)</li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>Alumni</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""#"">Public Directory</a></li>")
        'Response.Write("				<li><a href=""#"">Visitor Info</a></li>")
        'Response.Write("				<li><a href=""#"">GSB Office Directory</a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>News</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""http://www.gsb.stanford.edu/news/"">GSB News</a></li>")
        'Response.Write("				<li><a href=""#"">Knowledgebase</a></li>")
        'Response.Write("				<li><a href=""#"">@GSB Today</a></li>")
        'Response.Write("				<li><a href=""#"">Stanford Business Magazine</a></li>")
        'Response.Write("				<li>Press Contact: <a href=""#"">Helen K. Chang</a>, 650-723-3358</li>")
        'Response.Write("		</ul>")
        'Response.Write("		<div class=""mobileheadline""><strong>Date Book</strong></div>")
        'Response.Write("		<ul>")
        'Response.Write("				<li><a href=""#"">Academic Calendar</a></li>")
        'Response.Write("				<li><a href=""#"">Events at Stanford</a></li>")
        'Response.Write("		</ul>")
        'Response.Write("		<p>")
        'Response.Write("			<a href=""http://www.gsb.stanford.edu/corporate/""><b>Recruiters &amp; Companies</b></a><br />")
        'Response.Write("			<a href=""http://www.gsb.stanford.edu/giving/""><b>Giving to the GSB</b></a></p>")
        'Response.Write("		<p class=""mobilecopy"" style=""margin-left:0px;"">Copyright &copy;" & Now().Year & " Stanford Graduate School of Business</p>")
        'Response.Write("</div>")
        Response.Write("</div>")
        Response.Write("<!-- InstanceEndEditable -->")
        Response.Write("</body>")
        Response.Write("</html>")
    End Sub




    Sub RenderYesNo(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div>")
        Response.Write("<div class=""item-controls""><ul>")
        Response.Write("<li><input type=""radio"" name=""q" & rs("id") & """ ")
        If subValue = "Yes" Then Response.Write(" checked=""checked"" ")
        Response.Write("value=""Yes""/>")
        Response.Write("<label>Yes</label></li>")
        Response.Write("<li><input type=""radio"" name=""q" & rs("id") & """ ")
        If subValue = "No" Then Response.Write(" checked=""checked"" ")
        Response.Write("value=""No""/>")
        Response.Write("<label>No</label></li>")
        Response.Write("</ul>")
        Response.Write("</div></div>")
    End Sub

    Sub RenderNumericRating(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        Response.Write("<tr><td valign=""top"">" & rs("label") & "</td>")
        Response.Write("<td nowrap=""nowrap"">")
        sql2 = "select * from Application_Content_Numeric where application_question_id=" & rs("id")
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<table>")
            If rs2("use_labels") = 1 Then
                Response.Write("<tr>")
                Response.Write("<td align=""center"">" & rs2("label1") & "</td>")
                Response.Write("<td align=""center"">" & rs2("label2") & "</td>")
                Response.Write("<td align=""center"">" & rs2("label3") & "</td>")
                Response.Write("<td align=""center"">" & rs2("label4") & "</td>")
                Response.Write("<td align=""center"">" & rs2("label5") & "</td>")
                Response.Write("</tr>")
            Else
                Response.Write("<tr><td align=""center"">1</td><td align=""center"">2</td><td align=""center"">3</td><td align=""center"">4</td><td align=""center"">5</td></tr>")
            End If
            Response.Write("<tr>")
            Dim J
            For J = 1 To 5
                Response.Write("<td align=""center"">")
                Response.Write("<input type=""radio"" name=""q" & rs("id") & """ ")
                Try
                    If subValue = J Then Response.Write(" checked=""checked"" ")
                Catch
                End Try
                Response.Write(" value=""" & J & """/>")
                Response.Write("</td>")
            Next J
            Response.Write("</tr>")
            Response.Write("</table>")
        End If
        EQ2(rs2)
        Response.Write("</td>")
        Response.Write("</tr>")
    End Sub

    Sub RenderText(rs)
        Response.Write("<tr><td colspan=""2"">" & rs("label") & "</td>")
        Response.Write("</tr>")
    End Sub

    Sub RenderTextHeader(rs)
        Response.Write("<div class=""divider"">&nbsp;</div>")

        Response.Write("<div class=""section-title""><h3>" & rs("label") & "</h3></div>")

    End Sub

    Sub RenderTextMultiline(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        Response.Write("<div class=""item""><div class=""item-label"">" & rs("label"))
        If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")
        Response.Write("</div>")
        Response.Write("<div class=""item-controls""><textarea  ")
        If rs("required") = True Then Response.Write("class=""validate[required]"" ")
        Response.Write("rows=""6"" cols=""120"" id=""q" & rs("id") & """ name=""q" & rs("id") & """>" & subValue & "</textarea></div>")
        Response.Write("</div>")
    End Sub





    Sub ParseDateOfBirth(birthDate As String, ByRef month As String, ByRef day As String, ByRef year As String)

        Dim exp As Match
        birthDate = LTrim(RTrim(birthDate))
        exp = Regex.Match(birthDate, "^(\d{1,2})/(\d{1,2})/(\d{4})$", RegexOptions.IgnoreCase)

        If (exp.Success) Then

            month = exp.Groups(1).Value
            day = exp.Groups(2).Value
            year = exp.Groups(3).Value
        End If

    End Sub


    Sub RenderDateOfBirth(rs)
        Try
            Dim sql2, rs2
            Dim subValue = ""
            Dim month, day, year
            Dim MonthName(13) As String

            MonthName(1) = "January"
            MonthName(2) = "February"
            MonthName(3) = "March"
            MonthName(4) = "April"
            MonthName(5) = "May"
            MonthName(6) = "June"
            MonthName(7) = "July"
            MonthName(8) = "August"
            MonthName(9) = "September"
            MonthName(10) = "October"
            MonthName(11) = "November"
            MonthName(12) = "December"

            Dim J
            month = ""
            day = ""
            year = ""
            If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
                sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
                rs2 = BQ2(sql2)
                If rs2.Read() Then
                    subValue = rs2("value")
                    Try
                        ParseDateOfBirth(subValue, month, day, year)
                    Catch ex As Exception

                    End Try
                End If
                EQ2(rs2)

            End If

            Response.Write("<div class=""item"">")
            Response.Write("<div class=""item-label"">" & rs("label"))
            If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")
            Response.Write("</div>")
            Response.Write("<div class=""item-controls"">")

            Response.Write("<select id=""q1-" & rs("id") & "name=""q1-" & rs("id") & """ ") 'Revamp styling change
            If rs("required") = True Then Response.Write("class=""validate[required] required"" ")
            Response.Write(">")

            Response.Write("<option class=""group-title"" value="""">MONTH</option>")
            For J = 1 To 12
                Response.Write("<option value=""" & J & """")
                Try
                    If month = J Then Response.Write(" selected=""selected""")
                Catch
                End Try
                Response.Write(">")
                Response.Write(MonthName(J))
                Response.Write("</option>")
            Next
            Response.Write("</select>")

            Response.Write("<select id=""q2-" & rs("id") & " name=""q2-" & rs("id") & """ ")  'Revamp styling change
            If rs("required") = True Then Response.Write("class=""validate[required] required"" ")
            Response.Write(">")

            Response.Write("<option class=""group-title"" value="""">DAY</option>")
            For J = 1 To 31
                Response.Write("<option value=""" & J & """")
                Try
                    If day = J Then Response.Write(" selected=""selected""")
                Catch
                End Try
                Response.Write(">")
                Response.Write(J)
                Response.Write("</option>")
            Next
            Response.Write("</select>")

            Response.Write("<select id=""q3-" & rs("id") & " name=""q3-" & rs("id") & """ ")  'Revamp styling change
            If rs("required") = True Then Response.Write("class=""validate[required] required"" ")
            Response.Write(">")

            Response.Write("<option class=""group-title"" value="""">YEAR</option>")
            For J = 1905 To Now.Year()
                Response.Write("<option value=""" & J & """")
                Try
                    If year = J Then Response.Write(" selected=""selected""")
                Catch
                End Try
                Response.Write(">")
                Response.Write(J)
                Response.Write("</option>")
            Next
            Response.Write("</select>")


            Response.Write("</div>")
            Response.Write("</div>")
        Catch ex As Exception

        End Try
    End Sub





    Sub ParsePhone(subValue As String, ByRef subValue1 As String, ByRef subValue2 As String, ByRef subValue3 As String, ByRef subValue4 As String)


        Dim exp As Match
        subValue = LTrim(RTrim(subValue))
        exp = Regex.Match(subValue, "^[\+]?(\d{2,3})?[ ]?\((\d{3})\)[ ]?(\d{3})[-](\d{4}$)", RegexOptions.IgnoreCase)


        If (exp.Success) Then

            subValue1 = exp.Groups(1).Value
            subValue2 = exp.Groups(2).Value
            subValue3 = exp.Groups(3).Value
            subValue4 = exp.Groups(4).Value
        End If
    End Sub

    Sub RenderPhone(rs)
        Dim sql2, rs2
        Dim subValue = ""
        Dim subValue1, subValue2, subValue3, subValue4
        subValue1 = ""
        subValue2 = ""
        subValue3 = ""
        subValue4 = ""

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
                ParsePhone(subValue, subValue1, subValue2, subValue3, subValue4)
            End If
            EQ2(rs2)

        End If

        Response.Write("<div class=""item"">")
        Response.Write("<div class=""item-label"">" & rs("label"))
        If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")
        Response.Write("<div class=""body-emph"">Format US number as XXX-YYY-ZZZZ.<br />Include country codes for international number.</div></div>")
        Response.Write("<div class=""item-controls""><ul class=""grouped-box-item"">")

        Response.Write("<li><input type=""text"" maxlength=""3"" id=""q1-" & rs("id") & """ name=""q1-" & rs("id") & """ ")
        Response.Write("class=""group-box group-box-01 ")

        Response.Write(""" ")
        Response.Write("value=""" & subValue1 & """ /></li>")

        Response.Write("<li><input type=""text""  maxlength=""3"" id=""q2-" & rs("id") & """ name=""q2-" & rs("id") & """ ")
        Response.Write("class=""group-box group-box-02 ")
        If rs("required") = True Then Response.Write(" validate[required,custom[onlyNumberSp]] ")
        Response.Write(""" ")
        Response.Write("value=""" & subValue2 & """ /></li>")

        Response.Write("<li><input type=""text""  maxlength=""3"" id=""q3-" & rs("id") & """ name=""q3-" & rs("id") & """ ")
        Response.Write("class=""group-box group-box-02 ")
        If rs("required") = True Then Response.Write(" validate[required,custom[onlyNumberSp]] ")
        Response.Write(""" ")
        Response.Write("value=""" & subValue3 & """ /></li>")

        Response.Write("<li><input type=""text""  maxlength=""4"" id=""q4-" & rs("id") & """ name=""q4-" & rs("id") & """ ")
        Response.Write("class=""group-box group-box-03 ")
        If rs("required") = True Then Response.Write(" validate[required,custom[onlyNumberSp]] ")
        Response.Write(""" ")
        Response.Write("value=""" & subValue4 & """ /></li>")


        Response.Write("</ul></div>")
        Response.Write("</div>")

        'Script to remove background
        If subValue2 <> "" Then
            Response.Write("<script type=""text/javascript"">")
            Response.Write(vbCrLf & "$('.group-box-01').css(""background"",""none #fafafa"");")
            Response.Write(vbCrLf & "$('.group-box-02').css(""background"",""none #fafafa"");")
            Response.Write(vbCrLf & "$('.group-box-03').css(""background"",""none #fafafa"");")
            Response.Write(vbCrLf & "</script>")
        End If
    End Sub

    Sub RenderTextInput(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)

        End If

        Response.Write("<div class=""item"">")
        Response.Write("<div class=""item-label"">" & rs("label"))
        If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")
        Response.Write("</div>")
        Response.Write("<div class=""item-controls""><input type=""text"" id=""q" & rs("id") & """ name=""q" & rs("id") & """ ")
        If rs("required") = True Then Response.Write("class=""validate[required]"" ")
        Response.Write("value=""" & subValue & """ />")
        Response.Write("</div>")
        Response.Write("</div>")
    End Sub

    Sub RenderTextInputSize(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)

        End If
        Dim size = 40
        sql2 = "select size from application_content_input_size where application_question_id=" & rs("id")
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            size = rs2("size")
        End If
        EQ2(rs2)

        Response.Write("<div class=""item"">")
        Response.Write("<div class=""item-label"">" & rs("label") & "</div>")
        Response.Write("<div class=""item-controls""><input type=""text"" size=""" & size & """ id=""q" & rs("id") & """  name=""q" & rs("id") & """ value=""" & subValue & """ />")
        Response.Write("</div>")
        Response.Write("</div>")
    End Sub

    Sub RenderTextInputEmailLength(rs, Length)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If


        Response.Write("<div class=""item"">")
        Response.Write("<div class=""item-label"">" & rs("label"))
        If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")
        Response.Write("</div><div class=""item-controls""><input type=""text"" id=""q" & rs("id") & """ name=""q" & rs("id") & """ value=""" & subValue & """ ")
        If rs("required") = True Then Response.Write("class=""validate[required,custom[email]]"" ")

        Response.Write("maxlength=""" & length & """ />")
        Response.Write("</div>")
        Response.Write("</div>")


    End Sub



    Sub RenderTextInputLength(rs, length)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If


        Response.Write("<div class=""item"">")
        Response.Write("<div class=""item-label"">" & rs("label"))
        If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")
        Response.Write("</div><div class=""item-controls""><input type=""text"" id=""q" & rs("id") & """ name=""q" & rs("id") & """ value=""" & subValue & """ ")
        If rs("required") = True Then Response.Write("class=""validate[required]"" ")

        Response.Write("maxlength=""" & length & """ />")
        Response.Write("</div>")
        Response.Write("</div>")


    End Sub

    Sub RenderFile(rs)
        Dim sql2, rs2
        Response.Write("<div class=""item"">")
        Response.Write("<div class=""item-label"">" & rs("label") & "</div>")
        Response.Write("<div class=""item-controls""><div class=""file-browser""><input type=""file"" name=""q" & rs("id") & """/>")
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select filename from Application_submission_File where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                Response.Write("<br />" & rs2("filename") & " uploaded.  You may select a new file to replace this file.")
            End If
            EQ2(rs2)
        End If
        Response.Write("</div>")
        Response.Write("</div>")
        Response.Write("</div>")
    End Sub

    Sub RenderDropDown(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        ' Check to see if the drop down has elements
        sql2 = "select id,item from application_content_choice where application_question_id=" & rs("id") & " order by display_order,item"
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div>")
            Response.Write("<div class=""item-controls""><select name=""q" & rs("id") & """>")
            Do
                Response.Write("<option value=""" & rs2("id") & """")
                If CStr(subValue) = CStr(rs2("id")) Then Response.Write(" selected=""selected"" ")
                Response.Write(">" & rs2("item") & "</option>")
            Loop Until rs2.Read() = 0
            Response.Write("</select></div>")
            Response.Write("</div>")
        End If
        EQ2(rs2)
    End Sub

    Sub RenderState(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        ' Check to see if the drop down has elements
        sql2 = "select id,state from state order by state"
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div>")
            Response.Write("<div class=""item-controls""><select name=""q" & rs("id") & """>")
            Response.Write("<option selected=""selected"" class=""group-title"" value=""null"">Select " & rs("label") &"</option>")  'Adding label in default option for Revamp
            Do
                Response.Write("<option value=""" & rs2("id") & """")
                If CStr(subValue) = CStr(rs2("id")) Then Response.Write(" selected=""selected"" ")
                Response.Write(">" & rs2("state") & "</option>")
            Loop Until rs2.Read() = 0
            Response.Write("</select></div>")
            Response.Write("</div>")
        End If
        EQ2(rs2)
    End Sub

    Sub RenderSourceCode(rs)
        Dim sql2, rs2
        Dim subValue = ""

        ' Check to see if the drop down has elements

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select IsNull(case  when c.id>0 then 1 else 0 end,0) as checked, IsNull(c.alternate_text,'') as value, a.include_text_field,a.referral_id,case a.alternate_text when '' then b.referral else a.alternate_text end as label from application_flatbridge_referral a left outer join application_source_code c on c.referral_id=a.referral_id and c.application_question_id=a.application_question_id and c.application_submission_id=" & CInt(Session("application_submission_id")) & ", student_referral b where b.id=a.referral_id and a.application_question_id=" & rs("id") & " order by a.display_order,label"

        Else
            sql2 = "select '' as value,0 as checked, a.include_text_field,a.referral_id,case a.alternate_text when '' then b.referral else a.alternate_text end as label from application_flatbridge_referral a, student_referral b where b.id=a.referral_id and a.application_question_id=" & rs("id") & " order by a.display_order,label"

        End If




        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""section more-steps""><div class=""item more-steps"">" & vbCrLF & "<h4>" & rs("label") & "</h4>" & vbCrLf & "<div class=""item-controls"">" & vbCrLf & "<table>")

            Do
                Response.Write(vbCrLf & "<tr><td style=""width:550px;""><div class=""item-controls-col2-1"" ><input type=""checkbox""  value=""" & rs2("referral_id") & """ name=""q" & rs("id") & """")
                If rs2("checked") = 1 Then Response.Write(" checked=""checked"" ")
                Response.Write("/>")

                Response.Write("&nbsp;<label style=""font-weight:bold;"">" & Server.HTMLEncode(rs2("label")) & "&nbsp;</label></div></td>")
                If rs2("include_text_field") = 1 Then
                    Response.Write("<td><div class=""item-controls-col2-2""><input type=""text"" name=""q_source_code" & rs2("referral_id"))
                    Response.Write(""" value=""" & rs2("value") & " """)
                    Response.Write("/></div></td>")
                End If
                Response.Write("</tr>")

            Loop Until rs2.Read() = 0

            Response.Write(vbCrLf & "</table>")
            Response.Write(vbCrLf & "</div></div></div>")

        End If
        EQ2(rs2)
    End Sub


    Sub RenderIndustry(rs)
        Dim sql2, rs2
        Dim subValue = ""

        ' Check to see if the drop down has elements

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select IsNull(case  when c.id>0 then 1 else 0 end,0) as checked, IsNull(c.alternate_text,'') as value, a.include_text_field,a.industry_id,case a.alternate_text when '' then b.industry else a.alternate_text end as label from application_flatbridge_industry a left outer join application_industry c on c.industry_id=a.industry_id and c.application_question_id=a.application_question_id and c.application_submission_id=" & CInt(Session("application_submission_id")) & ", industry b where b.id=a.industry_id and a.application_question_id=" & rs("id") & " order by a.display_order,label"

        Else
            sql2 = "select '' as value,0 as checked, a.include_text_field,a.industry_id,case a.alternate_text when '' then b.industry else a.alternate_text end as label from application_flatbridge_industry a, industry b where b.id=a.industry_id and a.application_question_id=" & rs("id") & " order by a.display_order,label"

        End If

        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""section more-steps""><div class=""item more-steps"">" & vbCrLF & "<div class=""item-label"" >" & rs("label") & "</div>" & vbCrLf & "<div class=""item-controls"">" & vbCrLf & "<table>")
            Dim ICount = 0
            Do
                If ICount Mod 3 = 0 Then
                    Response.Write("<tr>")
                End If

                If rs2("include_text_field") = 1 Then
                    If ICount Mod 3 <> 0 Then
                        Response.Write("</tr><tr>")
                    End If
                End If

                Response.Write("<td style=""width:240px;""><input type=""radio""  value=""" & rs2("industry_id") & """ name=""q" & rs("id") & """")
                If rs2("checked") = 1 Then Response.Write(" checked=""checked"" ")
                Response.Write("/>")


                Response.Write("&nbsp;<label style=""font-weight:bold;"">" & Server.HTMLEncode(rs2("label")) & "</label>")
                If rs2("include_text_field") = 1 Then
                    Response.Write("<input type=""text"" style=""width:190px;"" name=""q_industry" & rs2("industry_id"))
                    Response.Write(""" value=""" & rs2("value") & " """)
                    Response.Write("/>")
                End If
                Response.Write("</td><td width=""40"" nowrap=""nowrap"">&nbsp;</td>")

                ICount = ICount + 1

                If ICount Mod 3 = 0 Then
                    Response.Write("</tr>")
                End If

            Loop Until rs2.Read() = 0

            If ICount Mod 3 <> 0 Then
                Response.Write("</tr>")
            End If

            Response.Write("</table></div></div></div>")

        End If
        EQ2(rs2)
    End Sub


    Sub RenderManagementFunction(rs)
        Dim sql2, rs2
        Dim subValue = ""

        ' Check to see if the drop down has elements

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select IsNull(case  when c.id>0 then 1 else 0 end,0) as checked, IsNull(c.alternate_text,'') as value, a.include_text_field,a.management_function_id,case a.alternate_text when '' then b.management_function else a.alternate_text end as label from application_flatbridge_management_function a left outer join application_management_function c on c.management_function_id=a.management_function_id and c.application_question_id=a.application_question_id and c.application_submission_id=" & CInt(Session("application_submission_id")) & ", management_function b where b.id=a.management_function_id and a.application_question_id=" & rs("id") & " order by a.display_order,label"

        Else
            sql2 = "select '' as value,0 as checked, a.include_text_field,a.management_function_id,case a.alternate_text when '' then b.management_function else a.alternate_text end as label from application_flatbridge_management_function a, management_function b where b.id=a.management_function_id and a.application_question_id=" & rs("id") & " order by a.display_order,label"

        End If

        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""section more-steps""><div class=""item more-steps"">" & vbCrLF & "<div class=""item-label"">" & rs("label") & "<br/></div>" & vbCrLf & "<div class=""item-controls"">" & vbCrLf)

            Response.Write("<table>")
            Dim ICount = 0

            Do
                If ICount Mod 2 = 0 Then
                    Response.Write("<tr>")
                End If
                If rs2("include_text_field") = 1 Then
                    If ICount Mod 2 = 1 Then
                        Response.Write("</tr><tr>")
                    End If
                End If
                Response.Write("<td style=""vertical-align:top;""><input type=""radio""  value=""" & rs2("management_function_id") & """ name=""q" & rs("id") & """")
                If rs2("checked") = 1 Then Response.Write(" checked=""checked"" ")
                Response.Write("/>")

                Response.Write("&nbsp;<label style=""font-weight:bold;"">" & Server.HTMLEncode(rs2("label")) & "</label>")
                If rs2("include_text_field") = 1 Then
                    Response.Write("<input type=""text"" name=""q_management_function" & rs2("management_function_id"))
                    Response.Write(""" value=""" & rs2("value") & " """)
                    Response.Write("/>")
                End If
                Response.Write("</td>")
                ICount = Icount + 1

                If ICount Mod 2 = 0 Then
                    Response.Write("</tr>")
                End If

            Loop Until rs2.Read() = 0

            If ICount Mod 2 = 1 Then
                Response.Write("</tr>")
            End If

            Response.Write("</table>")

            Response.Write("</div></div></div>")

        End If
        EQ2(rs2)
    End Sub


    Sub RenderCountry(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        ' Check to see if the drop down has elements
        sql2 = "select id,country from country order by case when id=242 then 1 else 2 end, case when id=224 then 1 else 2 end,country"
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div>")
            Response.Write("<div class=""item-controls""><select name=""q" & rs("id") & """>")
            Response.Write("<option selected=""selected"" class=""group-title"" value=""null"">Select " & rs("label") & "</option>") 'Adding label in default option for Revamp
            Do
                Response.Write("<option value=""" & rs2("id") & """")
                If CStr(subValue) = CStr(rs2("id")) Then Response.Write(" selected=""selected"" ")
                Response.Write(">" & rs2("country") & "</option>")
            Loop Until rs2.Read() = 0
            Response.Write("</select></div>")
            Response.Write("</div>")
        End If
        EQ2(rs2)
    End Sub

    Sub RenderSalaryRange(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        ' Check to see if the drop down has elements
        sql2 = "select id,web_field_description from salary_web_range where display_order>1 order by display_order"
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div>")
            Response.Write("<div class=""item-controls""><select name=""q" & rs("id") & """>")
            Response.Write("<option selected=""selected"" class=""group-title"" value=""1""> PLEASE SELECT</option>") 
            Do
                Response.Write("<option value=""" & rs2("id") & """")
                If CStr(subValue) = CStr(rs2("id")) Then Response.Write(" selected=""selected"" ")
                Response.Write(">" & rs2("web_field_description") & "</option>")
            Loop Until rs2.Read() = 0
            Response.Write("</select></div>")
            Response.Write("</div>")
        End If
        EQ2(rs2)
    End Sub



    Sub RenderSalutation(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        ' Check to see if the drop down has elements
        sql2 = "select id,salutation from salutation order by salutation"
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label"))
            If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")

            Response.Write("</div>")
            Response.Write("<div class=""item-controls""><select id=""q" & rs("id") & """  name=""q" & rs("id") & """ ")
            If rs("required") = True Then Response.Write("class=""validate[required] required"" ")
            Response.Write(">")
            Response.Write("<option  class=""group-title"" value="""">Select " & rs("label") & "</option>") 'Adding label in default option for Revamp

            Do
                Response.Write("<option value=""" & rs2("id") & """")
                If CStr(subValue) = CStr(rs2("id")) Then Response.Write(" selected=""selected"" ")
                Response.Write(">" & rs2("salutation") & "</option>")
            Loop Until rs2.Read() = 0
            Response.Write("</select></div>")
            Response.Write("</div>")
        End If
        EQ2(rs2)
    End Sub


    Sub RenderGender(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If
        ' Check to see if the drop down has elements
        sql2 = "select id,salutation from salutation order by salutation"
        rs2 = BQ2(sql2)
        If rs2.Read() Then


            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label"))
            If rs("required") = True Then Response.Write("<span class=""req-asterisk"">*</span>")

            Response.Write("</div>")
            Response.Write("<div class=""item-controls""><select id=""q" & rs("id") & """ name=""q" & rs("id") & """")
            If rs("required") = True Then Response.Write("class=""validate[required] required"" ")

            Response.Write(">")
            Response.Write("<option class=""group-title"" value="""">Select " & rs("label") & "</option>") 'Adding label in default option for Revamp

            Response.Write("<option value=""F"" ")
            If CStr(subValue) = "F" Then Response.Write(" selected=""selected"" ")
            Response.Write(">Female</option>")
            Response.Write("<option value=""M"" ")
            If CStr(subValue) = "M" Then Response.Write(" selected=""selected"" ")
            Response.Write(">Male</option>")
            Response.Write("<option value=""G"" ")
            If CStr(subValue) = "G" Then Response.Write(" selected=""selected"" ")
            Response.Write(">Gender nonconforming</option>")
            Response.Write("<option value=""P"" ")
            If CStr(subValue) = "P" Then Response.Write(" selected=""selected"" ")
            Response.Write(">Prefer not to answer</option>")

            Response.Write("</select></div>")
            Response.Write("</div>")
        End If
        EQ2(rs2)





    End Sub



    Sub RenderRadio(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If

        ' Check to see if we should double up
        Dim Count = 0
        sql2 = "select count(*) as count from application_content_choice where application_question_id=" & rs("id")
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Count = rs2("count")
        End If
        EQ2(rs2)


        ' Check to see if the drop down has elements
        sql2 = "select id,item from application_content_choice where application_question_id=" & rs("id") & " order by display_order,item"
        rs2 = BQ2(sql2)


        If rs2.Read() Then

            Response.Write("<div class=""item""><table><tr><td style=""vertical-align:top;""><div class=""item-label"">" & rs("label") & "</div></td><td style=""vertical-align:top;"">")
            Response.Write("<div class=""item-controls"">")
            Do
                Response.Write("<input type=""radio"" name=""q" & rs("id") & """ value=""" & rs2("id") & """")
                Try
                    If subValue = rs2("id") Then Response.Write(" checked=""checked"" ")
                Catch
                End Try
                Response.Write("/>&nbsp;<label style=""font-weight:bold;"">" & rs2("item") & "</label><br/>")
            Loop Until rs2.Read() = 0
            Response.Write("</div></td></tr></table>")
            Response.Write("</div>")

        End If
        EQ2(rs2)
    End Sub

    Sub RenderRadioMoney(rs)
        Dim sql2, rs2
        Dim subValue = ""
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
            End If
            EQ2(rs2)
        End If

        ' Check to see if we should double up
        Dim Count = 0
        sql2 = "select count(*) as count from application_content_choice_money where application_question_id=" & rs("id")
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Count = rs2("count")
        End If
        EQ2(rs2)


        ' Check to see if the drop down has elements
        sql2 = "select id,item from application_content_choice_money where application_question_id=" & rs("id") & " order by display_order,item"
        rs2 = BQ2(sql2)


        If rs2.Read() Then

            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div>")
            Response.Write("<div class=""item-controls""><ul>")

            Do

                Response.Write("<li><input type=""radio"" name=""q" & rs("id") & """ value=""" & rs2("id") & """")
                Try
                    If subValue = rs2("id") Then Response.Write(" checked=""checked"" ")
                Catch
                End Try
                Response.Write("/><label>" & rs2("item") & "</label></li>")

            Loop Until rs2.Read() = 0
            Response.Write("</ul></div>")
            Response.Write("</div>")


        End If
        EQ2(rs2)
    End Sub


    Sub RenderDropDownMultipleSelect(rs)
        Dim sql2, rs2
        Dim subValue
        Dim selArray() As String
        Dim I
        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
                selArray = split(subValue, ",")
            End If
            EQ2(rs2)
        End If
        ' Check to see if the drop down has elements
        sql2 = "select id,item from application_content_choice where application_question_id=" & rs("id") & " order by display_order,item"
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div>")
            Response.Write("<div class=""item-controls""><select name=""q" & rs("id") & """ multiple=""multiple"">")
            Do
                Response.Write("<option value=""" & rs2("id") & """")
                Try ' If they haven't selected any just skip
                    For I = 0 To selArray.Length - 1
                        If selArray(I) = rs2("id") Then Response.Write(" selected=""selected"" ")
                    Next
                Catch
                End Try
                Response.Write(">" & rs2("item") & "</option>")
            Loop Until rs2.Read() = 0
            Response.Write("</select></div>")
            Response.Write("</div>")
        End If
        EQ2(rs2)
    End Sub

    Sub RenderCheckBox(rs)
        Dim sql2, rs2
        Dim subValue
        Dim selArray() As String

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
                selArray = split(subValue, ",")
            End If
            EQ2(rs2)
        End If
        Dim Count = 0
        sql2 = "select count(*) as count from application_content_choice where application_question_id=" & rs("id")
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Count = rs2("count")
        End If
        EQ2(rs2)

        ' Check to see if the drop down has elements
        sql2 = "select id,item from application_content_choice where application_question_id=" & rs("id") & " order by display_order,item"
        rs2 = BQ2(sql2)
        If rs2.Read() Then

            Response.Write("<div class=""item"">")
            If rs("label") <> "<br>" Then
                Response.Write("<div class=""item-label"">" & rs("label") & "</div>")
            End If

            Response.Write("<div class=""item-controls""><table>")

            Do
                Response.Write("<tr><td><div class=""item-controls-col2-2""><input type=""checkbox"" name=""q" & rs("id") & """ value=""" & rs2("id") & """")
                Try
                    For Each thisObject As String In selArray
                        If (rs2("id") = thisObject) Then Response.Write(" checked=""checked"" ")
                    Next thisObject
                Catch
                End Try
                Response.Write("/>&nbsp;<label style=""font-weight:bold;"">" & rs2("item") & "</label><br/><br/></div></td></tr>")
            Loop Until rs2.Read() = 0
            Response.Write("</table>")
            Response.Write("</div></div>")

        End If
        EQ2(rs2)
    End Sub

    Sub RenderCheckBoxMoney(rs)
        Dim sql2, rs2
        Dim subValue
        Dim selArray() As String

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select value from Application_submission_values where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                subValue = rs2("value")
                selArray = split(subValue, ",")
            End If
            EQ2(rs2)
        End If
        Dim Count = 0
        sql2 = "select count(*) as count from application_content_choice_money where application_question_id=" & rs("id")
        rs2 = BQ2(sql2)
        If rs2.Read() Then
            Count = rs2("count")
        End If
        EQ2(rs2)

        ' Check to see if the drop down has elements
        sql2 = "select id,item from application_content_choice_money where application_question_id=" & rs("id") & " order by display_order,item"
        rs2 = BQ2(sql2)
        If rs2.Read() Then

            Response.Write("<div class=""item""><div class=""item-label"">" & rs("label") & "</div><div class=""item-controls""><ul>")

            Do

                Response.Write("<li><input type=""checkbox"" name=""q" & rs("id") & """ value=""" & rs2("id") & """")
                Try
                    For Each thisObject As String In selArray
                        If (rs2("id") = thisObject) Then Response.Write(" checked=""checked"" ")
                    Next thisObject

                Catch
                End Try
                Response.Write("/><label>" & rs2("item") & "</label></li>")

            Loop Until rs2.Read() = 0

            Response.Write("</ul>")
            Response.Write("</div></div>")

        End If
        EQ2(rs2)
    End Sub





    Sub RenderSubmitButton(course_id As Integer)

        If CStr(Session("application_user_id")) <> "" Then
            If CheckCourse(course_id) Then
                Response.Write("<div class=""divider"">&nbsp;</div><div id=""btn-forms-wrap""><ul>")

                If Request.QueryString("d") <> 1 Then
                    Response.Write("<li><a href=""javascript:void(0);"" id=""btn-back""><img src=""images/btn-back.png"" width=""92"" height=""35"" alt=""Back"" /></a></li><li>&nbsp;&nbsp;</li>")
                End If


                Response.Write("<li><a href=""#"" id=""btn-save""><img src=""images/btn-save.png""  width=""92"" height=""35"" alt=""Save"" /></a></li><li>&nbsp;&nbsp;</li>")

                Response.Write("<li><a href=""javascript:void(0);"" id=""btn-next""><img src=""images/btn-next.png"" width=""95"" height=""35"" alt=""Next"" /></a></li></ul></div>")
            Else

                Response.Write("<div class=""divider"">&nbsp;</div><div id=""btn-forms-wrap""><ul><li><a href=""javascript:void(0);"" id=""btn-next""><img src=""images/btn-next.png"" width=""95"" height=""35"" alt=""Next"" /></a></li></ul></div>")

            End If
        End If

    End Sub


    Sub RenderEventSubmitButton(course_id As Integer)
        If CStr(Session("application_user_id")) <> "" Then
            If CheckCourse(course_id) Then
                Response.Write("<div class=""divider"">&nbsp;</div><div id=""btn-forms-wrap""><ul>")


                Response.Write("<li><a href=""javascript:void(0);"" id=""btn-submit""><img src=""images/btn-pf-submit.png"" width=""95"" height=""35"" alt=""Submit"" /></a></li></ul></div>")
            Else

                Response.Write("<div class=""divider"">&nbsp;</div><div id=""btn-forms-wrap""><ul><li><a href=""javascript:void(0);"" id=""btn-submit""><img src=""images/btn-pf-submit.png"" width=""95"" height=""35"" alt=""Submit"" /></a></li></ul></div>")
            End If

        End If
    End Sub

    Sub RenderContinueButton()
        Response.Write("<tr><td colspan=""2"" align=""center"">")
        If CStr(Session("application_user_id")) <> "" Then
            Response.Write("&nbsp;&nbsp;<input type=""submit"" name=""continue"" value=""Continue"" />")
        End If
        Response.Write("</td></tr>")
    End Sub




    Sub RenderEmploymentRecord(rs, ByRef HeaderSeen)
        If headerSeen = True Then
            Response.Write("</div>")
            headerSeen = False
        End If
        Response.Write("<div class=""divider"">&nbsp;</div><div class=""section-title""><h3>" & rs("label") & "</h3>")
        Response.Write("</div>")

        Response.Write("<div class=""form-section-wrap"">")
        Dim sql2, rs2
        If rs("explanatory_text") <> "" Then
            Response.Write("<div class=""body-emph""><p>" & rs("explanatory_text") & "</p></div>")
        End If

        Response.Write("<div class=""body-emph""><p>List the past 3 positions you have held beginning with the most recent. Treat different assignments in the same firm as separate positions.</p></div>")

        Dim company(4), position(4), start_date(4), end_date(4)

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select * from application_employment_record  where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                Do
                    position(rs2("entry_num") - 1) = rs2("position")
                    company(rs2("entry_num") - 1) = rs2("company")
                    start_date(rs2("entry_num") - 1) = rs2("start_date")
                    end_date(rs2("entry_num") - 1) = rs2("end_date")

                Loop Until rs2.Read() = 0
            End If
            EQ2(rs2)
        End If
        Dim J
        J = 1

        Response.Write("<div class=""section-row"" id=""pr-educ-entry1""><ul>")
        Response.Write("<li><div class=""section-header"">Name of Company / Organization<br/>&nbsp;</div><input type=""text""  class=""frm-college"" name=""q_company_" & J & "_" & rs("id") & """ value=""" & company(J - 1) & """/></li><li><div class=""section-header"">Position<br/>&nbsp;</div><input type=""text""  class=""frm-clmn-st2"" name=""q_position_" & J & "_" & rs("id") & """ value=""" & position(J - 1) & """/></li><li><div class=""section-header"">Start Date<br /><span class=""body-emph"">(mm/dd/yyyy)</span></div><input type=""text""  class=""validate[custom[dateFormat]] frm-clmn-st2"" id=""q_start_date_" & J & "_" & rs("id") & """ name=""q_start_date_" & J & "_" & rs("id") & """ value=""" & start_date(J - 1) & """/></li><li><div class=""section-header"">End Date<br /><span class=""body-emph"">(mm/dd/yyyy)</span></div><input type=""text""  class=""validate[custom[dateFormat]] frm-clmn-st2"" name=""q_end_date_" & J & "_" & rs("id") & """  id=""q_end_date_" & J & "_" & rs("id") & """ value=""" & end_date(J - 1) & """/></li>")
        Response.Write("</ul></div>")




        For J = 2 To 3
            Response.Write("<div class=""section-row"" id=""pr-educ-entry""" & J & "><ul>")
            Response.Write("<li><input type=""text""  class=""frm-college"" name=""q_company_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & company(J - 1) & """")
            Response.Write("/></li><li><input type=""text""  class=""frm-clmn-st2"" name=""q_position_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & position(J - 1) & """")
            Response.Write("/></li><li><input type=""text""   class=""validate[custom[dateFormat]] frm-clmn-st2"" id=""q_start_date_" & J & "_" & rs("id") & """  name=""q_start_date_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & start_date(J - 1) & """")

            Response.Write("/></li><li><input type=""text""   class=""validate[custom[dateFormat]] frm-clmn-st2"" id=""q_end_date_" & J & "_" & rs("id") & """ name=""q_end_date_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & end_date(J - 1) & """")

            Response.Write("/></li></ul></div>")

        Next
        Response.Write("</div")



    End Sub

    Sub RenderContact()
        Response.Write("<table width=""800"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#FFFFFF"">")
        Response.Write("  <tr>")
        Response.Write("    <td width=""170"" align=""left"" valign=""top"" bgcolor=""#e9e3b1""><br /></td><td width=""630"" valign=""baseline"">")
        Response.Write("        <table width=""620"" border=""0"" align=""right"" cellpadding=""5"" cellspacing=""0"" bgcolor=""#e9e3b1"">")
        Response.Write("          <tr>")
        Response.Write("            <td width=""66""><a href=""http://www.gsb.stanford.edu""><img src=""https://www.gsb.stanford.edu/exed/images/SU_Seal.gif"" alt=""SU Seal"" width=""51"" height=""51"" border=""0"" align=""left"" /></a></td>")
        Response.Write("            <td width=""247""><span class=""contact"">Office of Executive Education <br />")
        Response.Write("Stanford Graduate School of Business<br />")
        Response.Write("518 Memorial Way <br />")
        Response.Write("            Stanford, CA 94305-5015</span></td>")
        Response.Write("            <td width=""277"" class=""contact"">Phone: 650.723.3341<br />")
        Response.Write("              Toll Free: 866.542.2205  (US and Canada Only) <br />")
        Response.Write("              Fax:   650.723.3950 <br />")
        Response.Write("            Email: <a href=""mailto:executive_education@gsb.stanford.edu"">executive_education@gsb.stanford.edu</a></td>")
        Response.Write("          </tr>")
        Response.Write("        </table>")
        Response.Write("")
        Response.Write("    </td>")
        Response.Write("  </tr>")
        Response.Write("</table><!--  10 PX WHITE SPACER TABLE -->")
        Response.Write("<table width=""800"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
        Response.Write("  <tr>")
        Response.Write("    <td><img src=""https://www.gsb.stanford.edu/exed/images/spacer.gif"" width=""800"" height=""10"" border=""0"" align=""top"" alt=""spacer"" /></td>")
        Response.Write("  </tr>")
        Response.Write("</table><!-- 10 PX WHITE SPACER TABLE -->")
        Response.Write("<!-- FOOTER TABLE --><table width=""800"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""footer"">")
        Response.Write("  <tr>")
        Response.Write("    <td><a href=""http://www.gsb.stanford.edu/policy/terms.html"">TERMS OF USE</a><b> :: </b><a href=""http://www.gsb.stanford.edu/policy/privacy.html"">PRIVACY POLICY</a><strong> :: </strong><a href=""http://www.gsb.stanford.edu/help.html"">HELP</a></td>")
        Response.Write("	<td align=""right""> &copy; " & Now().Year & " Stanford University Graduate School of Business</td>")
        Response.Write("  </tr>")
        Response.Write("</table>")
        Response.Write("<!-- END FOOTER TABLE -->	</td>")
        Response.Write("  </tr>")
        Response.Write("</table>")


    End Sub



    Sub RenderProfessionalEducation(rs, ByRef headerSeen)
        If headerSeen = True Then
            Response.Write("</div>")
            headerSeen = False
        End If
        Response.Write("<div class=""divider"">&nbsp;</div><div class=""section-title""><h3>" & rs("label") & "</h3>")
        Response.Write("</div>")
        If rs("explanatory_text") <> "" Then
            'Response.Write("<tr><td colspan=""2"">" & rs("explanatory_text") & "</td></tr>")
        End If
        Dim sql2, rs2
        Dim school(4), start_date(4), end_date(4)


        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select * from application_professional_education  where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                Do
                    school(rs2("entry_num") - 1) = rs2("school")
                    start_date(rs2("entry_num") - 1) = rs2("start_date")
                    end_date(rs2("entry_num") - 1) = rs2("end_date")

                Loop Until rs2.Read() = 0
            End If
            EQ2(rs2)
        End If
        Dim J
        J = 1
        Response.Write("<div class=""form-section-wrap"">")

        Response.Write("<div class=""section-row"" id=""pr-educ-entry1"">")
        Response.Write("<ul><li><div class=""section-header"">School / Program</div><br/><input type=""text"" size=""45"" name=""q_school_" & J & "_" & rs("id") & """  value=""" & school(J - 1) & """/></li><li><div class=""section-header"">Start Date<br/><span class=""body-emph"">(mm/dd/yyyy)</span></div><input type=""text"" class=""validate[custom[dateFormat]]"" size=""20"" name=""q_start_date_" & J & "_" & rs("id") & """ id=""q_start_date_" & J & "_" & rs("id") & """ value=""" & start_date(J - 1) & """/></li><li><div class=""section-header"">End Date<br/><span class=""body-emph"">(mm/dd/yyyy)</span></div><input type=""text"" class=""validate[custom[dateFormat]]""  size=""20"" name=""q_end_date_" & J & "_" & rs("id") & """ id=""q_end_date_" & J & "_" & rs("id") & """   value=""" & end_date(J - 1) & """/></li></ul></div>") 'Changes made for Revamp



        For J = 2 To 4
            Response.Write("<div class=""section-row"" id=""pr-educ-entry" & J & """><ul><li><input type=""text"" size=""45"" name=""q_school_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & school(J - 1) & """")
            Response.Write("/></li><li><input type=""text"" class=""validate[custom[dateFormat]]"" size=""20"" name=""q_start_date_" & J & "_" & rs("id") & """ id=""q_start_date_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & start_date(J - 1) & """")
            Response.Write("/></li><li><input type=""text"" class=""validate[custom[dateFormat]]""  size=""20"" name=""q_end_date_" & J & "_" & rs("id") & """ id=""q_end_date_" & J & "_" & rs("id") & """")

            Response.Write(" value=""" & end_date(J - 1) & """")

            Response.Write("/></li></ul></div>")

        Next
        Response.Write("</div>")

        'Response.Write("</div>")

    End Sub


    Sub RenderEducation(rs, ByRef HeaderSeen)
        If headerSeen = True Then
            Response.Write("</div>")
            headerSeen = False
        End If
        Dim sql2, rs2
        Response.Write("<div class=""divider"">&nbsp;</div><div class=""section-title""><h3>" & rs("label") & "</h3>")
        Response.Write("</div>")

        Response.Write("<div class=""form-section-wrap"">")


        Dim year_granted(3), college(3), major(3), application_degree_id(3)

        If CStr(Session("application_submission_id")) <> "" Then  ' Need to lookup value
            sql2 = "select * from application_education  where application_submission_id=" & CInt(Session("application_submission_id")) & " and application_question_id=" & rs("id")
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                Do
                    year_granted(rs2("entry_num") - 1) = rs2("year_granted")
                    college(rs2("entry_num") - 1) = rs2("college")
                    major(rs2("entry_num") - 1) = rs2("major")
                    application_degree_id(rs2("entry_num") - 1) = rs2("application_degree_id")

                Loop Until rs2.Read() = 0
            End If
            EQ2(rs2)
        End If

        Dim J
        For J = 1 To 3
            Response.Write("<div class=""section-row"" id=""educ-entry" & J & """>" & vbCrLf & "<ul>" & vbCrLf & "<li>")
            If J = 1 Then
                Response.Write(vbCrLf & "<div class=""section-header"">College / University</div>")
            End If
            Response.Write("<input type=""text"" size=""25"" maxlength=""200"" class=""frm-college"" name=""q_college_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & college(J - 1) & """")
            Response.Write("/></li>" & vbCrLf & "<li>")
            If J = 1 Then
                Response.Write(vbCrLf & "<div class=""section-header"">Year Granted</div>")
            End If
            Response.Write("<input type=""text"" size=""20"" maxlength=""50"" class=""frm-clmn-st2"" name=""q_year_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & year_granted(J - 1) & """")
            Response.Write("/></li>" & vbCrLf & "<li>")



            If J = 1 Then
                Response.Write(vbCrLf & "<div class=""section-header"">Degree Granted</div>")
            End If


            sql2 = "select id,degree from Application_Degree order by display_order"
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                Response.Write("<select name=""q_degree_" & J & "_" & rs("id") & """>")
                Do
                    Response.Write("<option value=""" & rs2("id") & """")
                    If rs2("id") = application_degree_id(J - 1) Then Response.Write(" selected=""selected"" ")
                    Response.Write(">" & rs2("degree") & "</option>" & vbCrLf)
                Loop Until rs2.Read() = 0
                Response.Write("</select></li>" & vbCrLf)
            End If
            EQ2(rs2)

            Response.Write("<li>" & vbCrLf)
            If J = 1 Then
                Response.Write(vbCrLf & "<div class=""section-header"">Major</div>")
            End If

            Response.Write("<input type=""text"" size=""20"" maxlength=""200"" class=""frm-clmn-st2"" name=""q_major_" & J & "_" & rs("id") & """")
            Response.Write(" value=""" & major(J - 1) & """")
            Response.Write("/></li>" & vbCrLf & "</ul></div>")

        Next
        Response.Write("</div>")
        'Response.Write("</div>")


    End Sub

    Sub ProcessContinue()


        ' Lookup page to see if it is last
        Dim sql, rs
        Dim application_id
        Dim pageFrom = CInt(Request.QueryString("p"))
        Dim pageTo
        Dim display_order = 0

        sql = "select IsNull(application_id,0) as application_id from Course where id=" & CInt(Request.QueryString("id"))
        rs = BQ(sql)
        If rs.Read() Then
            application_id = rs("application_id")
        Else
            Response.Write("<h1>An application has not been assigned to this program</h1>")
            Response.End()
        End If
        EQ(rs)

        sql = "select display_order from application_page_assigned where application_id=" & application_id & " and application_page_id=" & pageFrom
        rs = BQ(sql)
        If rs.Read() Then
            display_order = rs("display_order")
        Else
            display_order = 0
        End If
        EQ(rs)


        sql = "select top 1 application_page_id as pageTo, display_order from application_page_assigned where application_id=" & application_id & " and active=1 and display_order>" & display_order & " order by display_order"
        rs = BQ(sql)
        If rs.Read() Then
            Dim disp_order = rs("display_order")
            pageTo = rs("pageTo")
            EQ(rs)
            Response.Redirect("../app/apply.aspx?id=" & CInt(Request.QueryString("id")) & "&p=" & pageTo & "&d=" & disp_order)
            Response.End()
        Else
            Response.Redirect("../app/submit.aspx?id=" & CInt(Request.QueryString("id")))
        End If
        EQ(rs)
    End Sub

    Sub ProcessSubmission(Mode)
        AppCapture()
        Dim sql, rs, sql2, rs2
        Dim application_submission_id
        Dim application_user_id = Session("application_user_id")
        LogError("Application user id " & application_user_id)

        Dim application_status
        Try

            Select Case Mode
                Case "Submit"
                    application_status = 1
                Case "Save"
                    application_status = 4
            End Select

            sql = "ApplicationSubmitMultiple " & CInt(Request.Form("course_id")) & "," & CInt(application_status) & "," & CInt(application_user_id)
            LogError(sql)
            rs = BQ(sql)
            If rs.Read() Then
                application_submission_id = rs("id")
                LogError("Application Submission ID " & application_submission_id)
            End If
            EQ(rs)

            sql = "delete from Application_Page_Submission where application_page_id=" & CInt(Request.Form("page_id")) & " and application_submission_id=" & CInt(application_submission_id)
            rs = BQ(sql)
            EQ(rs)


            sql = "insert into Application_Page_Submission values(" & CInt(Request.Form("page_id")) & "," & CInt(application_submission_id) & ",getDate())"
            rs = BQ(sql)
            EQ(rs)


            sql = "ApplicationGetQuestions " & CInt(Request.Form("application_id")) & "," & CInt(Request.Form("page_id"))


            rs = BQ(sql)
            If rs.Read() Then
                Do
                    Try
                        Select Case rs("type")
                            Case "File"
                                ' Lets Get Posted File

                                Try
                                    Dim uploads As HttpFileCollection
                                    uploads = HttpContext.Current.Request.Files  ' Need to get them all
                                    Dim postedFile = uploads(CStr("q" & rs("id")))  'Find the one that matches this control
                                    Dim filename As String = Path.GetFileName(postedFile.FileName)
                                    Dim contentType As String = postedFile.ContentType
                                    Dim contentLength As Integer
                                    contentLength = postedFile.ContentLength
                                    Dim strConnect = ConfigurationManager.connectionStrings("pubsConnectionString").ConnectionString
                                    Dim cn = New System.Data.SqlClient.SqlConnection(strConnect)
                                    postedFile.InputStream.Position = 0
                                    Dim arrB(postedFile.InputStream.length) As Byte
                                    postedFile.InputStream.read(arrB, 0, postedFile.InputStream.length)
                                    If contentLength > 0 Then
                                        sql2 = "delete from Application_Submission_File where application_submission_id=" & CInt(application_submission_id) & " and application_question_id=" & CInt(rs("id"))
                                        rs2 = BQ2(sql2)
                                        EQ2(rs2)
                                        Dim cmd As New System.Data.SqlClient.SqlCommand("Insert into Application_Submission_File values (" & CInt(application_submission_id) & "," & rs("id") & ",@AttachedFile,'" & db(fileName) & "','" & db(contentType) & "')", cn)
                                        Dim P As New System.Data.SqlClient.SqlParameter("@AttachedFile", System.Data.SqlDbType.Image, arrB.Length - 1, System.Data.ParameterDirection.Input, False, 0, 0, Nothing, System.Data.DataRowVersion.Current, arrB)
                                        cmd.Parameters.Add(P)
                                        cn.Open()
                                        cmd.ExecuteNonQuery()
                                        cn.Close()
                                    Else
                                        Try
                                            'Let's check if they submitted one already and use that
                                            'sql2="insert into Application_Submission_File select " & application_submission_id & ",application_question_id, [file], filename,content_type from Application_Submission_File where application_submission_id=" & application_submission_id & " and application_question_id=" & rs("id")

                                            'rs2=BQ2(sql2)
                                            'EQ2(rs2)		
                                        Catch
                                        End Try
                                    End If
                                Catch exc As Exception


                                End Try
                            Case "Employment Record"
                                Dim J
                                For J = 1 To 3

                                    sql2 = "delete from Application_Employment_Record where application_submission_id=" & CInt(application_submission_id) & " and application_question_id=" & CInt(rs("id")) & " and entry_num=" & J
                                    rs2 = BQ2(sql2)
                                    EQ2(rs2)

                                    If Request.Form("q_company_" & J & "_" & rs("id")) <> "" Then

                                        sql2 = "insert into Application_Employment_Record values('" & db(Request.Form("q_company_" & J & "_" & rs("id"))) & "','" & db(Request.Form("q_position_" & J & "_" & rs("id"))) & "','" & db(Request.Form("q_start_date_" & J & "_" & rs("id"))) & "','" & db(Request.Form("q_end_date_" & J & "_" & rs("id"))) & "'," & J & "," & CInt(application_submission_id) & "," & rs("id") & ")"
                                        rs2 = BQ2(sql2)
                                        EQ2(rs2)
                                    End If
                                Next
                            Case "Professional Education"
                                Dim J
                                For J = 1 To 4
                                    sql2 = "delete from Application_Professional_Education where application_submission_id=" & CInt(application_submission_id) & " and application_question_id=" & CInt(rs("id")) & " and entry_num=" & J
                                    rs2 = BQ2(sql2)
                                    EQ2(rs2)
                                    If Request.Form("q_school_" & J & "_" & rs("id")) <> "" Then
                                        sql2 = "insert into Application_Professional_Education values('" & db(Request.Form("q_school_" & J & "_" & rs("id"))) & "','" & db(Request.Form("q_start_date_" & J & "_" & rs("id"))) & "','" & db(Request.Form("q_end_date_" & J & "_" & rs("id"))) & "'," & J & "," & CInt(application_submission_id) & "," & rs("id") & ")"
                                        rs2 = BQ2(sql2)
                                        EQ2(rs2)
                                    End If
                                Next

                            Case "Education"
                                Dim J
                                For J = 1 To 3
                                    sql2 = "delete from Application_Education where application_submission_id=" & CInt(application_submission_id) & " and application_question_id=" & CInt(rs("id")) & " and entry_num=" & J
                                    rs2 = BQ2(sql2)
                                    EQ2(rs2)
                                    If Request.Form("q_college_" & J & "_" & rs("id")) <> "" Then
                                        sql2 = "insert into Application_Education values('" & db(Request.Form("q_college_" & J & "_" & rs("id"))) & "','" & db(Request.Form("q_year_" & J & "_" & rs("id"))) & "'," & CInt(Request.Form("q_degree_" & J & "_" & rs("id"))) & ",'" & db(Request.Form("q_major_" & J & "_" & rs("id"))) & "'," & J & "," & CInt(application_submission_id) & "," & rs("id") & ")"
                                        rs2 = BQ2(sql2)
                                        EQ2(rs2)
                                    End If
                                Next
                            Case "Source Code"
                                sql2 = "delete from Application_Source_Code where application_submission_id=" & CInt(application_submission_id) & " and application_question_id=" & CInt(rs("id"))
                                rs2 = BQ2(sql2)
                                EQ2(rs2)

                                If Request.Form("q" & rs("id")) <> "" Then
                                    Dim selArray() As String
                                    selArray = split(Request.Form("q" & rs("id")), ",")
                                    Dim J
                                    For J = 0 To selArray.Length - 1
                                        sql2 = "insert into Application_Source_Code values(" & CInt(selArray(J)) & ",'" & db(Request.Form("q_source_code" & selArray(J))) & "'," & CInt(application_submission_id) & "," & rs("id") & ")"
                                        rs2 = BQ2(sql2)
                                        EQ2(rs2)
                                    Next
                                End If
                            Case "Industry"
                                sql2 = "delete from Application_Industry where application_submission_id=" & CInt(application_submission_id) & " and application_question_id=" & CInt(rs("id"))
                                rs2 = BQ2(sql2)
                                EQ2(rs2)
                                If Request.Form("q" & rs("id")) <> "" Then
                                    Dim selArray() As String
                                    selArray = split(Request.Form("q" & rs("id")), ",")
                                    Dim J
                                    For J = 0 To selArray.Length - 1
                                        sql2 = "insert into Application_Industry values(" & CInt(selArray(J)) & ",'" & db(Request.Form("q_industry" & selArray(J))) & "'," & CInt(application_submission_id) & "," & rs("id") & ")"
                                        rs2 = BQ2(sql2)
                                        EQ2(rs2)
                                    Next
                                End If
                            Case "Management Function"
                                sql2 = "delete from Application_Management_Function where application_submission_id=" & CInt(application_submission_id) & " and application_question_id=" & CInt(rs("id"))
                                rs2 = BQ2(sql2)
                                EQ2(rs2)
                                If Request.Form("q" & rs("id")) <> "" Then
                                    Dim selArray() As String
                                    selArray = split(Request.Form("q" & rs("id")), ",")
                                    Dim J
                                    For J = 0 To selArray.Length - 1
                                        sql2 = "insert into Application_Management_Function values(" & CInt(selArray(J)) & ",'" & db(Request.Form("q_management_function" & selArray(J))) & "'," & CInt(application_submission_id) & "," & rs("id") & ")"
                                        rs2 = BQ2(sql2)
                                        EQ2(rs2)
                                    Next
                                End If
                            Case "Email", "First Name", "Last Name", "Work Phone", "Fax", "Home Phone", "HR Work Phone", "HR Fax", "Emergency Contact Phone", "Cell Phone"
                                If Request.Form("q" & rs("id")) <> "" Then
                                    sql2 = "ApplicationSubmitValue " & CInt(application_submission_id) & "," & rs("id") & ",'" & db(Request.Form("q" & rs("id"))) & "'"
                                    Session(rs("type")) = Request.Form("q" & rs("id"))
                                    rs2 = BQ2(sql2)
                                    EQ2(rs2)
                                End If

                            Case "Date of Birth"
                                Dim month, day, year, sdate
                                month = Request.Form("q1-" & rs("id"))
                                day = Request.Form("q2-" & rs("id"))
                                year = Request.Form("q3-" & rs("id"))


                                If month <> "" And day <> "" And year <> "" Then
                                    sdate = month & "/" & day & "/" & year
                                Else
                                    sdate = ""
                                End If


                                sql2 = "ApplicationSubmitValue " & CInt(application_submission_id) & "," & rs("id") & ",'" & db(sdate) & "'"
                                rs2 = BQ2(sql2)
                                EQ2(rs2)

                            Case Else

                                sql2 = "ApplicationSubmitValue " & CInt(application_submission_id) & "," & rs("id") & ",'" & db(Request.Form("q" & rs("id"))) & "'"
                                rs2 = BQ2(sql2)

                                EQ2(rs2)

                        End Select
                    Catch ex As Exception
                        LogError(ex.ToString() & sql & ":" & sql2)
                    End Try
                Loop Until rs.Read() = 0

                If CStr(application_user_id) <> "" Then

                    sql2 = "delete from Application_User_App where application_submission_id=" & CInt(application_submission_id) & " and application_user_id=" & CInt(application_user_id)
                    rs2 = BQ2(sql2)
                    EQ2(rs2)


                    sql2 = "insert into Application_User_App values(" & CInt(application_submission_id) & "," & CInt(application_user_id) & ")"
                    rs2 = BQ2(sql2)
                    EQ2(rs2)
                End If


                Session("application_submission_id") = application_submission_id



                Select Case Mode
                    Case "Submit"
                        SendEmail()
                        Response.Redirect("../app/application_confirm.aspx?id=" & CInt(Request.Form("course_id")))

                    Case "Save Application"
                        Response.Redirect("../app/application_save_confirm.aspx?id=" & CInt(Request.Form("course_id")))
                End Select

            End If
            EQ(rs)
        Catch ex As Exception
            Response.Write(ex.ToString() & sql)
            LogError(ex.ToString() & sql & ":" & sql2)

        End Try
    End Sub


    Sub SendEmail()
        Dim email_title, description
        Dim sql, rs, sql2, rs2
        sql = "select email_title, description from SalesEmail a, Course b where b.id=" & CInt(Request.Form("course_id")) & " and b.application_email_script=a.id"
        rs = BQ(sql)
        If rs.Read() Then
            Try
                email_title = rs("email_title")
                description = rs("description")


                Dim myMail As New MailMessage()
                myMail.From = New MailAddress("executive_education@gsb.stanford.edu", "Stanford Graduate School of Business - Office of Executive Education")
                myMail.IsBodyHTML = True
                Dim toAddress As New MailAddress(Session("Email"))
                myMail.To.Add(toAddress)
                Try
                    sql2 = "select email from Course a, Staff b, Person c where c.id=b.person_id and b.id=a.application_staff_member_id and a.id=" & CInt(Request.Form("course_id"))
                    rs2 = BQ2(sql2)
                    If rs2.Read() Then
                        Dim staff_email = rs2("email")
                        If Trim(staff_email) <> "" Then
                            myMail.cc.Add(New MailAddress(staff_email))
                        End If
                    End If
                    EQ2(rs2)
                Catch
                End Try

                Try
                    sql2 = "select c.email from Staff a, Person c,  Staff_Contacts b where b.id=5 and a.person_id=c.id and b.staff_id=a.id"
                    rs2 = BQ2(sql2)
                    If rs2.Read() Then
                        Dim staff_email2 = rs2("email")
                        If Trim(staff_email2) <> "" Then
                            myMail.cc.Add(New MailAddress(staff_email2))
                        End If
                    End If
                    EQ2(rs2)
                Catch
                End Try



                myMail.Subject = email_title

                Dim Token
                Token = "!FirstName!"
                description = Replace(description, Token, Session("First Name"))
                Token = "!LastName!"
                description = Replace(description, Token, Session("Last Name"))

                myMail.Body = description
                Dim smtpSender As New SmtpClient(System.Configuration.ConfigurationManager.AppSettings("MailServer"))
                smtpSender.UseDefaultCredentials = True
                smtpSender.Port = 25

                smtpSender.Send(myMail)
            Catch ex As Exception

            End Try

        End If
        EQ(rs)
    End Sub



    Sub Page_Load_Application()
        Dim sql, rs, sql2, rs2
        Dim oExcep As Exception
        Dim course_id
        Dim application_id

        Try
            If CInt(Session("application_user_id")) = 0 Then
                Response.Redirect("../app/application_logon.aspx?id=" & CInt(Request.QueryString("id")))
                Response.End()
            End If
        Catch
            Response.Redirect("../app/application_logon.aspx?id=" & CInt(Request.QueryString("id")))
            Response.End()

        End Try

        Try
            Try
                course_id = CInt(Request.QueryString("id"))
            Catch
                oExcep = New Exception("An attempt has been made to render this page with an invalid program id.")
                Throw oExcep
            End Try
            If course_id = 0 Then
                oExcep = New Exception("An attempt has been made to render this page without a program id.")
                Throw oExcep
            End If
        Catch ex As Exception
            Response.Write(ex.ToString())
            Response.End()
        End Try

        If Not IsPostBack Then
            sql = "select IsNull(application_id,0) as application_id from Course where id=" & CInt(Request.QueryString("id"))
            rs = BQ(sql)
            If rs.Read() Then
                application_id = rs("application_id")
            Else
                Response.Write("<h1>An application has not been assigned to this program</h1>")
                Response.End()
            End If
            EQ(rs)

            If application_id = 0 Then
                Response.Write("<h1>An application has not been assigned to this program</h1>")
                Response.End()
            End If

            sql = "ApplicationGetDetail " & application_id
            rs = BQ(sql)
            If rs.Read() Then
                If CStr(Session("application_user_id")) = "" Then

                Else

                    'Need find associated submission
                    sql2 = "select application_submission_id as id from application_user_app a, application_submission b where a.application_user_id=" & CInt(Session("application_user_id")) & " and b.application_id=" & CInt(application_id) & " and b.course_id=" & CInt(Request.QueryString("id")) & " and b.id=a.application_submission_id"
                    rs2 = BQ2(sql2)
                    If rs2.Read() Then
                        Session("application_submission_id") = rs2("id")
                    Else
                        Session("application_submission_id") = "" ' Could not find submission
                    End If
                    EQ2(rs2)


                End If

            End If
            EQ(rs)
        End If
    End Sub


    Sub RenderFooter()

    End Sub

    Function ProcessLabel(label As String) As String
        label = Replace(label, "'", "\x27")    '  encode apostrophes 
        label = Replace(label, """", "\x22")   '  encode double-quotes 
        label = replace(label, vbCrLf, " ")
        label = replace(label, vbCr, " ")
        label = replace(label, vbLf, " ")
        Dim objRegExp As New RegEx("<(.|\n)+?>") ' replace html tags
        label = objRegExp.Replace(label, "")
        ProcessLabel = label
    End Function

    Sub GenPageContent(page_id As Integer, course_id As Integer, displayButtons As Boolean)
        Dim sql, rs, sql2, rs2
        Dim oExcep As Exception


        Dim application_id
        Try

            sql = "select id from Application_Page where id=" & page_id
            rs = BQ(sql)
            If rs.Read() Then
                page_id = rs("id")
            Else
                EQ(rs)
                Response.Write(page_id)
                oExcep = New Exception("The page does not exist.")
                Throw oExcep
            End If
            EQ(rs)



            Try
                course_id = CInt(course_id)
            Catch
                oExcep = New Exception("An attempt has been made to render this page with an invalid program id.")
                Throw oExcep
            End Try
            If course_id = 0 Then
                oExcep = New Exception("An attempt has been made to render this page without a program id.")
                Throw oExcep
            End If
        Catch ex As Exception
            Response.Write(ex.ToString())
            Response.End()
        End Try

        If Request.Form("save_continue") = "Save & Continue" Or Request.Form("save_continue") = "Continue" Then
            ProcessSubmission("Save")
            ProcessContinue()
        End If

        If Request.Form("faction") = "Next" Then
            ProcessSubmission("Save")
            ProcessContinue()
        End If


        If Request.Form("faction") = "Submit" Then
            ProcessSubmission("Save")
            ProcessEvent()
        End If


        If Request.Form("faction") = "Save" Then
            ProcessSubmission("Save")

            If Session("logged_in") <> 1 Then
                Dim fromPage = Request.Url.AbsoluteUri
                Session("from_page") = fromPage
                Response.Redirect("../app/application_register.aspx")
            End If
        End If




        sql = "select IsNull(application_id,0) as application_id from Course where id=" & course_id
        rs = BQ(sql)
        If rs.Read() Then
            application_id = rs("application_id")
        Else
            Response.Write("<h1>An application has not been assigned to this program</h1>")
            Response.End()
        End If
        EQ(rs)

        If application_id = 0 Then
            Response.Write("<h1>An application has not been assigned to this program</h1>")
            Response.End()
        End If

        Try
            sql = "select name from course where id=" & Request.QueryString("id")
            rs = BQ(sql)
            If rs.Read() Then

            End If
            EQ(rs)
        Catch
            oExcep = New Exception("An attempt has been made to render this page with an invalid program id.")
            Throw oExcep

        End Try
        Dim displaySubmit = False
        sql = "ApplicationGetQuestions " & application_id & "," & CInt(page_id)
        rs = BQ(sql)
        Dim headerSeen = False
        If rs.Read() Then

            Response.Write("<form method=""post"" id=""appForm""  enctype=""multipart/form-data"" onsubmit=""return true;"" action="""">" & vbCrLf)


            Response.Write("<input type=""hidden"" name=""application_id"" value=""" & application_id & """/>" & vbCrLf)
            Response.Write("<input type=""hidden"" name=""course_id"" value=""" & course_id & """/>" & vbCrLf)
            Response.Write("<input type=""hidden"" name=""page_id"" value=""" & page_id & """/>" & vbCrLf)
            Response.Write("<input type=""hidden"" id=""faction"" name=""faction"" value=""""/>" & vbCrLf)

            Do
                Response.Write(vbCrLf)
                Select Case (rs("type"))
                    Case "Yes/No"
                        RenderYesNo(rs)
                        displaySubmit = True
                    Case "Numeric Rating"
                        RenderNumericRating(rs)
                        displaySubmit = True
                    Case "Text"
                        RenderText(rs)
                    Case "Text - Header"
                        If headerSeen = True Then Response.Write("</div>")

                        RenderTextHeader(rs)
                        Response.Write("<div class=""form-section-wrap"">")
                        headerSeen = True
                    Case "Work Phone", "Fax", "Home Phone", "HR Work Phone", "HR Fax", "Emergency Contact Phone", "Cell Phone"
                        RenderTextInputLength(rs, 15)
                        displaySubmit = True
                    Case "Date of Birth"
                        RenderDateOfBirth(rs)
                        displaySubmit = True
                    Case "Text Input", "People Supervised"
                        RenderTextInput(rs)
                        displaySubmit = True
                    Case "Text Input (Specify Size)"
                        RenderTextInputSize(rs)
                        displaySubmit = True
                    Case "First Name", "Middle Name", "Last Name", "City", "Home City", "Billing City", "HR First Name", "HR Middle Name", "HR Last Name", "HR City"
                        RenderTextInputLength(rs, 100)
                        displaySubmit = True
                    Case "Title", "HR Title", "Address 1", "HR Address 1", "HR Address 2", "Home Address 1", "Address 2", "Home Address 2", "HR Email", "Billing Address 1", "Billing Address 2", "Billing State International", "Home State International", "HR State International", "State International", "HR Company Name", "Company Name", "Reports To"
                        RenderTextInputLength(rs, 200)
                        displaySubmit = True
                    Case "Email"
                        RenderTextInputEmailLength(rs, 200)
                        displaySubmit = True
                    Case "Nickname"
                        RenderTextInputLength(rs, 50)
                        displaySubmit = True
                    Case "Flatbridge Custom Field"
                        RenderTextInputLength(rs, 2000)
                        displaySubmit = True
                    Case "Zip", "Billing Zip", "Home Zip", "HR Zip"
                        RenderTextInputLength(rs, 13)
                        displaySubmit = True
                    Case "Emergency Contact Name", "Emergency Contact Relationship"
                        RenderTextInputLength(rs, 200)
                        displaySubmit = True
                    Case "State", "Billing State", "Home State", "HR State"
                        RenderState(rs)
                        displaySubmit = True
                    Case "Management Function"
                        RenderManagementFunction(rs)
                        displaySubmit = True
                    Case "Country", "Billing Country", "Home Country", "HR Country"
                        RenderCountry(rs)
                        displaySubmit = True
                    Case "Salary Range"
                        RenderSalaryRange(rs)
                        displaySubmit = True
                    Case "File"
                        RenderFile(rs)
                        displaySubmit = True
                    Case "Gender", "HR Gender"
                        RenderGender(rs)
                        displaySubmit = True
                    Case "Salutation", "HR Salutation"
                        RenderSalutation(rs)
                        displaySubmit = True
                    Case "Text - Multiline", "Text - Multiline - Sales Funnel Notes"
                        RenderTextMultiline(rs)
                        displaySubmit = True
                    Case "Drop Down"
                        RenderDropDown(rs)
                        displaySubmit = True
                    Case "Drop Down - Multiple Select"
                        RenderDropDownMultipleSelect(rs)
                        displaySubmit = True
                    Case "Radio"
                        RenderRadio(rs)
                        displaySubmit = True
                    Case "Radio (Money)"
                        RenderRadioMoney(rs)
                        displaySubmit = True
                    Case "Check Box"
                        RenderCheckBox(rs)
                        displaySubmit = True
                    Case "Check Box (Money)"
                        RenderCheckBoxMoney(rs)
                        displaySubmit = True
                    Case "Employment Record"
                        RenderEmploymentRecord(rs, headerSeen)
                        displaySubmit = True
                    Case "Professional Education"
                        RenderProfessionalEducation(rs, headerSeen)
                        displaySubmit = True
                    Case "Education"
                        RenderEducation(rs, headerSeen)
                        displaySubmit = True
                    Case "Source Code"
                        RenderSourceCode(rs)
                        displaySubmit = True
                    Case "Industry"
                        RenderIndustry(rs)
                        displaySubmit = True
                    Case Else
                        Response.Write("<tr><td colspan=""2"">" & rs("type") & " - not yet implemented</td></tr>")
                End Select

            Loop Until rs.Read() = 0
            If headerSeen = True Then Response.Write("</div>")

            Dim isEvent = False
            sql2 = "select IsNull(is_event,0) as is_event from course_webex a where a.course_id=" & CInt(course_id)
            rs2 = BQ2(sql2)
            If rs2.Read() Then
                isEvent = rs2("is_event")

            End If
            EQ2(rs2)

            If isEvent = True Then
                RenderEventSubmitButton(course_id)
            Else

                If displaySubmit = True And displayButtons = True Then

                    RenderSubmitButton(course_id)
                ElseIf displayButtons = True Then
                    RenderContinueButton()
                End If
            End If


            Response.Write("</form></div></div>")
        Else
            Response.Write("<table id=""applicationQuestions"">")
            Response.Write("<tr><td>This page has not been setup.<div class=""br""/></td></tr>")
            Response.Write("</table>")
        End If
        EQ(rs)

        RenderFooter()



    End Sub

    Sub ProcessEvent()
        Dim sql, rs
        Dim fee
        Dim application_id
        sql = "select application_id from course where id=" & CInt(Request.QueryString("id"))
        rs = BQ(sql)
        If rs.Read() Then
            application_id = rs("application_id")
        End If
        EQ(rs)

        SendEventEmail()

        sql = "GetApplicationInformation " & CInt(Session("application_user_id")) & "," & CInt(Request.QueryString("id")) & "," & CInt(application_id)
        rs = BQ(sql)
        If rs.Read() Then
            fee = rs("fee")
        End If
        EQ(rs)

        Dim isWebex = False

        sql = "select IsNull(webex_meeting_id,'') as webex_meeting_id from course_webex where course_id=" & CInt(Request.QueryString("id"))
        rs = BQ(sql)
        If rs.Read() Then
            If rs("webex_meeting_id") <> "" Then isWebex = True
        End If
        EQ(rs)

        If fee = 0 Then
            sql = "ApplicationSubmitMultiple " & CInt(Request.QueryString("id")) & ",1," & CInt(Session("application_user_id"))
            rs = BQ(sql)
            EQ(rs)
            Session("application_user_id") = ""
            If isWebex = True Then
                Response.Redirect("../app/thank_you_onstream.aspx?id=" & CInt(Request.QueryString("id")))
            Else
                Response.Redirect("../app/thank_you.aspx?id=" & CInt(Request.QueryString("id")))
            End If
            Response.End()
        Else
            Response.Redirect("../app/credit_card.aspx?id=" & CInt(Request.QueryString("id")))
            Response.End()
        End If


    End Sub

    Sub SendEventEmail()
        Dim email_title, description
        Dim sql, rs, sql2, rs2
        sql = "select email_title, description from SalesEmail a, Course b where b.id=" & CInt(Request.QueryString("id")) & " and b.application_email_script=a.id"
        rs = BQ(sql)
        If rs.Read() Then
            Try
                email_title = rs("email_title")
                description = rs("description")
                Dim staff_email, staff_email2


                Dim myMail As New MailMessage()
                myMail.From = New MailAddress("executive_education@gsb.stanford.edu", "Stanford Graduate School of Business - Office of Executive Education")
                myMail.IsBodyHTML = True
                Dim toAddress As New MailAddress(Session("Email"))
                myMail.To.Add(toAddress)
                Try
                    sql2 = "select email from Course a, Staff b, Person c where c.id=b.person_id and b.id=a.application_staff_member_id and a.id=" & CInt(Request.QueryString("id"))
                    rs2 = BQ2(sql2)
                    If rs2.Read() Then
                        staff_email = rs2("email")
                        If Trim(staff_email) <> "" Then
                            myMail.CC.Add(New MailAddress(staff_email))
                        End If
                    End If
                    EQ2(rs2)
                Catch
                End Try

                Try
                    sql2 = "select c.email from Staff a, Person c,  Staff_Contacts b where b.id=5 and a.person_id=c.id and b.staff_id=a.id"
                    rs2 = BQ2(sql2)
                    If rs2.Read() Then
                        staff_email2 = rs2("email")
                        If Trim(staff_email2) <> "" Then
                            myMail.CC.Add(New MailAddress(staff_email2))
                        End If
                    End If
                    EQ2(rs2)
                Catch ex As Exception

                End Try



                myMail.Subject = email_title

                Dim Token
                Token = "!FirstName!"
                description = Replace(description, Token, Session("First Name"))
                Token = "!LastName!"
                description = Replace(description, Token, Session("Last Name"))

                myMail.Body = description
                'myMail.Body = "Email 1:" & staff_email & "<br>" & "Email 2:" & staff_email2 & "<br>" & description
                Dim smtpSender As New SmtpClient(System.Configuration.ConfigurationManager.AppSettings("MailServer"))
                smtpSender.UseDefaultCredentials = True
                smtpSender.Port = 25

                smtpSender.Send(myMail)
            Catch ex As Exception
                Response.Write(ex.ToString())
            End Try

        End If
        EQ(rs)
    End Sub


    Sub GenerateLengthSelection(days As Integer, ByRef isSeen As Boolean)
        If isSeen = False Then
            Response.Write(" class=""")
        Else
            Response.Write(" ")
        End If
        Dim tag
        If days <= 1 Then
            tag = "length1_sect"
        ElseIf days <= 60 Then
            tag = "length2_sect"
        Else
            tag = "length3_sect"
        End If


        Response.Write(tag)

        isSeen = True

    End Sub
    Sub GenerateTimingSelection(course_id As Integer, ByRef isSeen As Boolean)
        If isSeen = False Then
            Response.Write(" class=""")
        Else
            Response.Write(" ")
        End If
        'Go get the class
        Dim sql, rs
        sql = "select Right('0' + Convert(VarChar(2), Month(start_date)), 2) as month from course where id=" & course_id
        rs = BQ(sql)
        If rs.Read() Then
            Response.Write("timing_" & rs("month") & "_sect")
        End If
        EQ(rs)

        isSeen = True
    End Sub


    Sub GenerateMonthSelection(course_id As Integer)
        Dim J
        Dim sql, rs, sql2, rs2
        'Get all occurrences for this course name within this year
        Dim TD(12) As String
        Dim SkipTD(12) As Integer

        For J = 0 To 11
            TD(J) = "<td>&nbsp;</td>" & vbCrLf
            SkipTD(J) = 0

        Next J
        'Get start day and month for each occurrence and get end day and month for each occurrence
        Dim startMonth
        Dim startDay
        Dim endMonth
        Dim endDay
        Dim endYear
        Dim startYear
        Dim startMonthName
        Dim endMonthName


        sql2 = "select IsNull(c.link_to_program_detail,'') as link_to_program_detail, a.id from course a left outer join course b on a.name=b.name left outer join course_webex c on c.course_id=a.id where a.start_date>getDate() and a.start_date <  dateadd(yy,1,convert(varchar,DATEADD(dd,-(DAY(getdate())-1),getdate()),101)) and a.active=1  and display_program_finder=1 and b.id=" & course_id & " order by a.start_date"
        rs2 = BQ2(sql2)
        If rs2.Read() Then


            Do
                sql = "select upper(left(datename(month, start_date), 3)) as startMonthName, upper(left(datename(month, end_date), 3)) as endMonthName, year(start_date) as startYear, year(end_date) as endYear, day(start_date) as startDay,day(end_date) as endDay,	datediff(mm,DATEADD(dd,-(DAY(getdate())-1),getdate()),DATEADD(dd,-(DAY(start_date)-1),start_date)) as startMonth,datediff(mm,DATEADD(dd,-(DAY(getdate())-1),getdate()),DATEADD(dd,-(DAY(end_date)-1),end_date)) as endMonth from course where id not in (select course_id from course_date_override) and id=" & rs2("id")
                rs = BQ(sql)
                If rs.Read Then
                    startMonth = rs("startMonth")
                    endMonth = rs("endMonth")
                    startDay = rs("startDay")
                    endDay = rs("endDay")
                    startYear = rs("startYear")
                    endYear = rs("endYear")
                    startMonthName = rs("startMonthName")
                    endMonthName = rs("endMonthName")

                End If
                EQ(rs)


                Dim Skip = 0
                For J = 0 To 11
                    If Skip > 0 Then
                        SkipTD(J) = 1
                        Skip = Skip - 1
                        Continue For
                    Else
                        If J = startMonth Then
                            If startMonth = endMonth Then
                                If startDay = endDay Then
                                    TD(J) = "<td class=""flagged""><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startDay & "</a></td>" & vbCrLf
                                Else
                                    TD(J) = "<td class=""flagged""><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startDay & "-" & endDay & "</a></td>" & vbCrLf
                                End If
                            Else

                                If endMonth - startMonth > 1 Then

                                    If endMonth > 11 Then endMonth = 11
                                    If endYear > startYear Then
                                        TD(J) = "<td class=""flagged"" colspan=""" & endMonth - startMonth + 1 & """><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startMonthName & " " & startDay & ", " & startYear & " - " & endMonthName & " " & endDay & ", " & endYear & "</a></td>" & vbCrLf
                                    Else
                                        TD(J) = "<td class=""flagged"" colspan=""" & endMonth - startMonth + 1 & """><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;""href=""" & rs2("link_to_program_detail") & """>" & startMonthName & " " & startDay & " - " & endMonthName & " " & endDay & "</a></td>" & vbCrLf

                                    End If
                                    Skip = endMonth - startMonth
                                Else
                                    If endMonth > 11 Then endMonth = 11
                                    TD(J) = "<td class=""flagged"" colspan=""" & endMonth - startMonth + 1 & """><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startDay & "-" & endDay & "</a></td>" & vbCrLf
                                    Skip = endMonth - startMonth

                                End If

                            End If
                        Else

                        End If
                    End If
                Next J
                sql = "select upper(left(datename(month, date), 3)) as startMonthName,upper(left(datename(month, dateadd(dd,days-1,date)), 3)) as endMonthName,year(date) as startYear,year(dateadd(dd,days-1,date)) as endYear,day(date) as startDay,day(dateadd(dd,days-1,date)) as endDay,datediff(mm,DATEADD(dd,-(DAY(getdate())-1),getdate()),DATEADD(dd,-(DAY(date)-1),date)) as startMonth,datediff(mm,DATEADD(dd,-(DAY(getdate())-1),getdate()),DATEADD(dd,-(DAY(dateadd(dd,days-1,date))-1),dateadd(dd,days-1,date))) as endMonth from course a, course_date_override b where	b.course_id=a.id and a.id=" & rs2("id") & " order by b.date"

                rs = BQ(sql)
                If rs.Read Then
                    For J = 0 To 11
                        TD(J) = "<td>&nbsp;</td>" & vbCrLf
                        SkipTD(J) = 0

                    Next J
                    Do
                        startMonth = rs("startMonth")
                        endMonth = rs("endMonth")
                        startDay = rs("startDay")
                        endDay = rs("endDay")
                        startYear = rs("startYear")
                        endYear = rs("endYear")
                        startMonthName = rs("startMonthName")
                        endMonthName = rs("endMonthName")
                        Skip = 0
                        For J = 0 To 11
                            If Skip > 0 Then
                                SkipTD(J) = 1
                                Skip = Skip - 1
                                Continue For
                            Else
                                If J = startMonth Then
                                    If startMonth = endMonth Then
                                        If startDay = endDay Then
                                            TD(J) = "<td class=""flagged""><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startDay & "</a></td>" & vbCrLf
                                        Else
                                            TD(J) = "<td class=""flagged""><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startDay & "-" & endDay & "</a></td>" & vbCrLf
                                        End If
                                    Else

                                        If endMonth - startMonth > 1 Then

                                            If endMonth > 11 Then endMonth = 11
                                            If endYear > startYear Then
                                                TD(J) = "<td class=""flagged"" colspan=""" & endMonth - startMonth + 1 & """><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startMonthName & " " & startDay & ", " & startYear & " - " & endMonthName & " " & endDay & ", " & endYear & "</a></td>" & vbCrLf
                                            Else
                                                TD(J) = "<td class=""flagged"" colspan=""" & endMonth - startMonth + 1 & """><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;""href=""" & rs2("link_to_program_detail") & """>" & startMonthName & " " & startDay & " - " & endMonthName & " " & endDay & "</a></td>" & vbCrLf

                                            End If
                                            Skip = endMonth - startMonth
                                        Else
                                            If endMonth > 11 Then endMonth = 11
                                            TD(J) = "<td class=""flagged"" colspan=""" & endMonth - startMonth + 1 & """><a style=""background-color:#f3e8cc; color:#a63a00; font-weight:700;"" href=""" & rs2("link_to_program_detail") & """>" & startDay & "-" & endDay & "</a></td>" & vbCrLf
                                            Skip = endMonth - startMonth

                                        End If

                                    End If
                                Else

                                End If
                            End If
                        Next J
                    Loop Until rs.Read() = 0
                    Dim firstSeen = False
                    For J = 0 To 11
                        If TD(J) = "<td>&nbsp;</td>" & vbCrLf Then
                            If firstSeen = True Then
                                TD(J) = "<td>Back at your office</td>" & vbCrLf
                            End If
                        Else
                            If firstSeen = False Then
                                firstSeen = True
                            Else
                                firstSeen = False
                            End If


                        End If
                    Next J
                End If
                EQ(rs)



            Loop Until rs2.Read() = 0
        End If
        EQ2(rs2)

        For J = 0 To 11
            If SkipTD(J) = 0 Then
                Response.Write(TD(J))
            End If
        Next J


    End Sub




    Sub GenerateCalendarMonths()
        Dim sql, rs, J
        Dim Months(12) As String

        J = 0
        sql = "GetMonths"
        rs = BQ(sql)
        If rs.Read() Then
            Do
                Months(J) = rs("month")
                J = J + 1
            Loop Until rs.Read() = 0
        End If
        EQ(rs)
        J = 0
        Response.Write("<div class=""month-list""><ul>")

        For J = 0 To 11
            Response.Write("<li ")
            If J = 0 Then Response.Write("class=""active"" ")
            Response.Write("rel=""" & J & """>" & Months(J) & "</li>" & vbCrLF)
        Next
        Response.Write("</ul></div>")
    End Sub


    Sub GenerateCalendarContent()
        Dim sql, rs
        Dim currentMonth = ""
        Dim itemSeen = False
        sql = "[GetProgramCalendar]"
        rs = BQ(sql)
        If rs.Read() Then
            Do
                If currentMonth <> rs("month_name") Then
                    If itemSeen = True Then
                        Response.Write("</table></div>")
                    End If
                    Response.Write("<div class=""item""><table border=""0"" cellspacing=""0"" cellpadding=""0"" id=""tb-" & rs("month_index") & """>")
                    Response.Write("<tr><th scope=""col"" colspan=""2"">" & rs("month_name") & "</th></tr>")
                End If
                Response.Write("<tr><td class=""event-date"">" & rs("date_range") & "</td>")
                Response.Write("<td class=""event-detail""><a href=""" & rs("link_to_program_detail") & """>" & rs("name") & "</a></td></tr>")

                itemSeen = True
                currentMonth = rs("month_name")
            Loop Until rs.Read() = 0
            Response.Write("</table></div>")
        End If

        EQ(rs)

    End Sub


    Sub GenerateProgramTableWrap()
        Dim sql, rs, sql2, rs2, J, sql3, rs3
        Dim Months(12) As String

        J = 0
        sql = "GetMonths"
        rs = BQ(sql)
        If rs.Read() Then
            Do
                Months(J) = rs("month")
                J = J + 1
            Loop Until rs.Read() = 0
        End If
        EQ(rs)
        sql = "select id, program_category from program_category where active=1 order by display_order"
        rs = BQ(sql)
        If rs.Read() Then
            Do

                Response.Write("<div class=""section-table-wrap"">" & vbCrlf)
                Response.Write("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""prog_cat" & rs("id") & "_sect"">")
                Response.Write("<tr>")
                Response.Write("<th scope=""col"" class=""th-title"">" & rs("program_category") & "</th>" & vbCrLf)

                J = 0
                For J = 0 To 11
                    Response.Write("<th scope=""col"" class=""th-month"">" & Months(J) & "</th>" & vbCrLF)
                Next
                Response.Write("</tr>")
                sql2 = "select datediff(dd,a.start_date,a.end_date)+1 as days,IsNull(d.link_to_program_detail,'') as link_to_program_detail,a.id,a.name,a.content, convert(varchar,start_date,101) as start_date,convert(varchar,application_deadline,101) as application_deadline, a.fee,IsNull(c.name,'') as location from course a left outer join location c on c.id=a.location_id, course_program_category b, course_webex d where d.course_id=a.id and display_program_finder=1 and a.start_date>getDate() and a.start_date <  dateadd(yy,1,convert(varchar,DATEADD(dd,-(DAY(getdate())-1),getdate()),101)) and a.active=1 and a.id=b.course_id and b.program_category_id=" & rs("id") & " order by a.name,a.start_date"
                rs2 = BQ2(sql2)

                Dim program_name = ""
                If rs2.Read() Then
                    Do
                        If program_name = rs2("name") Then

                        Else

                            program_name = rs2("name")

                            Response.Write("<tr")
                            'Go get the career stages and add the style
                            sql3 = "select career_stage_id from course_career_stage where course_id=" & rs2("id")
                            rs3 = BQ3(sql3)
                            Dim isSeen = False
                            If rs3.Read() Then
                                Do
                                    If isSeen = True Then
                                        Response.Write(" ")
                                    Else
                                        Response.Write(" class=""")
                                    End If
                                    Response.Write("car_stag" & rs3("career_stage_id") & "_sect")
                                    isSeen = True
                                Loop Until rs3.Read() = 0
                            End If
                            EQ3(rs3)

                            GenerateTimingSelection(rs2("id"), isSeen)
                            GenerateLengthSelection(rs2("days"), isSeen)

                            If isSeen = True Then Response.Write("""")



                            Response.Write(">")
                            Response.Write("<th scope=""row"">")
                            Response.Write("<div class=""program-wrap"">")
                            Response.Write("<div class=""program-name""><a style=""font-size:12px; font-weight:700; text-transform:uppercase; color:#ff8700;"" href=""" & rs2("link_to_program_detail") & """>" & rs2("name") & "</a></div>")
                            Response.Write("<div class=""program-details"">")
                            Response.Write("<p>" & rs2("content") & "</p>")
                            Response.Write("<div class=""body-emph"">")
                            Response.Write("<ul>")
                            Response.Write("<li class=""deadline"">Start Date: " & rs2("start_date") & "</li>")
                            Response.Write("<li class=""deadline"">Deadline: " & rs2("application_deadline") & "</li>")
                            Response.Write("<li class=""tuition"">Tuition: " & FormatCurrency(rs2("fee"), 0) & "</li>")


                            If rs2("location") <> "" Then
                                Response.Write("<li class=""location"">Location: " & rs2("location") & "</li>")
                            End If

                            Dim Finish = False
                            'Go get the next program
                            sql3 = "select a.id,a.name,a.content, convert(varchar,start_date,101) as start_date,convert(varchar,application_deadline,101) as application_deadline, a.fee,IsNull(c.name,'') as location from course a left outer join location c on c.id=a.location_id, course_program_category b, course_webex d where d.course_id=a.id and a.start_date>getDate() and a.start_date <  dateadd(yy,1,convert(varchar,DATEADD(dd,-(DAY(getdate())-1),getdate()),101)) and  display_program_finder=1 and a.active=1 and a.id=b.course_id and b.program_category_id=" & rs("id") & " and a.name='" & db(rs2("name")) & "' and a.id<>" & rs2("id") & " order by a.name,a.start_date"
                            rs3 = BQ3(sql3)

                            If rs3.Read() Then
                                Do


                                    Response.Write("<li><br/></li>")
                                    Response.Write("<li class=""deadline"">Start Date: " & rs3("start_date") & "</li>")
                                    Response.Write("<li class=""deadline"">Deadline: " & rs3("application_deadline") & "</li>")
                                    Response.Write("<li class=""tuition"">Tuition: " & FormatCurrency(rs3("fee"), 0) & "</li>")

                                    If rs3("location") <> "" Then
                                        Response.Write("<li class=""location"">Location: " & rs3("location") & "</li>")
                                    End If

                                Loop Until rs3.Read() = 0
                            End If
                            EQ3(rs3)



                            Response.Write("</ul>")
                            Response.Write("</div>")
                            Response.Write("</div>")
                            Response.Write("<div class=""btn-show-prog-details""></div>")
                            Response.Write("</div>")
                            Response.Write("</th>")

                            GenerateMonthSelection(rs2("id"))




                            Response.Write("</tr>")


                        End If

                    Loop Until rs2.Read() = 0
                End If
                EQ2(rs2)

                Response.Write("</table>")
                Response.Write("</div>")

            Loop Until rs.Read() = 0
        End If
        EQ(rs)

    End Sub
    Sub GenerateProgramCategory()
        Dim sql, rs
        sql = "select id,program_category from program_category where active=1 order by display_order"
        rs = BQ(sql)
        If rs.Read() Then
            Do
                Response.Write("<li><input type=""checkbox"" checked=""checked"" id=""prog_cat" & rs("id") & """/><label for=""prog_cat" & rs("id") & """>" & rs("program_category") & "</label></li>" & VbCrLf)
            Loop Until rs.Read() = 0
        End If
        EQ(rs)
    End Sub


    Sub GenerateCareerStage()
        Dim sql, rs
        sql = "select id,career_stage from career_stage where active=1 order by display_order"
        rs = BQ(sql)
        If rs.Read() Then
            Do
                Response.Write("<li><input type=""checkbox"" checked=""checked"" id=""car_stag" & rs("id") & """/><label for=""prog_cat" & rs("id") & """>" & rs("career_stage") & "</label></li>" & VbCrLf)
            Loop Until rs.Read() = 0
        End If
        EQ(rs)
    End Sub
End Class
