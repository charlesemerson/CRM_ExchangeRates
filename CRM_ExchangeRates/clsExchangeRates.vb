Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports Newtonsoft.Json

Public Class clsFunctionDetails
    Public ElapsedTime As String = ""
    Public EndTime As Object = Nothing
    Public ErrorMsg As String = ""
    Public FileName As String = ""
    Public RecordsAdded As Integer = 0
    Public RecordsDeleted As Integer = 0
    Public RecordsExported As Integer = 0
    Public RecordsUpdated As Integer = 0
    Public ReturnMsg As String = ""
    Public RoutineName As String = ""
    Public RoutineType As String = ""
    Public StartTime As Object = Nothing
    Public Success As Boolean = False
End Class

Public Class OpportunityDetails
    Public OID As Integer = 0
    Public SubscriberID As Integer = 0
    Public LastActivityDate As Date = Nothing
    Public Revenue As Integer = 0
    Public Updated As Date = Nothing
    Public UpdatedBy As String = ""
End Class

Public Class CurrencyDetails
    Public CID As Integer = 0
    Public CurrencyCode As String = ""
    Public CurrencyName As String = ""
    Public Updated As Date = Nothing
    Public UpdatedBy As String = ""
End Class

Public Class ExchangeRateDetails
    Public EID As Integer = 0
    Public CurrencyCode As String = ""
    Public ExchangeDate As Date = Nothing
    Public ExchangeRate As Decimal = 0
End Class

Public Class clsExchangeRates

#Region " Private Variables "
    Private cmdDelete As SqlCommand
    Private cmdSelect As SqlCommand
    Private cmdUpdate As SqlCommand
    Private da As SqlDataAdapter
    Private ds As DataSet
    Private dt As DataTable
    Private dtRow As DataRow
    Private strConn As String = ConfigurationManager.ConnectionStrings("ConnString").ToString
    Private strSQL As String
#End Region

#Region " Public Variables and Properties "
    'OpenExchangeRates API
    Public strAppID As String = ConfigurationManager.AppSettings("OER_API_Key").ToString
    'Function Results
    Public blnSuccess As Boolean = False
    Public strElapsedTime As String
    Public strErrorMsg As String
    Public strReturnMsg As String
    Public strRoutineName As String
    Public strRoutineType As String

    Public Property ConnString() As String
        Get
            Return strConn
        End Get
        Set(ByVal Value As String)
            strConn = Value
        End Set
    End Property

#End Region

#Region " Toolbox "

    Public Function CleanText(ByVal strTextIn As String) As String
        strTextIn = Replace(strTextIn, "'", "")
        Return strTextIn
    End Function

    Public Function Cs(ByVal strTextIn As String)
        strTextIn = Replace(strTextIn, "'", "''")
        strTextIn = Replace(strTextIn, "<", Chr(60))
        strTextIn = Replace(strTextIn, ">", Chr(62))
        Return strTextIn
    End Function

    Public Function GetStartDateForRoutine(strRoutineType As String) As String
        Dim strStartDate As String = ""
        Dim strSQL As String = ""
        strSQL += "SELECT StartTime "
        strSQL += "FROM tblDataBridgeLog "
        strSQL += "WHERE (RoutineType = '" & strRoutineType & "') "
        strSQL += "AND (Success=1) "
        strSQL += "ORDER BY StartTime DESC "
        Dim ds As Data.DataSet = FillDataset(strSQL)
        Dim dt As Data.DataTable = ds.Tables(0)
        ds = New Data.DataSet : ds = FillDataset(strSQL)
        dt = New Data.DataTable : dt = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            strStartDate = Nz(dt.Rows(0)("StartTime"), "")
        End If
        Return strStartDate
    End Function

    Public Function GetTimeStamp(ByVal strDateTime As String) As String
        Dim strTimeStamp As String = ""
        Dim strYear As String = ""
        Dim strMonth As String = ""
        Dim strDay As String = ""
        Dim strHour As String = ""
        Dim strMinute As String = ""
        Dim strSecond As String = ""
        Dim dtmDate As Date = Nothing
        If strDateTime = "" Then
            dtmDate = Date.Now
        Else
            dtmDate = DateTime.Parse(strDateTime)
        End If
        strYear = Year(dtmDate)
        strMonth = PadString(Month(dtmDate), 2, "0")
        strDay = PadString(Day(dtmDate), 2, "0")
        strHour = PadString(Hour(dtmDate), 2, "0")
        strMinute = PadString(Minute(dtmDate), 2, "0")
        strSecond = PadString(Second(dtmDate), 2, "0")
        strTimeStamp = "" & strYear & "_" & strMonth & "_" & strDay & "_" & strHour & "_" & strMinute & "_" & strSecond & ""
        Return strTimeStamp
    End Function

    Public Function FormatBit(ByVal bln As Boolean)
        Dim intOutput As Integer
        If bln = True Then
            intOutput = 1
        Else
            intOutput = 0
        End If
        Return intOutput
    End Function

    Public Function FormatDate(ByVal objDate As Object) As String
        Dim strDate As String = ""
        Dim dtmDate As Date
        If objDate Is Nothing Or objDate Is DBNull.Value Then
            Return strDate
        End If
        Try
            dtmDate = CDate(objDate)
            strDate = Format(dtmDate, "MM/dd/yyyy")
        Catch ex As Exception

        End Try
        Return strDate
    End Function

    Public Function FormatDBDate(ByVal dtm As Object) As String
        Try
            If dtm = Nothing Or Len(dtm.ToString) < 1 Then
                Return "NULL"
            Else
                Return "'" & dtm.ToString & "'"
            End If
        Catch ex As Exception
            Return "NULL"
        End Try
    End Function

    Public Function InitializeFunctionDetails(ByVal fd As clsFunctionDetails) As clsFunctionDetails
        fd.RoutineName = ""
        fd.RoutineType = ""
        fd.ReturnMsg = ""
        fd.ErrorMsg = ""
        fd.ElapsedTime = ""
        fd.Success = False
        Return fd
    End Function

    Public Sub InitializePublicVariables()
        strRoutineName = ""
        strRoutineType = ""
        strReturnMsg = ""
        strErrorMsg = ""
        strElapsedTime = ""
        blnSuccess = False
    End Sub

    Public Function IsValidDate(ByVal dtmDate As String) As Boolean
        Dim dtmDateTimeUS As System.DateTime
        Dim format As New System.Globalization.CultureInfo("en-US", True)
        Dim bln As Boolean = False
        If Not dtmDate = Nothing Then
            Try
                dtmDate = CDate(dtmDate)
                dtmDateTimeUS = System.DateTime.Parse(dtmDate, format)
                bln = True
            Catch ex As Exception
                bln = False
            End Try
        End If
        Return bln
    End Function

    Public Function GetElapsedTime(ByVal dtmStart As Date) As String
        Dim dtmEnd As Date = Date.Now
        Dim intHours As Integer
        Dim intMinutes As Integer
        Dim intSeconds As Integer
        Dim intTimeDiff As Integer
        Dim strElapsedTime As String
        'Get elapsed time in seconds
        intTimeDiff = DateDiff(DateInterval.Second, dtmStart, dtmEnd)
        'Get seconds part
        intSeconds = intTimeDiff Mod 60
        intTimeDiff = intTimeDiff - intSeconds
        intTimeDiff = intTimeDiff / 60
        'Get minutes part
        intMinutes = intTimeDiff Mod 60
        intTimeDiff = intTimeDiff - intMinutes
        'Get hours part
        intHours = intTimeDiff / 60
        strElapsedTime = PadString(CStr(intHours), 2, "0") & ":" & PadString(CStr(intMinutes), 2, "0") & ":" & PadString(CStr(intSeconds), 2, "0")
        Return strElapsedTime
    End Function

    Public Function Nz(ByVal objValue As Object, ByVal objNullValue As Object)
        If objValue Is DBNull.Value Then
            Return objNullValue
        Else
            Return objValue
        End If
    End Function

    Public Function PadString(ByVal strNo As String, ByVal intLength As Integer, ByVal PadChar As String) As String
        ' Pad number with zeros
        Dim intI As Integer
        For intI = 1 To intLength
            If Len(strNo) < intLength Then
                strNo = PadChar & strNo
            End If
        Next
        Return strNo
    End Function

    Public Function Zz(ByVal objValue As Object, ByVal objZeroValue As Object)
        If objValue = "" Then
            Return objZeroValue
        Else
            Return objValue
        End If
    End Function

#End Region

#Region " Database "

    'MSSQL
    Public Class Parameter
        Public Name As String
        Public Value
    End Class

    Public Function AddRecord(ByVal strSQL As String, ByVal strConn As String) As Integer
        Dim intRecordsAffected As Integer = 0
        Try
            Dim conn As SqlConnection
            conn = New SqlConnection(strConn)
            conn.ConnectionString = strConn
            conn.Open()
            Dim cmdInsert As New SqlCommand(strSQL, conn)
            cmdInsert.CommandTimeout = 0
            intRecordsAffected = cmdInsert.ExecuteNonQuery()
            conn.Close()
            conn.Dispose()
            cmdInsert.Dispose()
            conn = Nothing
            cmdInsert = Nothing
        Catch ex As Exception
            strReturnMsg = "Error Adding Record"
            strErrorMsg = ex.Message
            Console.WriteLine(strReturnMsg & " | " & strErrorMsg)
            blnSuccess = False
            GoTo wrap_up
        End Try
wrap_up:
        Return intRecordsAffected
    End Function

    Public Function AddRecordGetID(ByVal strSQL As String, ByVal strConn As String) As Integer
        'This function adds a record to a table
        'with identity primary key and retrieves newly added ID
        Dim intID As Integer = 0
        Dim conn As SqlConnection
        conn = New SqlConnection(strConn)
        conn.ConnectionString = strConn
        conn.Open()
        strSQL &= " SELECT @ID = @@identity"
        Dim cmdInsert As New SqlCommand(strSQL, conn)
        cmdInsert.CommandTimeout = 0
        Dim paramID As New SqlParameter("@ID", SqlDbType.Int)
        paramID.Direction = ParameterDirection.Output
        cmdInsert.Parameters.Add(paramID)
        cmdInsert.ExecuteNonQuery()
        'Get ID of added record
        intID = Nz(paramID.Value, 0)
        conn.Close()
        conn.Dispose()
        cmdInsert.Dispose()
        conn = Nothing
        cmdInsert = Nothing
        Return intID
    End Function

    Public Sub DeleteRecord(ByVal strSQL As String, ByVal strConn As String)
        Dim conn As SqlConnection
        conn = New SqlConnection(strConn)
        cmdDelete = New SqlCommand(strSQL, conn)
        conn.Open()
        cmdDelete.ExecuteNonQuery()
        conn.Close()
        conn.Dispose()
        cmdDelete.Dispose()
        conn = Nothing
    End Sub

    Public Function FillDataset(ByVal strSQL As String) As DataSet
        Dim conn As SqlConnection
        Dim ds As DataSet
        conn = New SqlConnection(strConn)
        cmdSelect = New SqlCommand(strSQL, conn)
        cmdSelect.CommandTimeout = 0
        da = New SqlDataAdapter
        da.SelectCommand = cmdSelect
        ds = New DataSet
        da.Fill(ds)
        Return ds
    End Function

    Public Function RecordExists(intID As Integer, strTableName As String) As Boolean
        Dim blnExists As Boolean = False
        Dim strSQL As String = ""
        Dim ds As DataSet = Nothing
        Dim dt As DataTable = Nothing
        strSQL = "SELECT ID FROM " & strTableName & " WHERE (ID = " & intID & ") "
        ds = FillDataset(strSQL)
        dt = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            blnExists = True
        End If
        dt.Dispose() : dt = Nothing
        ds.Dispose() : ds = Nothing
        Return blnExists
    End Function

    Public Sub TruncateTable(ByVal strTable As String, ByVal strConn As String)
        Dim strSQL As String
        strSQL = "TRUNCATE TABLE " & strTable
        UpdateRecord(strSQL, strConn)
    End Sub

    Public Function UpdateRecord(ByVal strSQL As String, ByVal strConn As String) As Integer
        Dim intRecordsAffected As Integer = 0
        Try
            Dim conn As SqlConnection
            conn = New SqlConnection(strConn)
            cmdUpdate = New SqlCommand(strSQL, conn)
            cmdUpdate.CommandTimeout = 0
            conn.Open()
            intRecordsAffected = cmdUpdate.ExecuteNonQuery()
            conn.Close()
            conn = Nothing
            cmdUpdate = Nothing
        Catch ex As Exception
            strReturnMsg = "Error Adding Record"
            strErrorMsg = ex.Message
            blnSuccess = False
            GoTo wrap_up
        End Try
wrap_up:
        Return intRecordsAffected
    End Function

#End Region

    Public Function CurrencyCodeExists(ByVal strCurrencyCode As String) As Boolean
        Dim blnExists As Boolean = False
        strSQL = "SELECT ID FROM tblCurrencies "
        strSQL += "WHERE (CurrencyCode = '" & strCurrencyCode & "') "
        Dim ds As Data.DataSet = FillDataset(strSQL)
        Dim dt As Data.DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            blnExists = True
        End If
        dt.Dispose() : dt = Nothing
        ds.Dispose() : ds = Nothing
        Return blnExists
    End Function

    Public Function CurrencyExchangeRateExists(ByVal strCurrencyCode As String, dtmExchangeDate As Date) As Boolean
        Dim blnExists As Boolean = False
        strSQL = "SELECT ID FROM tblCurrencyExchangeRates "
        strSQL += "WHERE (CurrencyCode = '" & strCurrencyCode & "') "
        strSQL += "AND (ExchangeDate = '" & dtmExchangeDate & "') "
        Dim ds As Data.DataSet = FillDataset(strSQL)
        Dim dt As Data.DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            blnExists = True
        End If
        dt.Dispose() : dt = Nothing
        ds.Dispose() : ds = Nothing
        Return blnExists
    End Function

    Public Function AddCurrency(ByVal strCurrencyCode As String, ByVal strCurrencyName As String) As String
        Dim strResult As String = ""
        'Check if Currency Exists
        Dim blnExists As Boolean = CurrencyCodeExists(strCurrencyCode)
        Dim intRecordsAffected As Integer = 0
        If blnExists = False Then
            strSQL = "INSERT INTO tblCurrencies ("
            strSQL += "CreatedByID, "
            strSQL += "CreatedDate,"
            strSQL += "CreatedUserName, "
            strSQL += "CurrencyCode, "
            strSQL += "CurrencyName "
            strSQL += ") VALUES ("
            strSQL += 999 & ", "
            strSQL += FormatDBDate(Date.Now) & ", "
            strSQL += "'ExchangeDataBridge', "
            strSQL += "'" & strCurrencyCode & "', "
            strSQL += "'" & strCurrencyName
            strSQL += "')"
            intRecordsAffected = AddRecord(strSQL, strConn)
            strResult = "Added Currency: " & strCurrencyCode & " - " & strCurrencyName
        Else
            strResult = "Currency Exists: " & strCurrencyCode & " - " & strCurrencyName
        End If
        Return strResult
    End Function

    Public Function AddCurrencyExchangeRate(ByVal strCurrencyCode As String, _
                                            ByVal dtmExchangeDate As Date, _
                                            ByVal dblExchangeRate As Double) As String
        Dim strResult As String = ""
        'Check if Currency Exchange Rate Exists
        Dim blnExists As Boolean = CurrencyExchangeRateExists(strCurrencyCode, dtmExchangeDate)
        Dim intRecordsAffected As Integer = 0
        If blnExists = False Then
            strSQL = "INSERT INTO tblCurrencyExchangeRates ("
            strSQL += "CurrencyCode, "
            strSQL += "ExchangeDate, "
            strSQL += "ExchangeRate "
            strSQL += ") VALUES ("
            strSQL += "'" & strCurrencyCode & "', "
            strSQL += "'" & dtmExchangeDate & "', "
            strSQL += dblExchangeRate & ")"
            intRecordsAffected = AddRecord(strSQL, strConn)
            strResult = "Added Currency Exchange Rate: " & strCurrencyCode & " - " & dtmExchangeDate
        Else
            strResult = "Currency Exchange Rate Exists: " & strCurrencyCode & " - " & dtmExchangeDate
        End If
        Return strResult
    End Function

    Public Function GetCurrencies() As String
        Dim strResult As String = ""
        Dim response As HttpWebResponse = Nothing
        Dim strJsonData As String = ""
        Dim strRequest As String = ""
        Dim strCurrencyCode As String = ""
        Dim strCurrencyName As String = ""
        Try
            Dim request As HttpWebRequest
            Dim reader As StreamReader
            If Len(strAppID) > 0 Then
                'Construct Url to get List of Currencies
                strRequest = "http://openexchangerates.org/api/currencies.json"
                strRequest += "?app_id=" & strAppID
                ' *** COMMENT: CREATE THE WEB REQUEST
                request = DirectCast(WebRequest.Create(strRequest), HttpWebRequest)
                ' *** COMMENT: GET THE RESPONSE
                response = DirectCast(request.GetResponse(), HttpWebResponse)
                ' *** COMMENT: GET THE RESPONSE STREAM INTO READER
                reader = New StreamReader(response.GetResponseStream())
                Do While reader.Peek <> -1
                    strJsonData = reader.ReadToEnd
                    Try
                        'Parse Simple Json Name / Value Pairs
                        Dim json_reader As New JsonTextReader(New StringReader(strJsonData))
                        While json_reader.Read()
                            If json_reader.Value IsNot Nothing Then
                                If json_reader.TokenType = JsonToken.PropertyName Then
                                    'Currency Code
                                    strCurrencyCode = json_reader.Value
                                    strCurrencyName = ""
                                    Console.WriteLine("{0}, {1}", json_reader.TokenType, json_reader.Value)
                                End If
                                If json_reader.TokenType = JsonToken.String Then
                                    'Currency Name
                                    strCurrencyName = json_reader.Value
                                    Console.WriteLine("{0}, {1}", json_reader.TokenType, json_reader.Value)
                                End If
                                'If strCurrencyCode <> "" And strCurrencyName <> "" Then
                                '    strResult += AddCurrency(strCurrencyCode, strCurrencyName) & vbCrLf
                                'End If
                            Else
                                'Json starting outer loop
                                Console.WriteLine("Json starting outer loop...")
                                'Console.WriteLine("TokenType: {0}", json_reader.TokenType)
                                'Console.WriteLine("ValueType: {0}", json_reader.ValueType)
                            End If
                            If strCurrencyCode <> "" And strCurrencyName <> "" Then
                                strResult += AddCurrency(strCurrencyCode, strCurrencyName) & vbCrLf
                                'Clear Variables
                                strCurrencyCode = ""
                                strCurrencyName = ""
                            End If
                        End While
                    Catch ex As Exception
                        Dim strError As String = ex.Message.ToString
                        strError = ex.Message.ToString & " Url:" & strRequest
                        Stop
                    End Try
                Loop
            Else
                'Log Error to tblExchangeRatesLog
                strResult = "No API AppID"
                fd.ReturnMsg = strRequest
                fd.RoutineName = "CRM_AddUpdateCurrencies"
                fd.StartTime = Date.Now
                fd.ErrorMsg = strResult
                fd.Success = False
                Update_ExchangeRatesLog(fd)
            End If
        Catch ex As Exception
            'Log Error to tblExchangeRatesLog
            strResult = ex.Message.ToString
            fd.ReturnMsg = strRequest
            fd.RoutineName = "CRM_AddUpdateCurrencies"
            fd.StartTime = Date.Now
            fd.ErrorMsg = strResult
            fd.Success = False
            Update_ExchangeRatesLog(fd)
        Finally
            If Not response Is Nothing Then response.Close()
        End Try
        Return strResult
    End Function

    Public Function GetLatestCurrencyExchangeRates() As String
        'Construct Url to get Current Exchange rates
        Dim strResult As String = ""
        Dim response As HttpWebResponse = Nothing
        Dim strJsonData As String = ""
        Dim strRequest As String = ""
        Dim intTimeStamp As Integer = 0
        Dim strCurrencyCode As String = ""
        Dim strExchangeDate As String = ""
        Dim dblExchangeRate As Decimal = 0
        Try
            Dim request As HttpWebRequest
            Dim reader As StreamReader
            If Len(strAppID) > 0 Then
                strRequest = "http://openexchangerates.org/api/latest.json"
                strRequest += "?app_id=" & strAppID
                ' *** COMMENT: CREATE THE WEB REQUEST
                request = DirectCast(WebRequest.Create(strRequest), HttpWebRequest)
                ' *** COMMENT: GET THE RESPONSE
                response = DirectCast(request.GetResponse(), HttpWebResponse)
                ' *** COMMENT: GET THE RESPONSE STREAM INTO READER
                reader = New StreamReader(response.GetResponseStream())
                Do While reader.Peek <> -1
                    strJsonData = reader.ReadToEnd
                    'TODO: Need to fix [ ] issue for rates array - ie: rates" & Chr(34) & " : " & "[{"   '}]}"

                    'Replace ": {" with ": [{"
                    strJsonData = strJsonData.Replace(": {", ": [{")
                    'Replace "}}" with "}]}"
                    strJsonData = strJsonData.Replace("}", "}]")
                    strJsonData = Mid(strJsonData, 1, Len(strJsonData) - 2) & "}"
                    Try
                        Dim item As Newtonsoft.Json.Linq.JObject
                        Dim jtoken As Newtonsoft.Json.Linq.JToken
                        Dim items As Newtonsoft.Json.Linq.JArray = Nothing
                        Dim jsonObject As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(strJsonData)
                        Dim objTimestamp As Object = DirectCast(jsonObject("timestamp"), Newtonsoft.Json.Linq.JValue).Value.ToString
                        intTimeStamp = CInt(Nz(objTimestamp, 0))
                        'Convert from Unix Timspan to DateTime
                        Dim time As DateTime = New DateTime(1970, 1, 1, 0, 0, 0)
                        Dim dtmExchangeDate As Date = time.AddSeconds(intTimeStamp)
                        'strExchangeDate = FormatDBDate(dtmDateTime)
                        'Parse Json rates into an array
                        items = DirectCast(jsonObject("rates"), Newtonsoft.Json.Linq.JArray)
                        'Loop through rows of array & Get CurrencyCodes & ExchangeRates
                        For i As Integer = 0 To items.Count - 1
                            item = DirectCast(items(i), Newtonsoft.Json.Linq.JObject)
                            jtoken = item.First
                            While jtoken IsNot Nothing
                                'Currency Code
                                strCurrencyCode = DirectCast(jtoken, Newtonsoft.Json.Linq.JProperty).Name.ToString()
                                'Exchange Rate
                                dblExchangeRate = CDbl(DirectCast(jtoken, Newtonsoft.Json.Linq.JProperty).Value.ToString)
                                'Add Record to tblExchangeRates
                                strResult = AddCurrencyExchangeRate(strCurrencyCode, dtmExchangeDate, dblExchangeRate)
                                'Next Record
                                jtoken = jtoken.[Next]
                            End While
                        Next
                    Catch ex As Exception
                        Dim strError As String = ex.Message.ToString
                        strError = ex.Message.ToString & " Url:" & strRequest
                        Stop
                    End Try
                Loop
            Else
                'Log Error to tblExchangeRatesLog
                strResult = "No API AppID"
                fd.ReturnMsg = strRequest
                fd.RoutineName = "CRM_ExchangeRates_Latest"
                fd.StartTime = Date.Now
                fd.ErrorMsg = strResult
                fd.Success = False
                Update_ExchangeRatesLog(fd)
            End If
        Catch ex As Exception
            'Log Error to tblExchangeRatesLog
            strResult = ex.Message.ToString
            fd.ReturnMsg = strRequest
            fd.RoutineName = "CRM_ExchangeRates_Latest"
            fd.StartTime = Date.Now
            fd.ErrorMsg = strResult
            fd.Success = False
            Update_ExchangeRatesLog(fd)
        Finally
            If Not response Is Nothing Then response.Close()
        End Try
        Return strResult
    End Function

    Public Sub Update_ExchangeRatesLog(ByVal fd As clsFunctionDetails)
        strSQL = "INSERT INTO tblExchangeRatesLog "
        strSQL += "(ElapsedTime, "
        strSQL += "EndTime, "
        strSQL += "ErrorMessage, "
        strSQL += "RecordsAdded, "
        strSQL += "RecordsDeleted, "
        strSQL += "RecordsExported, "
        strSQL += "RecordsUpdated, "
        strSQL += "ReturnMessage, "
        strSQL += "RoutineName, "
        strSQL += "RoutineType, "
        strSQL += "StartTime, "
        strSQL += "Success "
        strSQL += ") VALUES ("
        strSQL += "'" & fd.ElapsedTime & "', "
        strSQL += FormatDBDate(fd.EndTime) & ", "
        strSQL += "'" & Cs(fd.ErrorMsg) & "', "
        strSQL += fd.RecordsAdded & ", "
        strSQL += fd.RecordsDeleted & ", "
        strSQL += fd.RecordsExported & ", "
        strSQL += fd.RecordsUpdated & ", "
        strSQL += "'" & Cs(fd.ReturnMsg) & "', "
        strSQL += "'" & fd.RoutineName & "', "
        strSQL += "'" & fd.RoutineType & "', "
        strSQL += FormatDBDate(fd.StartTime) & ", "
        strSQL += FormatBit(fd.Success) & ")"
        AddRecord(strSQL, strConn)
    End Sub

    Public Sub UpdateOpportunityRecord(ByVal od As OpportunityDetails)
        Try
            Dim strSQL As String = ""
            strSQL = "UPDATE rptOpportunities "
            strSQL += "SET "
            strSQL += "Revenue = " & Nz(od.Revenue, 0) & ", "
            strSQL += "Updated = " & FormatDBDate(od.Updated) & ", "
            strSQL += "[Updated By] = '" & Cs(od.UpdatedBy) & "', "
            strSQL += "WHERE (OID = " & od.OID & ") "
            UpdateRecord(strSQL, strConn)
        Catch ex As Exception
            strReturnMsg = "Error-UpdateOpportunityReportRecord"
            strErrorMsg = ex.Message
            blnSuccess = False
            Console.WriteLine(strReturnMsg & " | " & strErrorMsg)
        End Try
    End Sub

    Public Function TestJson() As String
        Dim strJsonData As String = ""
        strJsonData += "{"
        strJsonData += Chr(34) & "disclaimer" & Chr(34) & " : " & Chr(34) & "Exchange rates" & Chr(34) & ","
        strJsonData += Chr(34) & "license" & Chr(34) & " : " & Chr(34) & "Data sourced" & Chr(34) & ","
        strJsonData += Chr(34) & "timestamp" & Chr(34) & " : " & 1354550408 & ","
        strJsonData += Chr(34) & "base" & Chr(34) & " : " & Chr(34) & "USD" & Chr(34) & ","
        strJsonData += Chr(34) & "rates" & Chr(34) & " : " & "[{"
        strJsonData += "    " & Chr(34) & "AED" & Chr(34) & " : " & 3.67282 & ","
        strJsonData += "    " & Chr(34) & "AFN" & Chr(34) & " : " & 52.039801 & ","
        strJsonData += "    " & Chr(34) & "SEK" & Chr(34) & " : " & 6.64219 & ","
        strJsonData += "    " & Chr(34) & "ZWL" & Chr(34) & " : " & 322.322775 & ","
        strJsonData += "    " & "}]"
        strJsonData += "}"
        '--------------------------------------------------------------------------------------------------
        Dim strResult As String = ""
        Dim intTimeStamp As Integer = 0
        Dim strCurrencyCode As String = ""
        Dim strExchangeDate As String = ""
        Dim dblExchangeRate As Decimal = 0
        Try
            Dim item As Newtonsoft.Json.Linq.JObject
            Dim jtoken As Newtonsoft.Json.Linq.JToken
            Dim items As Newtonsoft.Json.Linq.JArray = Nothing
            Dim jsonObject As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(strJsonData)
            Dim objTimestamp As Object = DirectCast(jsonObject("timestamp"), Newtonsoft.Json.Linq.JValue).Value.ToString
            intTimeStamp = CInt(Nz(objTimestamp, 0))
            'Convert from Unix Timspan to DateTime
            Dim time As DateTime = New DateTime(1970, 1, 1, 0, 0, 0)
            Dim dtmDateTime As Date = time.AddSeconds(intTimeStamp)
            strExchangeDate = FormatDBDate(dtmDateTime)
            'Parse Json rates into an array
            items = DirectCast(jsonObject("rates"), Newtonsoft.Json.Linq.JArray)
            'Loop through rows of array & Get CurrencyCode and ExchangeRate
            For i As Integer = 0 To items.Count - 1
                item = DirectCast(items(i), Newtonsoft.Json.Linq.JObject)
                jtoken = item.First
                While jtoken IsNot Nothing
                    'Currency Code
                    strCurrencyCode = DirectCast(jtoken, Newtonsoft.Json.Linq.JProperty).Name.ToString()
                    'Exchange Rate
                    dblExchangeRate = CDbl(DirectCast(jtoken, Newtonsoft.Json.Linq.JProperty).Value.ToString)
                    'Add Record to tblExchangeRates
                    strResult = AddCurrencyExchangeRate(strCurrencyCode, strExchangeDate, dblExchangeRate)
                    'Next Record
                    jtoken = jtoken.[Next]
                End While
            Next
        Catch ex As Exception
            Dim strError As String = ex.Message.ToString
            Stop
        End Try
        Return "Done"
    End Function

End Class
