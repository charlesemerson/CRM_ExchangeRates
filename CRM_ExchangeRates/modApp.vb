Imports System.Configuration
Imports System.Threading

Module modApp

#Region " Variables "
    'Classes
    Public fd As New clsFunctionDetails
    Public er As New clsExchangeRates
    'Email Variables
    Public strSMTPServer As String = ConfigurationManager.AppSettings("SMTPServer").ToString
    Public strSMTPPort As String = ConfigurationManager.AppSettings("SMTPPort").ToString
    Public strSMTPUser As String = ConfigurationManager.AppSettings("SMTPUser").ToString
    Public strSMTPPassword As String = ConfigurationManager.AppSettings("SMTPPassword").ToString
    Public strEmailFrom As String = ConfigurationManager.AppSettings("EmailFrom").ToString
    Public strEmailTo As String = ConfigurationManager.AppSettings("EmailTo").ToString
    Public strSubject As String = ""
    Public strBody As String = ""
    Public strMsg As String = ""
    Public emailDetails As New EmailHelper.clsEmail
    Public strReturn As String = ""
    Public dtmRunDate As Date = Date.Now
#End Region

#Region " Emailer "

    Public Sub SendCompletionEmail(ByVal fd As clsFunctionDetails)
        strSubject = "CRM Exchange Rates - Log"
        strMsg = "                    " & "Exchange Rates Finished.<br />"
        strMsg &= "                    " & "Start Time: " & fd.StartTime.ToString & "<br />"
        strMsg &= "                    " & "End Time: " & Date.Now.ToString & "<br />"
        strMsg &= "                    " & "Elapsed Time: " & fd.ElapsedTime & "<br />"
        strBody = "<html>" & vbCrLf
        strBody &= "    <body>" & vbCrLf
        strBody &= "        <table cellpadding='0' cellspacing = '4' border=0>" & vbCrLf
        strBody &= "            <tr>" & vbCrLf
        strBody &= "                <td>" & vbCrLf
        strBody &= strMsg
        strBody &= "                </td>" & vbCrLf
        strBody &= "            </tr>" & vbCrLf
        strBody &= "        </table>" & vbCrLf
        strBody &= "    </body>" & vbCrLf
        strBody &= "</html>"
        emailDetails = New EmailHelper.clsEmail
        With emailDetails
            .SMTPServer = strSMTPServer
            .SMTPUser = strSMTPUser
            .SMTPPassword = strSMTPPassword
            .EmailFrom = strEmailFrom
            .EmailTo = strEmailTo
            '.EmailCC = strEmailCC
            .Subject = strSubject
            .Body = strBody
        End With
        emailDetails.SendMailMessage()
        emailDetails = Nothing
    End Sub

    Public Sub SendErrorEmail(ByVal funcDetails As clsFunctionDetails)
        Dim strReturnMsg As String = ""
        Dim strSubject As String = "CRM Exchange Rates Error"
        strMsg = "                    " & funcDetails.RoutineName & "<br />"
        strMsg &= "                    " & funcDetails.ReturnMsg & "<br />"
        strMsg &= "                    " & "Date: " & funcDetails.StartTime.ToString & "<br />"
        strMsg &= "                    " & "Error Message: " & funcDetails.ErrorMsg & "<br />"
        strBody = "<html>" & vbCrLf
        strBody &= "    <body>" & vbCrLf
        strBody &= "        <table cellpadding='0' cellspacing = '4' border=0>" & vbCrLf
        strBody &= "            <tr>" & vbCrLf
        strBody &= "                <td>" & vbCrLf
        strBody &= strMsg
        strBody &= "                </td>" & vbCrLf
        strBody &= "            </tr>" & vbCrLf
        strBody &= "        </table>" & vbCrLf
        strBody &= "    </body>" & vbCrLf
        strBody &= "</html>"
        emailDetails = New EmailHelper.clsEmail
        With emailDetails
            .SMTPServer = strSMTPServer
            .SMTPUser = strSMTPUser
            .SMTPPassword = strSMTPPassword
            .EmailFrom = strEmailFrom
            .EmailTo = strEmailTo
            '.EmailCC = strEmailCC
            .Subject = strSubject
            .Body = strBody
        End With
        emailDetails.SendMailMessage()
        emailDetails = Nothing
    End Sub

#End Region

    Public Sub RunCRM_ExchangeRates(ByVal strStartDate As String)
        Dim er As New clsExchangeRates
        Dim strResult As String = ""
        er.ConnString = ConfigurationManager.ConnectionStrings("ConnString").ToString
        Dim dtmStartTime As DateTime = Date.Now
        '==============================================================================================
        Console.WriteLine("Updating Currency Exchange Rates via API: " & dtmStartTime.ToString)
        'Start Exchange Rates Update
        With fd
            .ReturnMsg = "Start CRM Exchange Rates Update..."
            .RoutineName = "CRM_ExchangeRates"
            .StartTime = dtmStartTime
            .Success = 1
        End With
        er.Update_ExchangeRatesLog(fd)
        'Run ExchangeRates Update
        '-------------------------------------------------------------------------------------------
        strResult = er.GetCurrencies()
        strResult = er.GetLatestCurrencyExchangeRates()
        'strResult = er.TestJson()

        'Update ExchangeRates Log
        Console.WriteLine(er.GetElapsedTime(dtmStartTime))
        With fd
            .ElapsedTime = er.GetElapsedTime(dtmStartTime)
            .EndTime = Date.Now
            .ReturnMsg = "CRM Exchange Rates Update Complete..." & vbCrLf & vbCrLf & strResult
            .RoutineName = "CRM_ExchangeRates"
            .RoutineType = ""
            .StartTime = dtmStartTime
            .Success = 1
        End With
        er.Update_ExchangeRatesLog(fd)
        fd = Nothing
    End Sub

    Sub Main()
        Dim strStartDate As String = ""
        RunCRM_ExchangeRates(strStartDate)
    End Sub

End Module
