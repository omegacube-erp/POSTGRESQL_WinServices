Imports System.Reflection
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports System.Net
Imports System.Net.Mail
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Text.RegularExpressions
Imports Microsoft.Exchange.WebServices
Imports Microsoft.Exchange.WebServices.Data
Imports System.IO
Imports System.Data
Imports System.Configuration
Imports System.Timers
Imports FE_SymmetricNamespace
Imports System.Threading
Imports System.Windows.Forms
Imports Npgsql
Public Class autodocumentdistribution_Service_pg

    Public sp As StreamWriter


    Dim da As NpgsqlDataAdapter
    Dim dbad As New NpgsqlDataAdapter
    Dim con As NpgsqlConnection
    Dim cmd As NpgsqlCommand

    Dim dbcon As NpgsqlConnection

    Dim pdfCreatedPath As String = String.Empty
    Dim ListSno As String = String.Empty

    Protected Shared AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
    Protected Shared strLogFilePath As String = AppPath + "\" + "Autodocumentdistribution_ExLog_pg.txt"
    Private Shared sw As StreamWriter = Nothing
    Dim scheduleTimer As New System.Timers.Timer()
    Dim WaitingTimer As New System.Timers.Timer()
    Public date_range As Date
    Dim str, str1, str2, str3, str4, str5, str6, str7, rname, wClause, sf, viewName, toemail, contactsQuery As String
    Dim sp1, sp2, sp3
    Dim reportPath As String = String.Empty
    Dim count As Integer = 0
    Dim nSchemas As Integer = 0
    Dim pkValue As String = String.Empty

    Public hours1 As String
    Public mins1 As String
    Public eCheckFreq As Integer
    Public hCount As Integer = 0

    '' Exchange Variables
    Public esb As New ExchangeService(ExchangeVersion.Exchange2007_SP1)
    Public hostName As String = String.Empty
    Public fltCondition As String = String.Empty
    Public domainName As String = String.Empty
    Public EmailUser As String = String.Empty
    Public EmailPassword As String = String.Empty

    Public sqlQuery As String = String.Empty

    Public strSNO As String = String.Empty
    Public strTO_ADDRESS As String = String.Empty
    Public strCC_ADDRESS As String = String.Empty
    Public strBCC_ADDRESS As String = String.Empty
    Public strMAIL_SUBJECT As String = String.Empty
    Public strATTACHMENT As String = String.Empty
    Public strREPORT_NAME As String = String.Empty
    Public strVIEW_NAME As String = String.Empty
    Public strCONDITION_PARAMETER As String = String.Empty
    Public strCONDITION_VALUE As String = String.Empty
    Public strSEND_CONFIRMATION As String = String.Empty
    Public strCREATED_BY As String = String.Empty
    Public strCREATED_DATE As String = String.Empty
    Public strCHANGED_BY As String = String.Empty
    Public strCHANGED_DATE As String = String.Empty
    Public strNEWS_LETTER As String = String.Empty
    Public strREPORT_FILE_NAME As String = String.Empty
    Public strFROM_ID As String = String.Empty
    Public strMAIL_BODY As String = String.Empty
    Public strSOURCE_NAME As String = String.Empty
    Public strSOURCE_TYPE As String = String.Empty

    Public strPreserveLog As String = String.Empty

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            scheduleTimer.Enabled = False
            strPreserveLog = ConfigurationManager.AppSettings("PreserveLog").ToString()
            If String.IsNullOrEmpty(strPreserveLog) Then
                strPreserveLog = "N"
            End If

            LogExceptions(strLogFilePath, Nothing, "Step 1")
            nSchemas = Convert.ToInt16(ConfigurationManager.AppSettings("NoOfSchemas").ToString())

            hours1 = ConfigurationManager.AppSettings("Hours").ToString()
            mins1 = ConfigurationManager.AppSettings("Mins").ToString()
            eCheckFreq = Convert.ToInt16(ConfigurationManager.AppSettings("EmailCheckFrequency").ToString())
            If eCheckFreq = 0 Then
                eCheckFreq = 1
            End If

            LogExceptions(strLogFilePath, Nothing, "Step 2")
            scheduleTimer.Interval = eCheckFreq * 60 * 1000
            AddHandler scheduleTimer.Elapsed, New ElapsedEventHandler(AddressOf TimerElapsed)

            LogExceptions(strLogFilePath, Nothing, "Started")
            scheduleTimer.Enabled = True

        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub
    Private Sub TimerElapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        Try
            scheduleTimer.Enabled = False

            hours1 = ConfigurationManager.AppSettings("Hours").ToString()
            mins1 = ConfigurationManager.AppSettings("Mins").ToString()

            Try
                Call auto_document_distribution_ForMultiSchemas()
            Catch ex As Exception
                LogExceptions(strLogFilePath, ex, "TimerElapsed - auto_document_distribution_ForMultiSchemas")
            End Try

            scheduleTimer.Enabled = True
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub
    Public Sub LogExceptions(ByVal filePath As String, Optional ByVal ex As Exception = Nothing, Optional ByVal msg As String = Nothing)
        If File.Exists(filePath) Then
            If strPreserveLog = "N" Then
                File.Delete(filePath)
            End If
        End If
        If False = File.Exists(filePath) Then
            Dim fs As New FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
            fs.Close()
        End If
        WriteExceptionLog(filePath, ex, msg)
    End Sub
    Private Sub WriteExceptionLog(ByVal strPathName As String, Optional ByVal objException As Exception = Nothing, Optional ByVal msg As String = Nothing)
        If (Not objException Is Nothing) AndAlso (Not String.IsNullOrEmpty(msg)) Then
            sw = New StreamWriter(strPathName, True)
            sw.WriteLine("Source     :" & objException.Source.ToString().Trim())
            sw.WriteLine("Method     : " & objException.TargetSite.Name.ToString())
            sw.WriteLine("Date       : " & DateTime.Now.ToLongTimeString())
            sw.WriteLine("Time       : " & DateTime.Now.ToShortDateString())
            sw.WriteLine("Error      : " & objException.Message.ToString().Trim())
            sw.WriteLine("Stack Trace: " & objException.StackTrace.ToString().Trim())
            sw.WriteLine("^^-------------------------------------------------------------------^^")
            sw.WriteLine(msg)
            sw.WriteLine("^^-------------------------------------------------------------------^^")
        ElseIf String.IsNullOrEmpty(msg) Then
            sw = New StreamWriter(strPathName, True)
            sw.WriteLine("Source     :" & objException.Source.ToString().Trim())
            sw.WriteLine("Method     : " & objException.TargetSite.Name.ToString())
            sw.WriteLine("Date       : " & DateTime.Now.ToLongTimeString())
            sw.WriteLine("Time       : " & DateTime.Now.ToShortDateString())
            sw.WriteLine("Error      : " & objException.Message.ToString().Trim())
            sw.WriteLine("Stack Trace: " & objException.StackTrace.ToString().Trim())
            sw.WriteLine("^^-------------------------------------------------------------------^^")
        Else
            sw = New StreamWriter(strPathName, True)
            sw.WriteLine(msg)
        End If
        sw.Flush()
        sw.Close()
    End Sub
    'For Auto Document Distrubution
    Public Sub auto_document_distribution_ForMultiSchemas()
        Dim strServerName As String = String.Empty
        Dim strUserName As String = String.Empty
        Dim strPassWord As String = String.Empty
        Dim rep_path As String = String.Empty
        Dim attachment_path As String = String.Empty
        Dim Host_Url As String = String.Empty
        Try
            If nSchemas = 0 Then
                nSchemas = 10
            End If

            For i As Integer = 1 To nSchemas

                strServerName = String.Empty
                strUserName = String.Empty
                strPassWord = String.Empty
                rep_path = String.Empty
                attachment_path = String.Empty

                strServerName = ConfigurationManager.AppSettings("Server" & i)
                strUserName = "x" 'ConfigurationManager.AppSettings("User" & i)
                strPassWord = "Y" 'ConfigurationManager.AppSettings("Pwd" & i)
                rep_path = ConfigurationManager.AppSettings("REPX_PATH" & i)
                attachment_path = ConfigurationManager.AppSettings("ATTACHMENTS_FOLDER" & i)
                Host_Url = ConfigurationManager.AppSettings("REPORT_URL" & i)
                Try
                    If Not String.IsNullOrEmpty(strServerName) Then
                        Try
                            If Not String.IsNullOrEmpty(strUserName) Then
                                Try
                                    If Not String.IsNullOrEmpty(strPassWord) Then
                                        Try
                                            Dim dbad As New NpgsqlDataAdapter
                                            dbcon = New NpgsqlConnection(strServerName)
                                            dbad.SelectCommand = New NpgsqlCommand
                                            dbad.SelectCommand.Connection = dbcon
                                            If (dbad.SelectCommand.Connection.State = ConnectionState.Closed) Then
                                                dbad.SelectCommand.Connection.Open()
                                            End If
                                            dbad.SelectCommand.Connection.Close()
                                            Call auto_document_distribution(strServerName, strUserName, strPassWord, rep_path, attachment_path, Host_Url)
                                        Catch ex As Exception
                                            Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
                                            LogExceptions(strLogFilePath, ex, Nothing)
                                            Continue For
                                        End Try
                                    End If
                                Catch ex As Exception
                                    Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
                                    LogExceptions(strLogFilePath, ex, Nothing)
                                    Continue For
                                End Try
                            End If
                        Catch ex As Exception
                            Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
                            LogExceptions(strLogFilePath, ex, Nothing)
                            Continue For
                        End Try
                    End If
                Catch ex As Exception
                    Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
                    LogExceptions(strLogFilePath, ex, Nothing)
                    Continue For
                End Try
            Next
        Catch ex As Exception
            Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub
    Public Sub Insert_Sys_Log(ByVal message As String, ByVal strServerName1 As String, ByVal strUserName1 As String, ByVal strPassWord1 As String)
        Try
            dbcon = New NpgsqlConnection(strServerName1)
            Dim sterr1, sterr2, sterr3, sterr4, sterr As String
            sterr = Replace(message, "'", "''")
            If (Len(sterr) > 4000) Then
                sterr1 = Mid(sterr, 1, 4000)
                If (Len(sterr) > 8000) Then
                    sterr2 = Mid(sterr, 4000, 8000)
                    If (Len(sterr) > 12000) Then
                        sterr3 = Mid(sterr, 8000, 12000)
                        If (Len(sterr) > 16000) Then
                            sterr4 = Mid(sterr, 12000, 16000)
                        Else
                            sterr4 = Mid(sterr, 12000, Len(sterr))
                        End If
                    Else
                        sterr3 = Mid(sterr, 8000, Len(sterr))
                        sterr4 = ""
                    End If
                Else
                    sterr2 = Mid(sterr, 4000, Len(sterr))
                    sterr3 = ""
                    sterr3 = ""
                    sterr4 = ""
                End If
            Else
                sterr1 = sterr
                sterr2 = ""
                sterr3 = ""
                sterr4 = ""
            End If
            dbad.InsertCommand = New NpgsqlCommand
            dbad.InsertCommand.Connection = dbcon
            dbad.InsertCommand.CommandText = "Insert into SYS_ACTIVATE_STATUS_LOG (LINE_NO, CHANGE_REQUEST_NO,  OBJECT_TYPE, OBJECT_NAME, ERROR_TEXT, STATUS,LOG_DATE,ERROR_TEXT1, ERROR_TEXT2, ERROR_TEXT3) values ((select COALESCE(max(line_no),0)+1 from SYS_ACTIVATE_STATUS_LOG),'EDGE','SERVICE','EDGE_QUERY','" & sterr1 & "','N',now(),'" & sterr2 & "','" & sterr3 & "','" & sterr4 & "')"
            If (dbad.InsertCommand.Connection.State = ConnectionState.Closed) Then
                dbad.InsertCommand.Connection.Open()
            End If
            dbad.InsertCommand.ExecuteNonQuery()
            If (dbad.InsertCommand.Connection.State = ConnectionState.Open) Then
                dbad.InsertCommand.Connection.Close()
            End If
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, "Insert_Sys_Log")
        End Try
    End Sub
    Public Sub auto_document_distribution(ByVal strServerName1 As String, ByVal strUserName1 As String, ByVal strPassWord1 As String, ByVal repPath1 As String, ByVal pdfPath1 As String, ByVal Host_Url As String)
        Try
            Dim dbad As New NpgsqlDataAdapter
            dbcon = New NpgsqlConnection(strServerName1)

            dbad.SelectCommand = New NpgsqlCommand
            dbad.InsertCommand = New NpgsqlCommand
            dbad.UpdateCommand = New NpgsqlCommand

            dbad.SelectCommand.Connection = dbcon
            dbad.InsertCommand.Connection = dbcon
            dbad.UpdateCommand.Connection = dbcon


            Dim ds As New DataSet
            Dim i As Integer
            dbad.SelectCommand.CommandText = "SELECT SETUP_ID, SETUP_DESCRIPTION, SETUP_TYPE, SETUP_FREQUENCY, REPEAT_HOURS, TO_FIELD_QUERY, CC_FIELD_QUERY, BCC_FIELD_QUERY, MAIL_SUBJECT, MAIL_BODY, SETUP_QUERY, REPORT_NAME, REPORT_VIEW_NAME, PROGRAM_INPUT1, PROGRAM_INPUT2, PROGRAM_INPUT3, PROGRAM_INPUT4, PROGRAM_INPUT5, PROGRAM_INPUT6, PROGRAM_INPUT_TYPE1, PROGRAM_INPUT_TYPE2, PROGRAM_INPUT_TYPE3, PROGRAM_INPUT_TYPE4, PROGRAM_INPUT_TYPE5, PROGRAM_INPUT_TYPE6, MAP_INPUT1, MAP_INPUT2, MAP_INPUT3, MAP_INPUT4, MAP_INPUT5, MAP_INPUT6, ACTIVE_FLAG, DEFAULT_PRINTER_NAME, UPDATE_QUERY, START_DATE, NEXT_EXECUTION_DATE FROM   SYS_AUTO_DOCUMENT_DISTRIBUTION WHERE ACTIVE_FLAG='Y' AND COALESCE(NEXT_EXECUTION_DATE,now())<=now() AND COALESCE(START_DATE,now())<=now()"
            ds.Clear()
            dbad.Fill(ds)
            LogExceptions(strLogFilePath, Nothing, "Step 3 Count: " & ds.Tables(0).Rows.Count)
            If (ds.Tables(0).Rows.Count > 0) Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If Not (Equals(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"), System.DBNull.Value)) Then
                        Try
                            pkValue = String.Empty
                            pkValue = "Auto_document_distribution - Data Source=" + strServerName1 + ";User Id=" + strUserName1 + "  SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"))) = "HOURLY") Then
                                Dim hr As Double
                                hr = 1
                                If Not (Equals(ds.Tables(0).Rows(i)("REPEAT_HOURS"), System.DBNull.Value)) Then
                                    If (IsNumeric(ds.Tables(0).Rows(i)("REPEAT_HOURS"))) Then
                                        hr = CDbl(ds.Tables(0).Rows(i)("REPEAT_HOURS"))
                                    End If
                                End If
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "UPDATE SYS_AUTO_DOCUMENT_DISTRIBUTION SET NEXT_EXECUTION_DATE=now() + interval '1' hour*" & hr & " WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log("auto_document_distribution1 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                End Try
                            End If
                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"))) = "DAILY") Then
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "UPDATE SYS_AUTO_DOCUMENT_DISTRIBUTION SET NEXT_EXECUTION_DATE=now() + interval '1' day*1 WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log("auto_document_distribution2 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                End Try
                            End If

                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"))) = "WEEKLY") Then
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "UPDATE SYS_AUTO_DOCUMENT_DISTRIBUTION SET NEXT_EXECUTION_DATE=now() + interval '1' day*7 WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log("auto_document_distribution3 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                End Try
                            End If

                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"))) = "MONTHLY") Then
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "UPDATE SYS_AUTO_DOCUMENT_DISTRIBUTION SET NEXT_EXECUTION_DATE=now() + interval '1' day*30 WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log("auto_document_distribution4 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                End Try
                            End If

                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"))) = "QUARTERLY") Then
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "UPDATE SYS_AUTO_DOCUMENT_DISTRIBUTION SET NEXT_EXECUTION_DATE=now() + interval '1' day*91 WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log("auto_document_distribution5 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                End Try

                            End If

                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"))) = "HALF YEARLY") Then
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "UPDATE SYS_AUTO_DOCUMENT_DISTRIBUTION SET NEXT_EXECUTION_DATE=now() + interval '1' day*182 WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log("auto_document_distribution6 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                End Try

                            End If

                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_FREQUENCY"))) = "YEARLY") Then
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "UPDATE SYS_AUTO_DOCUMENT_DISTRIBUTION SET NEXT_EXECUTION_DATE=now() + interval '1' day*365 WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log("auto_document_distribution7 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                End Try

                            End If
                            LogExceptions(strLogFilePath, Nothing, "Step 4 (Setup Query): " & (ds.Tables(0).Rows(i)("SETUP_QUERY")).replace("'", "''"))
                            If (1 = 1) Then

                                Dim dsseq As New DataSet
                                dbad.SelectCommand.CommandText = "SELECT  SETUP_QUERY FROM SYS_AUTO_DOCUMENT_DISTRIBUTION WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                dsseq.Clear()
                                dbad.Fill(dsseq)

                                Dim dquery As New DataSet
                                Dim i1, i2 As Integer
                                Dim to_email, cc_email, bcc_email, to_email1, cc_email1, bcc_email1, subject1, body1, sv As String
                                to_email = String.Empty
                                cc_email = String.Empty
                                bcc_email = String.Empty
                                to_email1 = String.Empty
                                cc_email1 = String.Empty
                                bcc_email1 = String.Empty
                                subject1 = String.Empty
                                body1 = String.Empty
                                sv = String.Empty
                                If (dsseq.Tables(0).Rows.Count > 0) Then
                                    dbad.SelectCommand.CommandText = dsseq.Tables(0).Rows(0)("SETUP_QUERY")
                                Else
                                    dbad.SelectCommand.CommandText = ds.Tables(0).Rows(i)("SETUP_QUERY")
                                End If

                                dquery.Clear()
                                dbad.Fill(dquery)
                                If (dquery.Tables(0).Rows.Count > 0) Then

                                    For i1 = 0 To dquery.Tables(0).Rows.Count - 1
                                        If Not (Equals(ds.Tables(0).Rows(i)("SETUP_TYPE"), System.DBNull.Value)) Then
                                            If (UCase(Trim(ds.Tables(0).Rows(i)("SETUP_TYPE"))) = "EMAIL") Then
                                                If Not (Equals(ds.Tables(0).Rows(i)("TO_FIELD_QUERY"), System.DBNull.Value)) Then
                                                    to_email = ds.Tables(0).Rows(i)("TO_FIELD_QUERY")
                                                Else
                                                    to_email = ""
                                                End If
                                                If Not (Equals(ds.Tables(0).Rows(i)("CC_FIELD_QUERY"), System.DBNull.Value)) Then
                                                    cc_email = ds.Tables(0).Rows(i)("CC_FIELD_QUERY")
                                                Else
                                                    cc_email = ""
                                                End If
                                                If Not (Equals(ds.Tables(0).Rows(i)("BCC_FIELD_QUERY"), System.DBNull.Value)) Then
                                                    bcc_email = ds.Tables(0).Rows(i)("BCC_FIELD_QUERY")
                                                Else
                                                    bcc_email = ""
                                                End If
                                                If Not (Equals(ds.Tables(0).Rows(i)("MAIL_SUBJECT"), System.DBNull.Value)) Then
                                                    subject1 = ds.Tables(0).Rows(i)("MAIL_SUBJECT")
                                                Else
                                                    subject1 = ""
                                                End If

                                                Dim dbody As New DataSet
                                                dbad.SelectCommand.CommandText = "SELECT  MAIL_BODY FROM SYS_AUTO_DOCUMENT_DISTRIBUTION WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                                dbody.Clear()
                                                dbad.Fill(dbody)
                                                If (dbody.Tables(0).Rows.Count > 0) Then
                                                    If Not (Equals(dbody.Tables(0).Rows(0)("MAIL_BODY"), System.DBNull.Value)) Then
                                                        body1 = dbody.Tables(0).Rows(0)("MAIL_BODY")
                                                    Else
                                                        body1 = ""
                                                    End If
                                                Else
                                                    If Not (Equals(ds.Tables(0).Rows(i)("MAIL_BODY"), System.DBNull.Value)) Then
                                                        body1 = Replace(ds.Tables(0).Rows(i)("MAIL_BODY"), "'", "''")
                                                    Else
                                                        body1 = ""
                                                    End If
                                                End If

                                                Dim uqry As New DataSet
                                                Dim ustr As String
                                                ustr = ""
                                                dbad.SelectCommand.CommandText = "SELECT  UPDATE_QUERY FROM SYS_AUTO_DOCUMENT_DISTRIBUTION WHERE SETUP_ID='" & ds.Tables(0).Rows(i)("SETUP_ID") & "'"
                                                uqry.Clear()
                                                dbad.Fill(uqry)
                                                If (uqry.Tables(0).Rows.Count > 0) Then
                                                    If Not (Equals(uqry.Tables(0).Rows(0)("UPDATE_QUERY"), System.DBNull.Value)) Then
                                                        ustr = uqry.Tables(0).Rows(0)("UPDATE_QUERY")
                                                    Else
                                                        ustr = ""
                                                    End If
                                                End If

                                                For i2 = 0 To dquery.Tables(0).Columns.Count - 1

                                                    If Not (Equals(dquery.Tables(0).Rows(i1)(dquery.Tables(0).Columns(i2).ColumnName), System.DBNull.Value)) Then

                                                        to_email = Replace(to_email, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", dquery.Tables(0).Rows(i1)(dquery.Tables(0).Columns(i2).ColumnName))
                                                        cc_email = Replace(cc_email, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", dquery.Tables(0).Rows(i1)(dquery.Tables(0).Columns(i2).ColumnName))
                                                        bcc_email = Replace(bcc_email, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", dquery.Tables(0).Rows(i1)(dquery.Tables(0).Columns(i2).ColumnName))
                                                        subject1 = Replace(subject1, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", dquery.Tables(0).Rows(i1)(dquery.Tables(0).Columns(i2).ColumnName))
                                                        body1 = Replace(body1, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", dquery.Tables(0).Rows(i1)(dquery.Tables(0).Columns(i2).ColumnName))
                                                        ustr = Replace(ustr, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", dquery.Tables(0).Rows(i1)(dquery.Tables(0).Columns(i2).ColumnName))
                                                    Else
                                                        to_email = Replace(to_email, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", "")
                                                        cc_email = Replace(cc_email, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", "")
                                                        bcc_email = Replace(bcc_email, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", "")
                                                        subject1 = Replace(subject1, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", "")
                                                        body1 = Replace(body1, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", "")
                                                        ustr = Replace(ustr, "{[" & dquery.Tables(0).Columns(i2).ColumnName & "]}", "")
                                                    End If
                                                Next


                                                Try
                                                    If (ustr <> "") Then
                                                        If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                                            dbad.UpdateCommand.Connection.Open()
                                                        End If
                                                        dbad.UpdateCommand.CommandText = ustr
                                                        dbad.UpdateCommand.ExecuteNonQuery()
                                                    End If

                                                Catch ex As Exception
                                                    Call Insert_Sys_Log("auto_document_distribution8 " & ustr, strServerName1, strUserName1, strPassWord1)
                                                End Try

                                                If (to_email <> "") Then
                                                    Try
                                                        Dim ds_to As New DataSet
                                                        dbad.SelectCommand.CommandText = to_email
                                                        ds_to.Clear()
                                                        dbad.Fill(ds_to)
                                                        If (ds_to.Tables(0).Rows.Count > 0) Then

                                                            to_email1 = ds_to.Tables(0).Rows(0)(0)
                                                            'cc
                                                            If (cc_email <> "") Then
                                                                Dim ds_cc As New DataSet
                                                                dbad.SelectCommand.CommandText = cc_email
                                                                ds_cc.Clear()
                                                                dbad.Fill(ds_cc)
                                                                If (ds_cc.Tables(0).Rows.Count > 0) Then
                                                                    cc_email1 = ds_cc.Tables(0).Rows(0)(0)
                                                                Else
                                                                    cc_email1 = ""
                                                                End If
                                                            End If
                                                            'bcc
                                                            If (bcc_email <> "") Then
                                                                Dim ds_bcc As New DataSet
                                                                dbad.SelectCommand.CommandText = bcc_email
                                                                ds_bcc.Clear()
                                                                dbad.Fill(ds_bcc)
                                                                If (ds_bcc.Tables(0).Rows.Count > 0) Then
                                                                    bcc_email1 = ds_bcc.Tables(0).Rows(0)(0)
                                                                Else
                                                                    bcc_email1 = ""
                                                                End If
                                                            End If

                                                            Dim sno As String
                                                            Dim dbmax As New DataSet
                                                            dbad.SelectCommand.CommandText = "select max(to_number(sno,'99999999')) + 1 as newSNo from sys_generic_email"
                                                            dbmax.Clear()
                                                            dbad.Fill(dbmax)
                                                            If (dbmax.Tables(0).Rows.Count > 0) Then
                                                                sno = dbmax.Tables(0).Rows(0)(0)

                                                                If (dbad.InsertCommand.Connection.State = ConnectionState.Closed) Then
                                                                    dbad.InsertCommand.Connection.Open()
                                                                End If
                                                                dbad.InsertCommand.CommandText = "Insert into SYS_GENERIC_EMAIL (SNO, TO_ADDRESS,CC_ADDRESS, BCC_ADDRESS, MAIL_SUBJECT, SEND_CONFIRMATION, CREATED_BY, CREATED_DATE, NEWS_LETTER, SOURCE_NAME, SOURCE_TYPE,MAIL_BODY) values('" & sno & "','" & Replace(to_email1, "'", "''") & "','" & Replace(cc_email1, "'", "''") & "','" & Replace(bcc_email1, "'", "''") & "','" & Replace(subject1, "'", "''") & "','K','MURALIT',now(),'N','AUTO_DOCUMENT_DISTRIBUTION','ALERT','" & body1 & "')"
                                                                dbad.InsertCommand.ExecuteNonQuery()
                                                                'update_clob(body1, "UPDATE SYS_GENERIC_EMAIL Set MAIL_BODY=:TEXT_DATA WHERE SNO='" + sno + "'", strServerName1, strUserName1, strPassWord1)

                                                                str2 = String.Empty
                                                                Try


                                                                    For i3 = 1 To 6
                                                                        sv = String.Empty
                                                                        sv = "MAP_INPUT" & i3
                                                                        If Not (Equals(ds.Tables(0).Rows(i)(sv), System.DBNull.Value)) Then
                                                                            If (ds.Tables(0).Rows(i)(sv) <> "") Then
                                                                                If String.IsNullOrEmpty(str2) Then
                                                                                    str2 = ds.Tables(0).Rows(i)(sv)
                                                                                Else
                                                                                    str2 = String.Concat(str2, "[_]", ds.Tables(0).Rows(i)(sv))
                                                                                End If
                                                                            End If

                                                                        End If
                                                                    Next
                                                                Catch ex As Exception

                                                                End Try
                                                                If Not str2.Contains("[_]") Then
                                                                    str2 = String.Concat(str2, "[_]")
                                                                End If

                                                                str4 = String.Empty
                                                                Try


                                                                    For i3 = 1 To 6
                                                                        sv = "PROGRAM_INPUT" & i3
                                                                        If Not (Equals(ds.Tables(0).Rows(i)(sv), System.DBNull.Value)) Then
                                                                            If (ds.Tables(0).Rows(i)(sv) <> "") Then
                                                                                If Not (Equals(dquery.Tables(0).Rows(i1)(ds.Tables(0).Rows(i)(sv)), System.DBNull.Value)) Then
                                                                                    If String.IsNullOrEmpty(str4) Then
                                                                                        str4 = dquery.Tables(0).Rows(i1)(ds.Tables(0).Rows(i)(sv))
                                                                                    Else
                                                                                        str4 = String.Concat(str4, "[_]", dquery.Tables(0).Rows(i1)(ds.Tables(0).Rows(0)(sv)))
                                                                                    End If
                                                                                End If
                                                                            End If

                                                                        End If
                                                                    Next
                                                                Catch ex As Exception

                                                                End Try
                                                                If Not str4.Contains("[_]") Then
                                                                    str4 = String.Concat(str4, "[_]")
                                                                End If

                                                                If Not (Equals(ds.Tables(0).Rows(i)("REPORT_NAME"), System.DBNull.Value)) Then
                                                                    If (ds.Tables(0).Rows(i)("REPORT_NAME") <> "") Then

                                                                        If ds.Tables(0).Rows(i)("REPORT_NAME").ToString().Contains(".repx") Then

                                                                            Try
                                                                                '  pdfCreatedPath = GenerateDevxReport(sno, ds.Tables(0).Rows(i)("REPORT_NAME"), str2, str4, strServerName1, strUserName1, strPassWord1, repPath1, pdfPath1)
                                                                                pdfCreatedPath = GenerateDevxReport(sno, ds.Tables(0).Rows(i)("REPORT_NAME"), str2, str4, strServerName1, strUserName1, strPassWord1, repPath1, pdfPath1, Host_Url)
                                                                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                                                                    dbad.UpdateCommand.Connection.Open()
                                                                                End If
                                                                                If File.Exists(pdfCreatedPath) Then
                                                                                    dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N',ATTACHMENT='" & pdfCreatedPath & "'  WHERE SNO='" & sno & "'"
                                                                                Else
                                                                                    dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='R' WHERE SNO='" & sno & "'"
                                                                                End If
                                                                                dbad.UpdateCommand.ExecuteNonQuery()
                                                                                Thread.Sleep(60000)
                                                                            Catch ex As Exception
                                                                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                                                                    dbad.UpdateCommand.Connection.Open()
                                                                                End If
                                                                                dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N' WHERE SNO='" & sno & "'"
                                                                                dbad.UpdateCommand.ExecuteNonQuery()
                                                                            End Try
                                                                        Else
                                                                            'New logic for list report creation
                                                                            Try

                                                                                Dim report_ID = String.Empty
                                                                                If Not (Equals(ds.Tables(0).Rows(i)("REPORT_NAME"), System.DBNull.Value)) Then
                                                                                    report_ID = ds.Tables(0).Rows(i)("REPORT_NAME").ToString.ToUpper
                                                                                    report_ID = report_ID.Replace(".ASPX", "")
                                                                                End If
                                                                                LogExceptions(strLogFilePath, Nothing, "Came to list report")
                                                                                Dim directoryString As String = pdfPath1 & "\" & sno
                                                                                LogExceptions(strLogFilePath, Nothing, "Attachment Path {App Config} :" & pdfPath1)
                                                                                LogExceptions(strLogFilePath, Nothing, "Email Folder :" & sno)
                                                                                LogExceptions(strLogFilePath, Nothing, "directoryString :" & directoryString)
                                                                                LogExceptions(strLogFilePath, Nothing, "rname :" & report_ID)
                                                                                If Not Directory.Exists(directoryString) Then
                                                                                    Directory.CreateDirectory(directoryString)
                                                                                End If
                                                                                ' SAVE THE CURRENT FOLDER PATH
                                                                                'Dim sp4 = Split(str4, "[_]")
                                                                                Dim currentFileName As String = (report_ID.Replace(".ASPX", "") & " ").Trim & ".PDF"
                                                                                reportPath = directoryString & "\" '& report_ID
                                                                                LogExceptions(strLogFilePath, Nothing, "reportPath :" & reportPath)
                                                                                'Dim REPORT_URL As String = String.Empty
                                                                                'REPORT_URL = ConfigurationManager.AppSettings("REPORT_URL")

                                                                                Dim list_report_name As String
                                                                                list_report_name = report_ID & "_PRINT.aspx"
                                                                                LogExceptions(strLogFilePath, Nothing, "Came to list report 1")
                                                                                Dim url As String

                                                                                url = Host_Url & "" & list_report_name & "?REPORT_NAME=" & report_ID & "&SPath=" & reportPath & "&Type=PDF&PF=PL"

                                                                                LogExceptions(strLogFilePath, Nothing, url)

                                                                                Dim thread As New Thread(New ThreadStart(Sub()
                                                                                                                             Dim WebBrowser1 As New WebBrowser()
                                                                                                                             WebBrowser1.Navigate(url)
                                                                                                                             LogExceptions(strLogFilePath, Nothing, "Appp 1111")

                                                                                                                             While WebBrowser1.ReadyState <> WebBrowserReadyState.Complete
                                                                                                                                 Application.DoEvents()
                                                                                                                             End While
                                                                                                                             WebBrowser1.Dispose()
                                                                                                                             LogExceptions(strLogFilePath, Nothing, "Appp 2222")
                                                                                                                         End Sub))
                                                                                thread.SetApartmentState(ApartmentState.STA)
                                                                                thread.Start()
                                                                                thread.Join()
                                                                                Thread.Sleep(60000)
                                                                                pdfCreatedPath = reportPath & report_ID & ".pdf"
                                                                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                                                                    dbad.UpdateCommand.Connection.Open()
                                                                                End If
                                                                                LogExceptions(strLogFilePath, Nothing, "Pdf path : " & pdfCreatedPath)
                                                                                'If File.Exists(pdfCreatedPath) Then
                                                                                LogExceptions(strLogFilePath, Nothing, "Came to list report 3")
                                                                                dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N',ATTACHMENT='" & pdfCreatedPath & "'  WHERE SNO='" & sno & "'"
                                                                                'Else
                                                                                '    LogExceptions(strLogFilePath, Nothing, "Came to list report 4")
                                                                                '    dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N' WHERE SNO='" & sno & "'"
                                                                                'End If
                                                                                dbad.UpdateCommand.ExecuteNonQuery()
                                                                            Catch ex As Exception
                                                                                LogExceptions(strLogFilePath, Nothing, "Came to Error")
                                                                            If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                                                                dbad.UpdateCommand.Connection.Open()
                                                                            End If
                                                                            dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N' WHERE SNO='" & sno & "'"
                                                                            dbad.UpdateCommand.ExecuteNonQuery()
                                                                        End Try
                                                                    End If
                                                                Else
                                                                    If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                                                        dbad.UpdateCommand.Connection.Open()
                                                                    End If
                                                                    dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N' WHERE SNO='" & sno & "'"
                                                                    dbad.UpdateCommand.ExecuteNonQuery()
                                                                End If


                                                                Else
                                                                    If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                                                        dbad.UpdateCommand.Connection.Open()
                                                                    End If
                                                                    dbad.UpdateCommand.CommandText = "UPDATE SYS_GENERIC_EMAIL SET SEND_CONFIRMATION='N' WHERE SNO='" & sno & "'"
                                                                    dbad.UpdateCommand.ExecuteNonQuery()
                                                                End If

                                                            End If
                                                        Else
                                                            to_email1 = ""
                                                        End If

                                                    Catch ex As Exception
                                                        LogExceptions(strLogFilePath, Nothing, pkValue)
                                                        Call Insert_Sys_Log("auto_document_distribution9 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                                                        LogExceptions(strLogFilePath, ex, pkValue)
                                                    End Try
                                                End If
                                            End If

                                        End If
                                    Next
                                End If
                            End If

                        Catch ex As Exception
                            'eventLogEQueryEx.Clear()
                            'eventLogEQueryEx.WriteEntry(pkValue.ToString)
                            'eventLogEQueryEx.WriteEntry(ex.Message.ToString)
                            Call Insert_Sys_Log("auto_document_distribution10 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                            LogExceptions(strLogFilePath, ex, pkValue)
                        End Try
                    End If
                Next
            End If
            dbad.SelectCommand.Connection.Close()
            dbad.InsertCommand.Connection.Close()
        Catch ex As Exception
            'eventLogEQueryEx.Clear()
            'eventLogEQueryEx.WriteEntry(pkValue.ToString)
            'eventLogEQueryEx.WriteEntry(ex.Message.ToString)
            Call Insert_Sys_Log("auto_document_distribution11 " & ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub



    Public Function GenerateDevxReport(ByVal _newSNo As String, ByVal rname As String, ByVal s1 As String, ByVal s2 As String, ByVal strServerName1 As String, ByVal strUserName1 As String, ByVal strPassWord1 As String, ByVal rep_Path1 As String, ByVal attachment_path1 As String, ByVal Host_Url As String) As String
        ''Dim reportPath As String = String.Empty
        If Not String.IsNullOrEmpty(rname) Then
            reportPath = String.Empty


            Dim directoryString As String = attachment_path1 & "\" & _newSNo
            If Not Directory.Exists(directoryString) Then
                Directory.CreateDirectory(directoryString)
            End If
            ' SAVE THE CURRENT FOLDER PATH
            sp3 = Split(s2, "[_]")
            Dim currentFileName As String = (rname.Replace(".repx", "") & " ").Trim & "_" & sp3(0).ToString.Trim & ".pdf"
            reportPath = directoryString & "\" & currentFileName
            LogExceptions(strLogFilePath, Nothing, "Step 7 (Report Path): " & reportPath)

            Dim url As String
            url = Host_Url & "related_link_report_auto.aspx?pid=AUTO_DCUMENT_DISTRIBUTION&pmifields=PRIMARY_KEY1[_]&pmifieldvalues=" & s2 & "&mtype=REPX-REPORT&apifields=" & s1 & "&rptnm=" & rname & "&vwnm=PACKING_SLIP_REPORT_SEARCH&toemail=&sf=no&actiontypes=&rptprint=N&rptpreview=Y&RELATED_LINK_ID=PRNT_PS_WOC&rptcopies=1&SPath=" & reportPath & "&Type=PDF&PF=PL"
            LogExceptions(strLogFilePath, Nothing, "Step 7 (URL): " & url)
            Dim thread1 As New Thread(New ThreadStart(Sub()
                                                          Dim WebBrowser1 As New WebBrowser()
                                                          WebBrowser1.Navigate(url)
                                                          LogExceptions(strLogFilePath, Nothing, "Appp 1111")

                                                          While WebBrowser1.ReadyState <> WebBrowserReadyState.Complete
                                                              Application.DoEvents()
                                                          End While
                                                          WebBrowser1.Dispose()
                                                          LogExceptions(strLogFilePath, Nothing, "Appp 2222")
                                                      End Sub))
            thread1.SetApartmentState(ApartmentState.STA)
            thread1.Start()
            thread1.Join()
            Thread.Sleep(100000)
            thread1.Abort()
            'rpt1 = GenerateReport(str1, rname, wClause, s1, s2, strServerName1, strUserName1, strPassWord1, rep_Path1)
            ' B: CREATE A FOLDER UNDER "emailAttachments"

            Try
                'rpt1.CreateDocument()
                'rpt1.ExportToPdf(reportPath)
            Catch ex As Exception
                LogExceptions(strLogFilePath, Nothing, "Step 8 : Error")
                LogExceptions(strLogFilePath, ex, Nothing)
            End Try
            LogExceptions(strLogFilePath, Nothing, "Step 8 : Report Created and saved in path Success")
        End If
        Return reportPath
    End Function


    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub


End Class
