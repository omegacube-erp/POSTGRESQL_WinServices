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
Imports System.IO
Imports System.Data
Imports System.Configuration
Imports System.Timers
Imports FE_SymmetricNamespace
Imports Npgsql
Public Class Alert_execution_Service_PG


    Public sp As StreamWriter


    Dim da As NpgsqlDataAdapter
    Dim dbad As New NpgsqlDataAdapter
    Dim con As NpgsqlConnection
    Dim cmd As NpgsqlCommand

    Dim dbcon As NpgsqlConnection

    Dim pdfCreatedPath As String = String.Empty

    Protected Shared AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
    Protected Shared strLogFilePath As String = AppPath + "\" + "AlertExecution_ExLog_pg.txt"
    Private Shared sw As StreamWriter = Nothing
    Dim scheduleTimer As New System.Timers.Timer()
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
                Call Code_Execte_ForMultiSchemas()
            Catch ex As Exception
                LogExceptions(strLogFilePath, ex, "TimerElapsed - Code_Execte_ForMultiSchemas")
            End Try


            scheduleTimer.Enabled = True
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub
    'For Message Queue Service
    Public Sub Code_Execte_ForMultiSchemas()
        Dim strServerName As String = String.Empty
        Dim strUserName As String = String.Empty
        Dim strPassWord As String = String.Empty
        Try
            If nSchemas = 0 Then
                nSchemas = 10
            End If

            For i As Integer = 1 To nSchemas

                strServerName = String.Empty
                strUserName = String.Empty
                strPassWord = String.Empty

                strServerName = ConfigurationManager.AppSettings("Server" & i)
                strUserName = "X" 'ConfigurationManager.AppSettings("User" & i)
                strPassWord = "Y" 'ConfigurationManager.AppSettings("Pwd" & i)
                Try
                    If Not String.IsNullOrEmpty(strServerName) Then
                        Try
                            If Not String.IsNullOrEmpty(strUserName) Then
                                Try
                                    If Not String.IsNullOrEmpty(strPassWord) Then
                                        Try
                                            Call code_execute(strServerName, strUserName, strPassWord)
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
    Public Sub code_execute(ByVal strServerName1 As String, ByVal strUserName1 As String, ByVal strPassWord1 As String)
        Dim msg1 As String = String.Empty
        Dim msg2 As String = String.Empty
        Dim msg3 As String = String.Empty
        Dim msg4 As String = String.Empty
        Dim msg5 As String = String.Empty

        Try
            Dim dbad As New NpgsqlDataAdapter
            dbcon = New NpgsqlConnection(strServerName1)

            dbad.SelectCommand = New NpgsqlCommand
            dbad.InsertCommand = New NpgsqlCommand
            dbad.UpdateCommand = New NpgsqlCommand
            dbad.DeleteCommand = New NpgsqlCommand

            dbad.SelectCommand.Connection = dbcon
            dbad.InsertCommand.Connection = dbcon
            dbad.UpdateCommand.Connection = dbcon
            dbad.DeleteCommand.Connection = dbcon
            Try
                If (dbad.InsertCommand.Connection.State = ConnectionState.Closed) Then
                    dbad.InsertCommand.Connection.Open()
                End If
                ''dbad.InsertCommand.CommandText = "INSERT INTO SYS_OBJECT_REPOSITORY (ID, OBJECT_NAME, CREATED_BY, CREATED_DATE, MAINTENANCE_LEVEL) SELECT (SELECT COALESCE(MAX(ID),0) FROM SYS_OBJECT_REPOSITORY )+ROWNUM,NAME,'AUTO',now(),'2' FROM V_SYS_OBJECT_LIST WHERE NAME NOT IN (SELECT OBJECT_NAME FROM SYS_OBJECT_REPOSITORY)"
                dbad.InsertCommand.CommandText = "INSERT INTO SYS_OBJECT_REPOSITORY (ID, OBJECT_NAME, CREATED_BY, CREATED_DATE, MAINTENANCE_LEVEL,OBJECT_TYPE) SELECT (SELECT COALESCE(MAX(ID),0) FROM SYS_OBJECT_REPOSITORY )+ROW_NUMBER() OVER(),NAME,'AUTO',now(),'2',OBJECT_TYPE FROM V_SYS_OBJECT_LIST WHERE NAME NOT IN (SELECT OBJECT_NAME FROM SYS_OBJECT_REPOSITORY)"
                dbad.InsertCommand.ExecuteNonQuery()

            Catch ex As Exception
                Call Insert_Sys_Log(ex.Message.ToString(), strServerName1, strUserName1, strPassWord1)
                LogExceptions(strLogFilePath, ex, Nothing)
            End Try



            Try
                Dim dbautogen, dbss, dbuser As New DataSet
                Dim dlabel, dtext, utext, user_id As String
                Dim i4, i1, i2, i3, i5 As Integer
                i4 = 0

                'dbad.SelectCommand.CommandText = "SELECT USER_ID FROM SYS_USER_MASTER WHERE ACTIVE_FLAG='Y' ORDER BY USER_ID"
                msg1 = String.Empty
                'msg1 = dbad.SelectCommand.CommandText
                'dbuser.Clear()
                'dbad.Fill(dbuser)
                'If (dbuser.Tables(0).Rows.Count > 0) Then
                '    For i5 = 0 To dbuser.Tables(0).Rows.Count - 1
                '        user_id = dbuser.Tables(0).Rows(i5)("USER_ID")
                ' LogExceptions(strLogFilePath, Nothing, user_id & " Start :-" + DateTime.Now)
                dbad.SelectCommand.CommandText = "select DOCUMENT_ID, DOCUMENT_LABEL, DISPLAY_MESSAGE, OPEN_URL,STATUS_VIEW_TYPE, EXTERNAL_URL,ALERT_FREQUENCY_UNIT, ALERT_FREQUENCY from SYS_STATUS_VIEWS_HEADER where COALESCE(ACTIVE_FLAG,'N')='Y' AND UPPER(ALERT_FREQUENCY_UNIT)!='INSTANT' AND COALESCE(LAST_EXECUTION_DATE,now())<=now()" ' AND  document_id in ((select document_id from SYS_STATUS_VIEWS_USER_GROUPS a,SYS_USER_ACCESS b where a.user_group= b.user_group_id and UPPER(b.user_id)=upper('" & user_id & "')) union all (Select document_id from SYS_STATUS_VIEWS_USERS a, sys_user_master b where a.user_id=b.user_id and UPPER(a.user_id)=upper('" & user_id & "')))"
                msg2 = String.Empty
                msg2 = dbad.SelectCommand.CommandText.ToString
                dbautogen.Clear()
                dbad.Fill(dbautogen)
                If (dbautogen.Tables(0).Rows.Count > 0) Then

                    Dim st1, st2 As String


                    For i1 = 0 To dbautogen.Tables(0).Rows.Count - 1
                        If (dbad.DeleteCommand.Connection.State = System.Data.ConnectionState.Closed) Then
                            dbad.DeleteCommand.Connection.Open()
                        End If
                        dbad.DeleteCommand.CommandText = "Delete from SYS_STATUS_ALERTS where DOCUMENT_ID='" & dbautogen.Tables(0).Rows(i1)("DOCUMENT_ID") & "'"
                        dbad.DeleteCommand.ExecuteNonQuery()
                        ' LogExceptions(strLogFilePath, Nothing, dbautogen.Tables(0).Rows(i1)("DOCUMENT_ID") & " start :-" + DateTime.Now)
                        pkValue = String.Empty
                        pkValue = "Code_Execute - Data Source=" + strServerName1 + ";Schema Name=" + strUserName1 + " DOCUMENT_ID='" & dbautogen.Tables(0).Rows(i1)("DOCUMENT_ID") & "'"
                        Try
                            st1 = ""
                            st2 = ""
                            If Not (Equals(dbautogen.Tables(0).Rows(i1)("STATUS_VIEW_TYPE"), System.DBNull.Value)) Then
                                st1 = dbautogen.Tables(0).Rows(i1)("STATUS_VIEW_TYPE")
                            End If
                            If Not (Equals(dbautogen.Tables(0).Rows(i1)("EXTERNAL_URL"), System.DBNull.Value)) Then
                                st2 = dbautogen.Tables(0).Rows(i1)("EXTERNAL_URL")
                            End If

                            Try  '' Nelwly Added - START  Only Try Catch Block
                                Dim dbautogen1 As New System.Data.DataSet
                                Dim ac As Integer = 0
                                dbad.SelectCommand.CommandText = "SELECT SQL_QUERY FROM SYS_STATUS_VIEWS_HEADER WHERE DOCUMENT_ID='" & dbautogen.Tables(0).Rows(i1)("DOCUMENT_ID") & "'"
                                dbautogen1.Clear()
                                dbad.Fill(dbautogen1)
                                If (dbautogen1.Tables(0).Rows.Count > 0) Then
                                    dbad.SelectCommand.CommandText = dbautogen1.Tables(0).Rows(0)("SQL_QUERY")
                                    msg5 = String.Empty
                                    msg5 = dbad.SelectCommand.CommandText.ToString
                                    dbss.Clear()
                                    dbad.Fill(dbss)

                                    If (dbss.Tables(0).Rows.Count > 0) Then
                                        ac = dbss.Tables(0).Rows.Count
                                    End If
                                End If
                                Try
                                    If Not (Equals(dbautogen.Tables(0).Rows(i1)("ALERT_FREQUENCY_UNIT"), System.DBNull.Value)) Then
                                        If Not (Equals(dbautogen.Tables(0).Rows(i1)("ALERT_FREQUENCY"), System.DBNull.Value)) Then
                                            If (dbad.UpdateCommand.Connection.State = System.Data.ConnectionState.Closed) Then
                                                dbad.UpdateCommand.Connection.Open()
                                            End If
                                            If (dbautogen.Tables(0).Rows(i1)("ALERT_FREQUENCY_UNIT") = "MINUTES") Then
                                                dbad.UpdateCommand.CommandText = "UPDATE SYS_STATUS_VIEWS_HEADER SET alert_count=" & ac & ",LAST_EXECUTION_DATE=now()+ interval '1' minute * " & dbautogen.Tables(0).Rows(i1)("ALERT_FREQUENCY") & " WHERE DOCUMENT_ID='" & dbautogen.Tables(0).Rows(i1)("DOCUMENT_ID") & "'"
                                                dbad.UpdateCommand.ExecuteNonQuery()
                                            End If
                                            If (dbautogen.Tables(0).Rows(i1)("ALERT_FREQUENCY_UNIT") = "HOURS") Then
                                                dbad.UpdateCommand.CommandText = "UPDATE SYS_STATUS_VIEWS_HEADER SET alert_count=" & ac & ",LAST_EXECUTION_DATE=now()+ interval '1' hour *" & dbautogen.Tables(0).Rows(i1)("ALERT_FREQUENCY") & " WHERE DOCUMENT_ID='" & dbautogen.Tables(0).Rows(i1)("DOCUMENT_ID") & "'"
                                                dbad.UpdateCommand.ExecuteNonQuery()
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    Call Insert_Sys_Log(ex.Message.ToString() & " " & pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5, strServerName1, strUserName1, strPassWord1)
                                    LogExceptions(strLogFilePath, ex, pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5)
                                End Try

                            Catch ex As Exception
                                'eventLogEQueryEx.Clear()
                                'eventLogEQueryEx.WriteEntry(ex.Message.ToString)
                                'eventLogEQueryEx.WriteEntry(pkValue)
                                Call Insert_Sys_Log(ex.Message.ToString() & " " & pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5, strServerName1, strUserName1, strPassWord1)
                                LogExceptions(strLogFilePath, ex, pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5)
                            End Try  '' Nelwly Added - END Only Try Catch Block
                        Catch ex As Exception
                            'eventLogEQueryEx.Clear()
                            'eventLogEQueryEx.WriteEntry(ex.Message.ToString)
                            'eventLogEQueryEx.WriteEntry(pkValue)
                            Call Insert_Sys_Log(ex.Message.ToString() & " " & pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5, strServerName1, strUserName1, strPassWord1)
                            LogExceptions(strLogFilePath, ex, pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5)
                        End Try

                    Next

                End If

            Catch ex As Exception
                'eventLogEQueryEx.Clear()
                'eventLogEQueryEx.WriteEntry(ex.Message.ToString)
                'eventLogEQueryEx.WriteEntry(pkValue)
                Call Insert_Sys_Log(ex.Message.ToString() & " " & pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5, strServerName1, strUserName1, strPassWord1)
                LogExceptions(strLogFilePath, ex, pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5)
            End Try

            If (dbad.UpdateCommand.Connection.State = ConnectionState.Open) Then
                dbad.UpdateCommand.Connection.Close()
            End If
            If (dbad.SelectCommand.Connection.State = ConnectionState.Open) Then
                dbad.SelectCommand.Connection.Close()
            End If
            If (dbad.InsertCommand.Connection.State = ConnectionState.Open) Then
                dbad.InsertCommand.Connection.Close()
            End If
            If (dbad.DeleteCommand.Connection.State = ConnectionState.Open) Then
                dbad.DeleteCommand.Connection.Close()
            End If
            dbad.Dispose()
        Catch ex As Exception
            'eventLogEQueryEx.Clear()
            'eventLogEQueryEx.WriteEntry(ex.Message.ToString)
            'eventLogEQueryEx.WriteEntry(pkValue)
            Call Insert_Sys_Log(ex.Message.ToString() & " " & pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5, strServerName1, strUserName1, strPassWord1)
            LogExceptions(strLogFilePath, ex, pkValue & ";##" & msg1 & ";##" & msg2 & ";##" & msg3 & ";##" & msg4 & ";##" & msg5)
        End Try
    End Sub
    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
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
End Class
