Imports System.IO
Imports System.Timers
Imports System.Threading
Imports System.Configuration
Imports System.Data
Imports System.Reflection
Imports Npgsql

Public Class event_scheduler_PG

    Public sp As StreamWriter

    Dim da As NpgsqlDataAdapter
    Dim dbcon As NpgsqlConnection

    Dim pdfCreatedPath As String = String.Empty

    Protected Shared AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
    Protected Shared strLogFilePath As String = AppPath + "\" + "EDGEevent_schedulerLog_pg.txt"
    Private Shared sw As StreamWriter = Nothing
    Dim scheduleTimer As New System.Timers.Timer()
    Public date_range As Date
    Dim contactsQuery As String
    Dim sp1, sp2, sp3
    Dim reportPath As String = String.Empty
    Dim count As Integer = 0
    Dim nSchemas As Integer = 0
    Dim pkValue As String = String.Empty

    Public hours1 As String = String.Empty
    Public mins1 As String = String.Empty

    Public dbad As New NpgsqlDataAdapter



    Protected Overrides Sub OnStart(ByVal args() As String)
        Try
            LogExceptions(strLogFilePath, Nothing, "Step 1")
            nSchemas = Convert.ToInt16(ConfigurationManager.AppSettings("NoOfSchemas").ToString())
            hours1 = ConfigurationManager.AppSettings("Hours").ToString()
            mins1 = ConfigurationManager.AppSettings("Mins").ToString()

            LogExceptions(strLogFilePath, Nothing, "Step 2")
            scheduleTimer.Interval = 1000
            AddHandler scheduleTimer.Elapsed, New ElapsedEventHandler(AddressOf TimerElapsed)
            scheduleTimer.Enabled = True
            LogExceptions(strLogFilePath, Nothing, "Service Started")
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        Try
            Thread.Sleep(1000)
            scheduleTimer.Enabled = False
            LogExceptions(strLogFilePath, Nothing, "Service Stopped")
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub

    Private Sub TimerElapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        Try
            scheduleTimer.Enabled = False

            hours1 = ConfigurationManager.AppSettings("Hours").ToString()
            mins1 = ConfigurationManager.AppSettings("Mins").ToString()

            Call Code_Execte_ForMultiSchemas()

            scheduleTimer.Enabled = True
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub

    Public Shared Sub LogExceptions(ByVal filePath As String, Optional ByVal ex As Exception = Nothing, Optional ByVal msg As String = Nothing)
        If File.Exists(filePath) Then
            File.Delete(filePath)
        End If
        If False = File.Exists(filePath) Then
            Dim fs As New FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
            fs.Close()
        End If
        WriteExceptionLog(filePath, ex, msg)
    End Sub

    Private Shared Sub WriteExceptionLog(ByVal strPathName As String, Optional ByVal objException As Exception = Nothing, Optional ByVal msg As String = Nothing)
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

    Public Sub Code_Execte_ForMultiSchemas()
        Try
            If nSchemas = 0 Then
                nSchemas = 10
            End If

            For i As Integer = 1 To nSchemas

                Dim strServerName As String = String.Empty
                Dim strUserName As String = String.Empty
                Dim strPassWord As String = String.Empty

                strServerName = ConfigurationManager.AppSettings("Server" & i)


                If Not String.IsNullOrEmpty(strServerName) Then
                    dbcon = New NpgsqlConnection(strServerName)
                    dbad.SelectCommand = New NpgsqlCommand
                    dbad.SelectCommand.Connection = dbcon
                    Dim ds As New Data.DataSet
                    If (dbad.SelectCommand.Connection.State = ConnectionState.Closed) Then
                        dbad.SelectCommand.Connection.Open()
                    End If
                    dbad.SelectCommand.CommandText = "SELECT MESSAGE_ID, DOC_TYPE, ACTION_PROGRAM FROM sys_event_scheduler WHERE STATUS_FLAG='N' and DOC_TYPE <>'' AND ACTION_PROGRAM <>'' ORDER BY CREATED_DATE"
                    ds.Clear()
                    Try
                        dbad.Fill(ds)
                    Catch ex As Exception
                    End Try
                    If (dbad.SelectCommand.Connection.State = ConnectionState.Open) Then
                        dbad.SelectCommand.Connection.Close()
                    End If
                    ''dbad.SelectCommand.Connection.Close()
                    If (ds.Tables(0).Rows.Count > 0) Then
                        For i1 = 0 To ds.Tables(0).Rows.Count - 1
                            If (ds.Tables(0).Rows(i1)("DOC_TYPE") = "FUNCTION") Then
                                dbad.UpdateCommand = New NpgsqlCommand
                                dbad.UpdateCommand.Connection = dbcon
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "update sys_event_scheduler set EXECUTION_START=now() WHERE MESSAGE_ID='" & ds.Tables(0).Rows(i1)("MESSAGE_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception

                                End Try
                                dbad.UpdateCommand.Connection.Close()
                                Try
                                    Call execute_dbfunction(ds.Tables(0).Rows(i1)("MESSAGE_ID"), ds.Tables(0).Rows(i1)("ACTION_PROGRAM"), strServerName, strUserName, strPassWord)
                                Catch ex As Exception
                                    Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
                                End Try
                            End If
                            If (ds.Tables(0).Rows(i1)("DOC_TYPE") = "PROCEDURE") Then
                                dbad.UpdateCommand = New NpgsqlCommand
                                dbad.UpdateCommand.Connection = dbcon
                                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                                    dbad.UpdateCommand.Connection.Open()
                                End If
                                dbad.UpdateCommand.CommandText = "update sys_event_scheduler set EXECUTION_START=now() WHERE MESSAGE_ID='" & ds.Tables(0).Rows(i1)("MESSAGE_ID") & "'"
                                Try
                                    dbad.UpdateCommand.ExecuteNonQuery()
                                Catch ex As Exception
                                    Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
                                End Try

                                dbad.UpdateCommand.Connection.Close()
                                Dim p1
                                p1 = Split(ds.Tables(0).Rows(i1)("ACTION_PROGRAM"), ",")
                                If (UBound(p1) > 2) Then
                                    Try
                                        Call execute_storeProcedure(ds.Tables(0).Rows(i1)("MESSAGE_ID"), p1(0), p1(1), p1(2), p1(3), strServerName)
                                    Catch ex As Exception
                                        Call Insert_Sys_Log(ex.Message.ToString(), strServerName, strUserName, strPassWord)
                                    End Try
                                End If
                            End If
                        Next
                    End If


                End If
            Next
        Catch ex As Exception
            dbad.SelectCommand.Connection.Close()
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub
    Public Sub execute_dbfunction(ByVal mid1 As String, ByVal fname As String, ByVal strServerName1 As String, ByVal strUserName1 As String, ByVal strPassWord1 As String)
        dbcon = New NpgsqlConnection(strServerName1)

        Try

            dbad.SelectCommand = New NpgsqlCommand
            dbad.SelectCommand.Connection = dbcon
            Dim ds As New Data.DataSet
            dbad.SelectCommand.CommandText = fname
            ds.Clear()
            If (dbad.SelectCommand.Connection.State = ConnectionState.Closed) Then
                dbad.SelectCommand.Connection.Open()
            End If
            dbad.Fill(ds)
            If (dbad.SelectCommand.Connection.State = ConnectionState.Open) Then
                dbad.SelectCommand.Connection.Close()
            End If
            ''dbad.SelectCommand.Connection.Close()

            dbad.UpdateCommand = New NpgsqlCommand
            dbad.UpdateCommand.Connection = dbcon
            If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                dbad.UpdateCommand.Connection.Open()
            End If
            dbad.UpdateCommand.CommandText = "update sys_event_scheduler set STATUS_FLAG='Y', STATUS_MESSAGE='SUCCESS',CHANGED_BY='AUTO', CHANGED_DATE=now(),EXECUTION_END=now() WHERE MESSAGE_ID='" & mid1 & "'"
            dbad.UpdateCommand.ExecuteNonQuery()
            If (dbad.UpdateCommand.Connection.State = ConnectionState.Open) Then
                dbad.UpdateCommand.Connection.Close()
            End If
            ''dbad.UpdateCommand.Connection.Close()
        Catch ex As Exception
            dbad.SelectCommand.Connection.Close()
            LogExceptions(strLogFilePath, ex, Nothing)
            dbad.UpdateCommand = New NpgsqlCommand
            dbad.UpdateCommand.Connection = dbcon
            If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                dbad.UpdateCommand.Connection.Open()
            End If
            dbad.UpdateCommand.CommandText = "update sys_event_scheduler set STATUS_FLAG='R', STATUS_MESSAGE='" & Replace(Mid(ex.Message, 1, 3990), "'", "''") & "',CHANGED_BY='AUTO', CHANGED_DATE=now(),EXECUTION_END=now() WHERE MESSAGE_ID='" & mid1 & "'"
            dbad.UpdateCommand.ExecuteNonQuery()
            dbad.UpdateCommand.Connection.Close()
        End Try


    End Sub
    Public Sub execute_storeProcedure(ByVal mid1 As String, ByVal fname As String, ByVal plist As String, ByVal plist1 As String, ByVal plist2 As String, ByVal strServerName1 As String)
        dbcon = New NpgsqlConnection(strServerName1)
        Dim s, s1 As String
        s = ""
        s1 = ""
        Dim pu1, pu2, pu3
        Dim rct, pp As Integer
        If (plist <> "") Then
            pu1 = Split(plist, "#")
            pu2 = Split(plist1, "#")
            pu3 = Split(plist2, "#")
            If (UBound(pu1) > 0) Then
                rct = UBound(pu1)
            Else
                rct = 0
            End If
        Else
            rct = -1
        End If
        Dim rvalue As String
        dbad.SelectCommand = New NpgsqlCommand
        dbad.SelectCommand.Connection = dbcon
        dbad.SelectCommand.Parameters.Clear()
        dbad.SelectCommand.CommandType = Data.CommandType.Text
        If (rct = -1) Then
            dbad.SelectCommand.CommandText = ("{call " & fname & "()}")
        Else
            If (rct = 0) Then
                dbad.SelectCommand.CommandText = ("{call " & fname & "(?)}")
            End If
        End If
        If (rct >= 0) Then
            If (rct = 0) Then
                If (UCase(plist2) = "N") Then
                    s1 = s1 & plist
                Else
                    s1 = s1 & "'" & plist & "'"
                End If
            Else
                For pp = 0 To rct
                    If (UCase(pu3(pp)) = "N") Then
                        s1 = s1 & pu1(pp) & ","
                    Else
                        s1 = s1 & "'" & pu1(pp) & "',"
                    End If
                Next
                s1 = Mid(s1, 1, Len(s1) - 1)
            End If
        End If
        If (s1 <> "") Then
            dbad.SelectCommand.CommandText = "call " & fname & "(" & s1 & ")"
        End If
        If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Closed) Then
            dbad.SelectCommand.Connection.Open()
        End If
        Try
            If (dbad.SelectCommand.Connection.State = ConnectionState.Closed) Then
                dbad.SelectCommand.Connection.Open()
            End If
            dbad.SelectCommand.ExecuteNonQuery()
            Try
                dbad.UpdateCommand = New NpgsqlCommand
                dbad.UpdateCommand.Connection = dbcon
                If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                    dbad.UpdateCommand.Connection.Open()
                End If
                dbad.UpdateCommand.CommandText = "update sys_event_scheduler set STATUS_FLAG='Y', STATUS_MESSAGE='SUCCESS',CHANGED_BY='AUTO', CHANGED_DATE=now(),EXECUTION_END=now() WHERE MESSAGE_ID='" & mid1 & "'"
                dbad.UpdateCommand.ExecuteNonQuery()
                If (dbad.UpdateCommand.Connection.State = ConnectionState.Open) Then
                    dbad.UpdateCommand.Connection.Close()
                End If
                ''dbad.UpdateCommand.Connection.Close()
            Catch ex As Exception
                LogExceptions(strLogFilePath, ex, Nothing)
            End Try

        Catch ex As Exception
            dbad.SelectCommand.Connection.Close()
            LogExceptions(strLogFilePath, ex, Nothing)
            dbad.UpdateCommand = New NpgsqlCommand
            dbad.UpdateCommand.Connection = dbcon
            If (dbad.UpdateCommand.Connection.State = ConnectionState.Closed) Then
                dbad.UpdateCommand.Connection.Open()
            End If
            dbad.UpdateCommand.CommandText = "update sys_event_scheduler set STATUS_FLAG='R', STATUS_MESSAGE='" & Replace(Mid(ex.Message, 1, 3990), "'", "''") & "',CHANGED_BY='AUTO', CHANGED_DATE=now(),EXECUTION_END=now() WHERE MESSAGE_ID='" & mid1 & "'"
            dbad.UpdateCommand.ExecuteNonQuery()
            If (dbad.UpdateCommand.Connection.State = ConnectionState.Open) Then
                dbad.UpdateCommand.Connection.Close()
            End If
        End Try
        If (dbad.SelectCommand.Connection.State = ConnectionState.Open) Then
            dbad.SelectCommand.Connection.Close()
        End If
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
            dbad.InsertCommand.CommandText = "Insert into SYS_ACTIVATE_STATUS_LOG (LINE_NO, CHANGE_REQUEST_NO,  OBJECT_TYPE, OBJECT_NAME, ERROR_TEXT, STATUS,LOG_DATE,ERROR_TEXT1, ERROR_TEXT2, ERROR_TEXT3) values ((select COALESCE(max(line_no),0)+1 from SYS_ACTIVATE_STATUS_LOG),'EDGE','SERVICE','MESSAGE_QUEUE_NEW','" & sterr1 & "','N',now(),'" & sterr2 & "','" & sterr3 & "','" & sterr4 & "')"
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

End Class
