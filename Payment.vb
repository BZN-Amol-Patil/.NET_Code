Imports System.IO

Module Payment

    Dim objLogCls As New ClsErrLog
    Dim objGetSetINI As New ClsShared
    Dim objBaseClass As ClsBase
    Public Payment_T_Record As String

    Public Function GenerateNestleOutputFile(ByRef dtNESTLE As DataTable, ByVal strFileName As String) As Boolean


        Dim objstrWriter As StreamWriter
        Dim strOutputStream As String
        Dim FileCounter As String
        Dim intRowCount As Integer

        Try

            GenerateNestleOutputFile = False

            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")

            intRowCount = 0
            FileCounter = objGetSetINI.GetINISettings("General", "File Counter", My.Application.Info.DirectoryPath & "\settings.ini")
            FileCounter += 1
            If Len(FileCounter) < 3 Then
                FileCounter = FileCounter.PadLeft(4, "0").Trim()
                FileCounter = FileCounter.Substring(FileCounter.Length - 3, 3)
            End If
            If nmConventionSFTP.ToString.Trim().ToUpper() = "Y".ToString.Trim().ToUpper() Then
                gstrOutputFile = strDomainID & "_" & strClientCode & "_" & strClientCode & Format(Now.Date(), "ddMM") & "." & FileCounter
            Else
                gstrOutputFile = strClientCode & Format(Now.Date(), "ddMM") & "." & FileCounter
            End If

            If Not objstrWriter Is Nothing Then
                objstrWriter.Close()
                objstrWriter.Dispose()
            End If

            objstrWriter = New StreamWriter(strTempFolderPath & "\" & gstrOutputFile)

            objLogCls.LogEntry("[Output File Name] = " & gstrOutputFile)
            For Each dtRow As DataRow In dtNESTLE.Select("[TXN_NO]<>0", "[TXN_NO]")
                ''OUTPUT TRANSACTION
                strOutputStream = ""
                strOutputStream = Check_Comma(Pad_Length(dtRow("Customer Reference No").ToString(), 15))
                strOutputStream = strOutputStream & Check_Comma(Pad_Length(dtRow("Beneficiary Name").ToString(), 70))
                strOutputStream = strOutputStream & Check_Comma(Pad_Length(dtRow("Beneficiary Account Number").ToString(), 30))
                strOutputStream = strOutputStream & Check_Comma(Pad_Length(dtRow("IFSC Code").ToString(), 11))
                strOutputStream = strOutputStream & Check_Comma(Pad_Length(dtRow("Account Type").ToString(), 2))
                strOutputStream = strOutputStream & Check_Comma(Pad_Length(dtRow("Instrument Amount").ToString(), 15))
                strOutputStream = strOutputStream & Check_Comma(Pad_Length(dtRow("Value Date").ToString(), 8))

                ''For Last Value Contain (,)
                strOutputStream = strOutputStream.Substring(0, strOutputStream.Length - 1)
                objstrWriter.WriteLine(strOutputStream, gstrOutputFile)
                intRowCount += 1
            Next

            Call objGetSetINI.SetINISettings("General", "File Counter", Val(FileCounter), My.Application.Info.DirectoryPath & "\settings.ini")

            If Not objstrWriter Is Nothing Then
                objstrWriter.Close()
                objstrWriter.Dispose()
            End If

        Catch ex As Exception
            GenerateNestleOutputFile = True
            Call objLogCls.Handle_Error(ex, "Payment", Err.Number, "GenerateNestleOutputFile")

        Finally

            If Not objstrWriter Is Nothing Then
                objstrWriter.Close()
                objstrWriter.Dispose()
            End If

            If Not objBaseClass Is Nothing Then
                objBaseClass.Dispose()
                objBaseClass = Nothing
            End If

        End Try

    End Function

    Private Sub ClearArray(ByRef pArr() As String)
        'Added by Jaiwant dtd 03-06-2011
        Try
            For I As Int16 = 0 To pArr.Length - 1
                pArr(I) = ""
            Next

        Catch ex As Exception

        End Try

    End Sub

    Private Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then
                Check_Comma = Chr(34) & strTemp & Chr(34) & ","
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            objLogCls.WriteErrorToTxtFile(Err.Number, Err.Description, "Payment", "Check_Comma")

        End Try
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen).Trim()

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objLogCls.Handle_Error(ex, "Payment", Err.Number, "Pad_Length")

        End Try
    End Function

End Module
