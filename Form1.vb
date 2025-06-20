Imports System
Imports System.IO


Public Class Form1
    Dim objBaseClass As ClsBase
    Dim objFileValidate As ClsValidation
    Dim objGetSetINI As ClsShared

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            Timer1.Enabled = True
            Timer1.Interval = 1000
            blnErrorLog = False

            Generate_SettingFile()

        Catch ex As Exception
            Call objBaseClass.Handle_Error(ex, "Form", Err.Number, "Form_Load")

        End Try

    End Sub
    Private Sub Generate_SettingFile()
        Dim strConverterCaption As String = ""
        Dim strSettingsFilePath As String = My.Application.Info.DirectoryPath & "\settings.ini"

        Try
            objGetSetINI = New ClsShared

            '-Genereate Settings.ini File-
            If Not File.Exists(strSettingsFilePath) Then
                '-General Section-
                Call objGetSetINI.SetINISettings("General", "Date", Format(Now, "dd/MM/yyyy"), strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Audit Log", My.Application.Info.DirectoryPath & "\Audit", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Error Log", My.Application.Info.DirectoryPath & "\Error", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Input Folder", My.Application.Info.DirectoryPath & "\INPUT", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderSuc", My.Application.Info.DirectoryPath & "\Archive\Success", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderUnSuc", My.Application.Info.DirectoryPath & "\Archive\Unsuccess", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Temp Folder", My.Application.Info.DirectoryPath & "\Temp", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Output Folder", My.Application.Info.DirectoryPath & "\Output", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Report Folder", My.Application.Info.DirectoryPath & "\Report", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Validation", My.Application.Info.DirectoryPath & "\Validation\Validation.xls", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Special Character Validation", My.Application.Info.DirectoryPath & "\Validation\Special Character Mapping.xls", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Converter Caption", "Automotive Axle Converter", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Process Output File Ignoring Invalid Transactions", "N", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "File Counter", "0", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "==", "==", strSettingsFilePath) 'Separator

                Call objGetSetINI.SetINISettings("Client Details", "Naming Convention With SFTP (Y/N)", "N", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Client Name", "Automotive Axle", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Client Code", "SALARYNEFT", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Domain ID", "AUTOAXLE", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Input Date Format", "DD/MM/YYYY", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "==", "==", strSettingsFilePath) 'Separator

                Call objGetSetINI.SetINISettings("Input Details", "Account type", "11", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Input Details", "==", "==", strSettingsFilePath)




                '-Encryption Section-
                Call objGetSetINI.SetINISettings("Encryption", "Encryption Required (Y/N)", "N", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Encryption", "Batch File Path", "C:\GenericEncryption_Client\encryptdaemon.bat", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Encryption", "PICKDIR Path", "C:\GenericEncryption_Client\datafiles\clearfiles", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Encryption", "DROPDIR Path", "C:\GenericEncryption_Client\datafiles\encfiles", strSettingsFilePath)

            End If

            '-Get Converter Caption from Settings-
            If File.Exists(strSettingsFilePath) Then
                strConverterCaption = objGetSetINI.GetINISettings("General", "Converter Caption", strSettingsFilePath)
                If strConverterCaption <> "" Then
                    Text = strConverterCaption.ToString() & " - Version " & Mid(Application.ProductVersion.ToString(), 1, 3)
                Else
                    MsgBox("Either settings.ini file does not contains the key as [ Converter Caption ] or the key value is blank" & vbCrLf & "Please refer to " & strSettingsFilePath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End
                End If
            End If

        Catch ex As Exception
            MsgBox("Error - " & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error while Generating Settings File")
            End

        Finally
            If Not objGetSetINI Is Nothing Then
                objGetSetINI.Dispose()
                objGetSetINI = Nothing
            End If

        End Try

    End Sub


    Private Sub Conversion_Process()
        Dim objFolderAll As DirectoryInfo


        Try
            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")

            If objBaseClass Is Nothing Then
                objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            End If

            '-Get Settings- 
            If GetAllSettings() = True Then
                MsgBox("Either file path is invalid or any key value is left blank in settings.ini file", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in Settings.ini file")
                Exit Sub
            End If
            '---Input file
            objFolderAll = New DirectoryInfo(strInputFolderPath)
            If objFolderAll.GetFiles.Length = 0 Then
                objFolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started for INPUT Files")
                For Each objFileOne As FileInfo In objFolderAll.GetFiles("*.*")
                    objBaseClass.isCompleteFileAvailable(objFileOne.FullName)
                    If Mid(objFileOne.FullName, objFileOne.FullName.Length - 3, 4).ToString().ToUpper() <> ".BAK" Then
                        'objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("INPUT File [ " & objFileOne.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)

                        Process_Each(objFileOne.FullName)

                        objFolderAll.Refresh()
                    End If
                Next

            End If

        Catch ex As Exception
            'MsgBox("Error - " & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Conversion_Process")
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Conversion_Process")

        Finally
            If Not objBaseClass Is Nothing Then
                objBaseClass.Dispose()
                objBaseClass = Nothing
            End If

        End Try

    End Sub

    Private Function GetAllSettings() As Boolean

        Try
            GetAllSettings = False

            If Not File.Exists(My.Application.Info.DirectoryPath & "\settings.ini") Then
                GetAllSettings = True
                MsgBox("Either settings.ini file does not exists or invalid file path" & vbCrLf & My.Application.Info.DirectoryPath & "\settings.ini", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            End If

            '-Audit Folder Path-
            If strAuditFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strAuditFolderPath) Then
                    Directory.CreateDirectory(strAuditFolderPath)
                    If Not Directory.Exists(strAuditFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Audit Log folder. Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", True)
                        End If
                        MsgBox("Invalid path for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                        Exit Function
                    End If
                End If
            End If

            '-Error Folder Path-
            If strErrorFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Error Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strErrorFolderPath) Then
                    Directory.CreateDirectory(strErrorFolderPath)
                    If Not Directory.Exists(strErrorFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Error Log folder. Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Error Log folder." & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Input Folder Path-
            If strInputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Input folder" & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strInputFolderPath) Then
                    Directory.CreateDirectory(strInputFolderPath)
                    If Not Directory.Exists(strInputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Input Folder. Please check [ settings.ini ] file, the key as [ Input Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Input Folder." & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            ''-Advice Folder Path-
            'If strAdviceFolderPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Advice folder" & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not Directory.Exists(strAdviceFolderPath) Then
            '        Directory.CreateDirectory(strAdviceFolderPath)
            '        If Not Directory.Exists(strAdviceFolderPath) Then
            '            GetAllSettings = True
            '            If Not objBaseClass Is Nothing Then
            '                objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Advice Folder. Please check [ settings.ini ] file, the key as [ Advice Folder ] contains invalid path specification.", True)
            '            End If
            '            MsgBox("Invalid path for Advice Folder." & vbCrLf & "Please check settings.ini file, the key as [ Advice Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '        End If
            '    End If
            'End If


            '-Temp Folder Path-
            If strTempFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Temp folder" & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strTempFolderPath) Then
                    Directory.CreateDirectory(strTempFolderPath)
                    If Not Directory.Exists(strTempFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Temp Folder. Please check [ settings.ini ] file, the key as [ Temp Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Temp Folder." & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If


            '-Output Folder Path-
            If strOutputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Output folder" & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strOutputFolderPath) Then
                    Directory.CreateDirectory(strOutputFolderPath)
                    If Not Directory.Exists(strOutputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Output Folder. Please check [ settings.ini ] file, the key as [ Output Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Output Folder." & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If



            '-Report Folder Path-
            If strReportFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Report Folder." & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strReportFolderPath) Then
                    Directory.CreateDirectory(strReportFolderPath)
                    If Not Directory.Exists(strReportFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Report Folder. Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Report Folder." & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Archive Successful-
            If strArchivedFolderSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archive Suc folder" & vbCrLf & "Please check settings.ini file, the key as [ Archive Suc Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderSuc) Then
                    Directory.CreateDirectory(strArchivedFolderSuc)
                    If Not Directory.Exists(strArchivedFolderSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archive Suc Folder. Please check settings.ini file, the key as [ Archive Suc Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archive Suc folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
                    End If
                End If
            End If

            '-Archive Unsuccessful-
            If strArchivedFolderUnSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archive UnSuc folder" & vbCrLf & "Please check settings.ini file, the key as [ Archive UnSuc Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderUnSuc) Then
                    Directory.CreateDirectory(strArchivedFolderUnSuc)
                    If Not Directory.Exists(strArchivedFolderUnSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archive UnSuc Folder. Please check settings.ini file, the key as [ Archive UnSuc Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archive UnSuc folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
                    End If
                End If
            End If

            '-Validation Folder Path-
            If strValidationPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Validation file." & vbCrLf & "Please check settings.ini file, the key as [ Validation ] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not File.Exists(strValidationPath) Then
                    GetAllSettings = True
                    If Not objBaseClass Is Nothing Then
                        objBaseClass.LogEntry("Error in settings.ini file, Validation file does not exist or invalid file path", True)
                    End If
                    MsgBox("Validation file does not exist or invalid file path" & vbCrLf & strValidationPath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                End If
            End If

        Catch ex As Exception
            GetAllSettings = True
            'MsgBox("Error - " & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error While Getting Log Path from Settings.ini File")
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "GetAllSettings")

        End Try

    End Function

    Private Sub Process_Each(ByVal StrInputFileName As String)
        Dim strAns As String
        Try
            '-Verify Input File Name-
            If StrInputFileName.ToString().Trim() = "" Then
                objBaseClass.LogEntry("Input file does not exist or Invalid file path", True)
                Exit Sub
            End If



            gstrInputFolder = StrInputFileName.ToString().Substring(0, StrInputFileName.ToString().LastIndexOf("\"))
            gstrInputFile = Path.GetFileName(StrInputFileName.ToString())



            '-Verify Client Code-
            If strClientCode.ToString().Trim().Length = 0 Then
                objBaseClass.LogEntry("Client Code cannot be Blank", True)
                MsgBox("Client Code cannot be blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
                Exit Sub
            End If


            'Conversion Process

            objBaseClass.LogEntry("", False)
            objBaseClass.LogEntry("Process Started")
            objBaseClass.LogEntry("Reading Input file [ " & gstrInputFile & " ]", False)


            objFileValidate = New ClsValidation(StrInputFileName, objBaseClass.gstrIniPath, StrInputFileName)

            If objFileValidate.CheckValidateFile() = True Then
                objBaseClass.LogEntry("Input File Reading Completed Successfully")

                If objFileValidate.DtUnSucInput.Rows.Count = 0 Or strProceed.ToUpper = "Y" Then
                    objBaseClass.LogEntry("Input File Validated Successfully")


                    If objFileValidate.DtInput.Rows.Count > 0 Then
                        objBaseClass.LogEntry("Output File Generation Process Started")

                        If GenerateNestleOutputFile(objFileValidate.DtInput, gstrOutputFile) = True Then
                            objBaseClass.LogEntry("Output File Generation process failed due to Error", True)
                            objBaseClass.LogEntry("Output File Generation process failed due to Error", False)
                        Else
                            If strEncrypt.ToUpper() = "Y" Then
                                objBaseClass.LogEntry("Performing Output File Encryption", False)
                                objBaseClass.FileMove(strTempFolderPath & "\" & gstrOutputFile, strPICKDIRpath & "\" & gstrOutputFile)
                                objBaseClass.Execute_Batch_file()
                                objBaseClass.FileMove(strDROPDIRPath & "\" & gstrOutputFile, strOutputFolderPath & "\" & gstrOutputFile)
                                objBaseClass.LogEntry("Encryption performed Successfully", False)
                            Else
                                objBaseClass.FileMove(strTempFolderPath & "\" & gstrOutputFile, strOutputFolderPath & "\" & gstrOutputFile)
                            End If

                            '-Write Summary Report-
                            Dim strSummaryFileName As String
                            strSummaryFileName = Path.GetFileNameWithoutExtension(gstrInputFile)
                            objBaseClass.LogEntry("Writing Summary Report File [ Summary_" & strSummaryFileName & ".txt ]")
                            Summary_Report()
                            objBaseClass.LogEntry("Summary Report File Generated Successfully")

                            'System.Windows.Forms.Application.DoEvents()

                            '-Write Transaction Report-
                            objBaseClass.LogEntry("Writing Transaction Report File [ Transaction_Report_" & strSummaryFileName & ".csv ]")

                            'System.Windows.Forms.Application.DoEvents()

                            Payment_Report()

                            objBaseClass.LogEntry("Transaction Report File Generated Successfully")

                            'System.Windows.Forms.Application.DoEvents()


                            '-Output Success-
                            objBaseClass.LogEntry("Output File [ " & gstrOutputFile & " ] is Generated Successfully", False)
                            'MessageBox.Show("Output File [ " & gstrOutputFile & "] is Generated Successfully!", strClientName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                            objBaseClass.FileMove(StrInputFileName, strArchivedFolderSuc & "\" & Path.GetFileName(StrInputFileName))
                            objBaseClass.LogEntry("[ " & gstrInputFile & " ] files moved to Archived Folder Successful")
                        End If


                    Else
                        objBaseClass.LogEntry("No Valid Record present in Input File")
                        'MessageBox.Show("No Valid Record present in Input File", strClientName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))
                        objBaseClass.FileMove(gstrAdviceFolder & "\" & gstrAdviceFile, strArchivedFolderUnSuc & "\" & gstrAdviceFile)

                        objBaseClass.LogEntry("[ " & gstrInputFile & " ] files moved to Archived Folder UnSuccessful")

                    End If

                Else
                    objBaseClass.LogEntry("No Valid Record present in Input File")
                    'MessageBox.Show("No Valid Record present in Input File", strClientName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))
                    objBaseClass.FileMove(gstrAdviceFolder & "\" & gstrAdviceFile, strArchivedFolderUnSuc & "\" & gstrAdviceFile)
                    objBaseClass.LogEntry("[ " & gstrInputFile & " ] files moved to Archived Folder UnSuccessful")

                End If


                If objFileValidate.DtUnSucInput.Rows.Count > 0 Then
                    If objFileValidate.DtUnSucInput.Select("[Reason] <>''").Length > 0 Then
                        objBaseClass.LogEntry("Input File Contains Following Discrepancies")
                        objBaseClass.LogEntry("Writing Transaction instruction failed for Following Beneficiary Name")
                        With objFileValidate.DtUnSucInput
                            For Each _dtRow As DataRow In .Select("[TXN_NO]<>0")
                                If _dtRow("Beneficiary Name").ToString().Trim() <> "" Then
                                    objBaseClass.LogEntry("Beneficiary Name : " & _dtRow("Beneficiary Name").ToString & StrDup(5, " "))
                                End If
                                objBaseClass.LogEntry(_dtRow("REASON").ToString)
                            Next
                        End With
                    End If
                End If


                If strAns = 7 Then
                    objBaseClass.LogEntry("Output File Generation Failed")
                    objBaseClass.LogEntry("Proccess Terminated")
                End If

            Else
                objBaseClass.LogEntry("[" & gstrInputFile & " ] is not valid Input File", False)
                objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))
                objBaseClass.LogEntry("[ " & gstrInputFile & " ] files moved to Archived Folder UnSuccessful")
                'MessageBox.Show("[" & gstrInputFile & " ] is not Valid Input File", strClientName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            objBaseClass.LogEntry("Process Completed")

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Process_Each")

        Finally

            '-Error Log Link-

            If Not objFileValidate.DtInput Is Nothing Then
                objBaseClass.ObjectDispose(objFileValidate.DtInput)
            End If

            If Not objFileValidate.DtUnSucInput Is Nothing Then
                objBaseClass.ObjectDispose(objFileValidate.DtUnSucInput)
            End If

            If Not objFileValidate Is Nothing Then
                objFileValidate.Dispose()
                objFileValidate = Nothing
            End If

        End Try

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Interval = 1000
        Timer1.Enabled = False

        Conversion_Process()

        Timer1.Enabled = True
    End Sub
    Private Sub Summary_Report()
        Dim strSumFileName As String
        Try
            strSumFileName = "Summary_" & Path.GetFileNameWithoutExtension(gstrInputFile) & ".txt"

            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Summary Report As On [" & Format(Now, "dd-MM-yyyy hh:mm:ss") & "]")
            objBaseClass.WriteSummaryTxt(strSumFileName, StrDup(105, "-"))

            '-Summary of Input File-
            objBaseClass.WriteSummaryTxt(strSumFileName, "Input File Details")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Input File Name :  " & gstrInputFile)


            '-Summary of Output File-
            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Output File Details")
            objBaseClass.WriteSummaryTxt(strSumFileName, gstrOutputFile)

            'For iCount As Int32 = 0 To UBound(gstrOutputFileListing) - 1
            '    objBaseClass.WriteSummaryTxt(strSumFileName, ("Output File Name " & iCount + 1).ToString.PadRight(25, " ") & ":  " & gstrOutputFileListing(iCount))
            'Next

            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            If objFileValidate.DtInput.Rows.Count > 0 Then
                objBaseClass.WriteSummaryTxt(strSumFileName, ("Total No of Successfull Records").PadRight(40, " ") & ":" & objFileValidate.DtInput.Select("[TXN_NO]<> '0'").Length().ToString().PadLeft(15, " ") & StrDup(8, " ") & ("Total Amount").PadRight(25, " ") & ":" & GetPaymentAmount(True, "").ToString().PadLeft(20, " "))
                objBaseClass.WriteSummaryTxt(strSumFileName, StrDup(105, "-"))
            End If

            If objFileValidate.DtUnSucInput.Rows.Count > 0 Then
                objBaseClass.WriteSummaryTxt(strSumFileName, ("Total No of UnSuccessfull Records").PadRight(40, " ") & ":" & objFileValidate.DtUnSucInput.Select("[TXN_NO]<> '0'").Length().ToString().PadLeft(15, " ") & StrDup(8, " ") & ("Total Amount").PadRight(25, " ") & ":" & GetPaymentAmount(False, "").ToString().PadLeft(20, " "))
                objBaseClass.WriteSummaryTxt(strSumFileName, StrDup(105, "-"))
            End If

            objBaseClass.WriteSummaryTxt(strSumFileName, StrDup(105, "-"))

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Summary_Report")

        End Try

    End Sub

    Private Sub Payment_Report()
        Dim strSumFileName As String
        Dim strTranRepName As String

        Try
            strSumFileName = "Transaction_Report_" & Path.GetFileNameWithoutExtension(gstrInputFile) & ".csv"

            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "[" & Format(Now, "dd-MM-yyyy hh:mm:ss") & "]")

            objBaseClass.WriteSummaryTxt(strSumFileName, "Transaction Report for Input File " & gstrInputFile)
            objBaseClass.WriteSummaryTxt(strSumFileName, "Beneficiary Name,Amount,Status,Reason")

            For Each row As DataRow In objFileValidate.DtInput.Select
                objBaseClass.WriteSummaryTxt(strSumFileName, Replace(row("Beneficiary Name").ToString, ",", "") & "," & row("Instrument Amount") & ",Successfull," & row("Reason").ToString)
            Next
            For Each row As DataRow In objFileValidate.DtUnSucInput.Select
                objBaseClass.WriteSummaryTxt(strSumFileName, Replace(row("Beneficiary Name").ToString, ",", "") & "," & "," & row("Instrument Amount") & ",UnSuccessfull," & row("Reason").ToString())
            Next

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Payment_Report")

        End Try

    End Sub

    Private Function GetPaymentAmount(ByVal IsSuccess As Boolean, ByVal PayTpye As String) As Double
        Dim DblAmount As Double = 0

        Try
            If IsSuccess = True Then
                For Each Row As DataRow In objFileValidate.DtInput.Select("[TXN_NO]<>0")
                    DblAmount += Val(Row("Instrument Amount").ToString())
                Next


            Else
                For Each Row As DataRow In objFileValidate.DtUnSucInput.Select("[TXN_NO]<>0")
                    DblAmount += Val(Row("Instrument Amount").ToString())
                Next
            End If

            GetPaymentAmount = DblAmount
            GetPaymentAmount = Convert.ToDecimal(DblAmount).ToString("0.00")

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "GetPaymentAmount")

        End Try

    End Function

End Class
