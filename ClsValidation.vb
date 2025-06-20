Imports System
Imports System.Data
Imports System.IO

Public Class ClsValidation

    Implements IDisposable

    Private ObjBaseClass As ClsBase         ''need to be dispose 

    Private DtValidation As DataTable       ''need to be dispose
    Private DtSpCharValidation As DataTable

    Private DtTemp As DataTable             ''need to be dispose

    Public DtInput As DataTable             ''need to be dispose
    Public DtUnSucInput As DataTable        ''need to be dispose
    'Public DtInputTemp As DataTable         ''need to be dispose

    Private ValidationPath As String
    Private ReverseValidationPath As String
    Private SpCharValidationPath As String
    Private StrFilePath As String
    Private InputFilePath As String
    Private AdviceFilePath As String
    'Private ValidationPath As String
    Public ErrorMessage As String

    Public Sub New(ByVal _strFilePath As String, ByVal _SettINIPath As String, ByVal strInputFilePath As String)

        StrFilePath = _strFilePath

        Try
            ObjBaseClass = New ClsBase(_SettINIPath)
            ValidationPath = ObjBaseClass.GetINISettings("General", "Validation", _SettINIPath)

            InputFilePath = strInputFilePath

            DtInput = New DataTable("Input")
            DefineColumnForNestle(DtInput)
            DtUnSucInput = New DataTable("UnSucInput")
            DefineColumnForNestle(DtUnSucInput)

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Constructor")

        End Try

    End Sub

    Private Sub DefineColumnForHLL(ByRef DtInput As DataTable)


        DtInput.Columns.Add(New DataColumn("Record Identifier"))    '0
        DtInput.Columns.Add(New DataColumn("Transaction Type")) '1
        DtInput.Columns.Add(New DataColumn("Beneficiary/Vendor Code"))  '2
        DtInput.Columns.Add(New DataColumn("Mail To"))  '3
        DtInput.Columns.Add(New DataColumn("Bene Mailing Address")) '4
        DtInput.Columns.Add(New DataColumn("Bene Bank + A/c #"))    '5
        DtInput.Columns.Add(New DataColumn("Pay To"))   '6
        DtInput.Columns.Add(New DataColumn("Instrument/Cheque Number")) '7
        DtInput.Columns.Add(New DataColumn("Transaction/Cheque Date"))  '8
        DtInput.Columns.Add(New DataColumn("Instrument Amount"))    '9
        DtInput.Columns.Add(New DataColumn("Hundi"))    '10
        DtInput.Columns.Add(New DataColumn("Currency Code"))    '11
        DtInput.Columns.Add(New DataColumn("Payment Location")) '12
        DtInput.Columns.Add(New DataColumn("Annexure Text Link Ref"))   '13
        DtInput.Columns.Add(New DataColumn("Payment Date")) '14
        DtInput.Columns.Add(New DataColumn("Number of Records in Annexure Text"))   '15
        DtInput.Columns.Add(New DataColumn("Print Location"))   '16
        DtInput.Columns.Add(New DataColumn("Bene Bank IFSC code"))  '17
        DtInput.Columns.Add(New DataColumn("Bene A/c type"))    '18
        DtInput.Columns.Add(New DataColumn("Bene Bank Name"))  '19
        DtInput.Columns.Add(New DataColumn("Bene Bank  A/c"))   '20
        DtInput.Columns.Add(New DataColumn("Bene Bank Branch")) '21
        DtInput.Columns.Add(New DataColumn("Bene Bank Location"))   '22
        DtInput.Columns.Add(New DataColumn("Bene Mail Id")) '23
        DtInput.Columns.Add(New DataColumn("Invoice No")) '24
        DtInput.Columns.Add(New DataColumn("Gross Amount")) '25
        DtInput.Columns.Add(New DataColumn("Deductions")) '26
        DtInput.Columns.Add(New DataColumn("Net Amount")) '27

        'DtInput.Columns.Add(New DataColumn("TXN_NO"))     '26 26
        DtInput.Columns.Add(New DataColumn("TXN_NO", System.Type.GetType("System.Int32")))    '28
        'DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   '27 27
        DtInput.Columns.Add(New DataColumn("SUBTXN_NO", System.Type.GetType("System.Int32")))   '29
        DtInput.Columns.Add(New DataColumn("REASON"))     '30




    End Sub

    Private Sub DefineColumnForNestle(ByRef DtInput As DataTable)


        DtInput.Columns.Add(New DataColumn("Customer Reference No"))  '0
        DtInput.Columns.Add(New DataColumn("Beneficiary Name"))  '1
        DtInput.Columns.Add(New DataColumn("Beneficiary Account Number")) '2
        DtInput.Columns.Add(New DataColumn("IFSC Code"))  '3
        DtInput.Columns.Add(New DataColumn("Account Type"))  '4
        DtInput.Columns.Add(New DataColumn("Instrument Amount"))  '5
        DtInput.Columns.Add(New DataColumn("Value Date"))  '6

        'DtInput.Columns.Add(New DataColumn("EmailId"))  '6



        'DtInput.Columns.Add(New DataColumn("TXN_NO"))   '29 28
        DtInput.Columns.Add(New DataColumn("TXN_NO", System.Type.GetType("System.Int32")))   '7
        'DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))    '30 29
        'DtInput.Columns.Add(New DataColumn("SUBTXN_NO", System.Type.GetType("System.Int32")))   '8
        DtInput.Columns.Add(New DataColumn("REASON"))   '8
        'DtInput.Columns.Add(New DataColumn("Exception")) '9



    End Sub

    Public Function CheckValidateFile() As Boolean

        Try
            If Not File.Exists(StrFilePath) Then
                Call ObjBaseClass.Handle_Error(New ApplicationException("Input file path is incorrect or not file found. [" & StrFilePath & "]"), "ClsValidation", -123, "CheckValidateFile")
                CheckValidateFile = False
                Exit Function
            End If


            If File.Exists(strValidationPath) Then
                CheckValidateFile = Validate()
            Else
                Call ObjBaseClass.Handle_Error(New ApplicationException("Validation file path is incorrect. [" & strValidationPath & "]"), "ClsValidation", -123, "CheckValidateFile")
            End If



        Catch ex As Exception
            CheckValidateFile = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "CheckValidateFile")
        End Try

    End Function

    Private Function RemoveBlankRow(ByRef _DtTemp As DataTable)
        'To Remove Blank Row Exists in DataTable
        Dim blnRowBlank As Boolean

        Try
            For Each vRow As DataRow In _DtTemp.Rows
                blnRowBlank = True

                For intCol As Int32 = 0 To _DtTemp.Columns.Count - 1
                    If vRow.Item(intCol).ToString().Trim() <> "" Then
                        blnRowBlank = False
                        Exit For
                    End If
                Next

                If blnRowBlank = True Then
                    _DtTemp.Rows(vRow.Table.Rows.IndexOf(vRow)).Delete()
                End If

            Next
            _DtTemp.AcceptChanges()

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveBlankRow")

        End Try

    End Function

    Private Function Validate() As Boolean

        Validate = False

        Dim DrValidOutputColumn() As DataRow = Nothing

        Dim InputLineNumber As Int32 = 0
        Dim StrDataRow(8) As String

        Dim ArrLineData As Object
        Dim intPosField As Integer
        Dim intPosition As Integer
        Dim intPosLength As Integer


        Dim intI As Integer

        Dim TXN_NO As Integer



        Try
            ErrorMessage = ""

            DtValidation = ObjBaseClass.GetDataTable_ExcelSheet(ValidationPath, "Sheet1")
            RemoveBlankRow(DtValidation)
            DrValidOutputColumn = DtValidation.Select("[SR NO] <> 0  ", "[SR NO]")

            DtTemp = ObjBaseClass.GetDataTable_ExcelSQL(strInputFolderPath & "\" & gstrInputFile, 1, "")
            RemoveBlankRow(DtTemp)

            If DtValidation.Rows.Count >= 7 Then
                InputLineNumber = 1
                TXN_NO = 0


                intPosition = 0
                intPosField = 2
                intPosLength = 0



                TXN_NO = 0

                For Each dtRow As DataRow In DtTemp.Rows

                    If dtRow("BANK KEY").ToString().ToUpper() <> "CASH".ToString().ToUpper() OrElse dtRow("BANK").ToString().ToUpper() <> "CASH".ToString().ToUpper() Then
                        ArrLineData = dtRow.ItemArray

                        intPosField = 3
                        intPosition = 2
                        intPosLength = 4

                        ClearArray(StrDataRow)
                        TXN_NO += 1



                        For strIndex As Int32 = 0 To DrValidOutputColumn.Length - 1

                            If Val(DrValidOutputColumn(strIndex)(intPosField).ToString.Trim()) <> 0 Then

                                StrDataRow(strIndex) = GetValueFormArray(ArrLineData, Val(DrValidOutputColumn(strIndex)(intPosField).ToString.Trim().ToString.Trim())).Trim()

                                If StrDataRow(strIndex) = "~ERROR~" Then
                                    StrDataRow(8) = StrDataRow(8).ToString() & "Line No.:" & InputLineNumber & ", Invalid input field position defined in validation file. [ Reference : Input data array length = " & ArrLineData.Length & " , Field Position = " & Val(DrValidOutputColumn(strIndex)(intPosField).ToString.Trim()) & "]" & "| "
                                End If
                            Else
                                StrDataRow(strIndex) = ""
                            End If


                            If StrDataRow(strIndex) <> "" Then
                                StrDataRow(strIndex) = RemoveJunk(StrDataRow(strIndex).ToString)
                            End If
                        Next


                        intI = 0
                        For Each VROW As DataRow In DtValidation.Rows


                            If VROW(1).ToString().Trim().ToUpper() = "Instrument Amount".ToUpper() Then
                                StrDataRow(intI) = Val(StrDataRow(intI)).ToString("0.00")
                                If Val(StrDataRow(intI)) <= 0 Then
                                    StrDataRow(8) = StrDataRow(8) & "Line NO: " & InputLineNumber & " Column Name [" & VROW(1).ToString() & "] Column Value [" & StrDataRow(intI) & "] is either blank zero or negative amount" & "| "  ''Reason
                                End If


                            ElseIf VROW(1).ToString().Trim().ToUpper() = "Account Type".ToUpper() Then

                                StrDataRow(intI) = "11"

                            ElseIf VROW(1).ToString().Trim().ToUpper() = "Value Date".ToUpper() Then

                                '    StrDataRow(intI) = Format(CDate(Now.Date()), "yyyyMMdd")
                                Dim inputDate As String = StrDataRow(intI)
                                Dim dt As DateTime

                                If DateTime.TryParseExact(inputDate, "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, dt) Then
                                    ' Format is correct; convert to "yyyyMMdd"
                                    Dim formattedDate As String = dt.ToString("yyyyMMdd")
                                    StrDataRow(intI) = formattedDate

                                Else
                                    ' Invalid date format
                                    StrDataRow(8) = StrDataRow(8) & "Line No: " & InputLineNumber & "  Column Name [" & VROW(1).ToString & "] value [" & StrDataRow(intI) & "] of Invalid date format" & "|"
                                    Console.WriteLine("Invalid date format: " & inputDate)
                                End If
                            End If

                            '' Checking Duplication Account Number.
                            'Dim drMasterRow As DataRow()
                            'If VROW(1).ToString().Trim().ToUpper() = "Beneficiary Account Number".ToUpper() Then
                            '    If (StrDataRow(intI).ToString.Trim <> "") Then
                            '        drMasterRow = DtInput.Select("[Beneficiary Account Number]='" & StrDataRow(intI).ToString & "'")
                            '        If drMasterRow.Length > 0 Then
                            '            StrDataRow(9) = StrDataRow(9) & "Line No: " & InputLineNumber & "  Column Name [" & VROW(1).ToString & "] value [" & StrDataRow(intI) & "] of duplicate Beneficiary Account Number" & "|"
                            '        End If
                            '    End If

                            'End If


                            '---Checking Length
                            If intPosLength <> 0 Then
                                If StrDataRow(intI).Length > VROW(intPosLength) Then
                                    StrDataRow(8) = StrDataRow(8) & "Line No: " & InputLineNumber & "  Column Name [" & VROW(1).ToString & "] value [" & StrDataRow(intI) & "] exceeds " & VROW(intPosLength) & " characters" & "|"
                                End If
                            End If
                            '-Mandatory Fields-
                            If intPosition <> 0 Then
                                If StrDataRow(intI).ToString() = "" And VROW(intPosition).ToString().Trim().ToUpper() = "M" Then
                                    StrDataRow(8) = StrDataRow(8) & "Line NO: " & InputLineNumber & " Column Name [" & VROW(1).ToString & "] is Mandatory Field but value is Blank" & "| "

                                End If
                            End If


                            intI += 1
                        Next
                        StrDataRow(7) = TXN_NO

                        If StrDataRow(8) = "" Then
                            DtInput.Rows.Add(StrDataRow)
                        Else
                            DtUnSucInput.Rows.Add(StrDataRow)
                        End If
                    End If

                    InputLineNumber += 1
                Next


            Else
                Call ObjBaseClass.Handle_Error(New ApplicationException("Validation is not maintained properly in " & Path.GetFileName(ValidationPath) & " validation file. It must be atleast 24 columns defination."), "ClsValidation", -123, "Validate")
            End If

            Validate = True
ValidateLine:
        Catch ex As Exception
            Validate = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Validate")

        Finally

            DrValidOutputColumn = Nothing
            ObjBaseClass.ObjectDispose(DtValidation)
            ObjBaseClass.ObjectDispose(DtTemp)

        End Try

    End Function


    Public Function CheckValidEmailId(ByVal strMail As String) As Boolean
        Dim objMail As Object
        Try
            CheckValidEmailId = True
            objMail = New System.Net.Mail.MailAddress(strMail)
        Catch ex As Exception
            CheckValidEmailId = False
        Finally
            ObjBaseClass.ObjectDispose(objMail)
        End Try


    End Function


    Public Shared Function CheckIfAlphaNumeric(ByVal Str As String) As Boolean
        Dim IsAlphaNumeric As Boolean = True
        Dim c As Char

        Try

            For i As Integer = 0 To Str.Length - 1
                c = Str.Chars(i)
                If IsNumeric(c) Or c Like "[A-Z]" Or c Like "[a-z]" Then
                    IsAlphaNumeric = True
                Else
                    IsAlphaNumeric = False
                End If
            Next

        Catch ex As Exception
            IsAlphaNumeric = False
        End Try

        Return IsAlphaNumeric

    End Function

    Public Function SpCharValidation(ByVal StringValue As String, ByRef _dtSpChar As DataTable) As String
        'Public Function SpCharValidation(ByVal StringValue As String) As String

        ''Added by Jaiwant dtd  03-Dec-2010
        Dim ArrSpChar(0) As String
        Dim intSpCharRow As Integer
        ''---

        ClearArray(ArrSpChar)
        Array.Resize(ArrSpChar, _dtSpChar.Select.Length)
        intSpCharRow = 0

        For Each SVRow As DataRow In _dtSpChar.Rows
            ArrSpChar(intSpCharRow) = SVRow(0).ToString
            intSpCharRow += 1
        Next

        ''Added By Jaiwant dtd  03-Dec-2010 ''For All Special Characters
        Dim StrOriginalValue As String = ""
        'Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&", " "} ''Commented by Lakshmi dtd 22-03-2012
        Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&"} ''Added by Lakshmi dtd 22-03-2012

        Try
            'To remove special chars from array which need to ignore.
            For iIChar As Int16 = 0 To ArrSpChar.Length - 1
                For iSChar As Int16 = 0 To arrSpecialChar.Length - 1
                    If ArrSpChar(iIChar) = arrSpecialChar(iSChar) Then
                        arrSpecialChar(iSChar) = Nothing
                    End If
                Next
            Next

            SpCharValidation = ""
            Dim i As Integer
            For i = 0 To arrSpecialChar.Length - 1
                If InStr(StringValue, arrSpecialChar(i), CompareMethod.Binary) <> 0 Then
                    SpCharValidation = SpCharValidation & arrSpecialChar(i)
                End If
            Next

            Return SpCharValidation

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "SpCharValidation")

        End Try
    End Function

    Private Sub AddRowsToDataTable(ByRef dtTempData As DataTable)

        Try
            Dim IsInvalidTransaction As Boolean

            If dtTempData Is Nothing Then Exit Sub

            ''Validating Successfull Set of Transactions only.
            For Each _Row As DataRow In dtTempData.Select
                IsInvalidTransaction = False

                For Each vldROW As DataRow In dtTempData.Select("TXN_NO ='" & _Row("TXN_NO") & "' and Reason <> ''")
                    IsInvalidTransaction = True
                    Exit For
                Next

                If IsInvalidTransaction = True Then
                    Dim unsucRow As DataRow = DtUnSucInput.NewRow
                    unsucRow = _Row
                    DtUnSucInput.Rows.Add(unsucRow.ItemArray)
                Else
                    Dim sucRow As DataRow = DtInput.NewRow
                    sucRow = _Row
                    DtInput.Rows.Add(sucRow.ItemArray)
                End If
            Next

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "AddRowsToDataTable")

        End Try

    End Sub

    Private Function GetSubstring(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring = "~ERROR~"
                Else
                    ''Added By Jaiwant dtd 29-dec-2010
                    ''GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    If Len(Mid(pStrValue, pStartPos + 1, Len(pStrValue))) < (pEndPos - pStartPos) Then
                        GetSubstring = Mid(pStrValue, pStartPos + 1, pEndPos - pStartPos)
                    Else
                        GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    End If
                End If
            End If

        Catch ex As Exception
            GetSubstring = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")

        End Try
    End Function

    Private Function GetValueFormArray(ByRef pArray() As Object, ByVal pPosition As Int16) As String

        Try

            If pArray.Length >= pPosition Then
                GetValueFormArray = pArray(pPosition - 1).ToString()
            Else
                ErrorMessage = "~ERROR~"
            End If


        Catch ex As Exception
            ErrorMessage = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetValueFormArray")

        End Try

    End Function

    Private Function GetShort(ByVal pStr As String) As String

        pStr = pStr.ToUpper

        If InStr(pStr, "D") > 0 Then
            GetShort = "D"
        ElseIf InStr(pStr, "M") > 0 Then
            GetShort = "M"
        ElseIf InStr(pStr, "Y") > 0 Then
            GetShort = "Y"
        End If

    End Function

    Private Sub ClearArray(ByRef ArrRow() As String)

        Try
            For i As Integer = 0 To ArrRow.Length - 1
                ArrRow(i) = ""
            Next

        Catch ex As Exception

        End Try
    End Sub

    Public Function GetColumValue(ByVal strString As String, ByVal intStart As Integer, ByVal intEnd As Integer)

        Try

            intStart = intStart - 1
            GetColumValue = strString.Substring(intStart, intEnd - intStart).Trim()

        Catch ex As Exception
            GetColumValue = ""
        End Try
    End Function

    Public Function RemoveJunk(ByVal sText As String) As String
        ''-To remove Junk Characters-
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsAlpha(sChar) Then
                        RemoveJunk = RemoveJunk + sChar
                    End If
                Next
            End If

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsRBIValidation", "RemoveJunk")

        End Try

    End Function

    Private Function IsAlpha(ByVal sChr As String) As Boolean
        '-To remove Junk Characters-

        IsAlpha = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" _
        Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[;]" Or sChr Like "[:]" _
        Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[?]" Or sChr Like "[/]" _
        Or sChr Like "[']" Or sChr Like "[""]" Or sChr Like "[|]" Or sChr Like "[\]" _
        Or sChr Like "[{]" Or sChr Like "[[]" Or sChr Like "[}]" Or sChr Like "[]]" _
        Or sChr Like "[+]" Or sChr Like "[=]" Or sChr Like "[_]" Or sChr Like "[-]" _
        Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[*]" Or sChr Like "[&]" _
        Or sChr Like "[^]" Or sChr Like "[%]" Or sChr Like "[$]" Or sChr Like "[#]" _
        Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[`]" Or sChr Like "[~]" Or sChr Like "[ ]"

    End Function

    Private Function GetValidateDate(ByRef pStrDate As String) As Boolean

        Try

            strInputDateFormat = strInputDateFormat.ToUpper()

            Dim TmpstrInputDateFormat() As String
            Dim TempStrDateValue() As String = pStrDate.Split(" ")

            If InStr(TempStrDateValue(0), "/") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("/")
                TmpstrInputDateFormat = strInputDateFormat.Split("/")
            ElseIf InStr(TempStrDateValue(0), "-") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("-")
                TmpstrInputDateFormat = strInputDateFormat.Split("-")
            End If

            Dim HsUserDate As New Hashtable
            Dim HsSystemDate As New Hashtable
            Dim StrFinalDate As String

            If TempStrDateValue.Length = 3 Then
                For IntStr As Integer = 0 To TempStrDateValue.Length - 1
                    HsUserDate.Add(GetShort(TmpstrInputDateFormat(IntStr)), TempStrDateValue(IntStr))
                Next
                Dim SysDate() As String
                Dim dtSys As String = System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern.ToUpper()
                If InStr(dtSys, "/") > 0 Then
                    SysDate = dtSys.Split("/")
                ElseIf InStr(dtSys, "-") > 0 Then
                    SysDate = dtSys.Split("-")
                End If

                StrFinalDate = ""
                For IntStr As Integer = 0 To SysDate.Length - 1
                    If StrFinalDate = "" Then
                        StrFinalDate += HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    Else
                        StrFinalDate += "/" & HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    End If
                Next

                Try
                    pStrDate = CDate(StrFinalDate)

                    GetValidateDate = True

                Catch ex As Exception
                    GetValidateDate = False

                End Try
            Else
                GetValidateDate = False
            End If

        Catch ex As Exception
            GetValidateDate = False

        End Try

    End Function

#Region " IDisposable Support "
    Public Sub Dispose() Implements IDisposable.Dispose

        If Not ObjBaseClass Is Nothing Then ObjBaseClass.Dispose()
        If Not DtValidation Is Nothing Then DtValidation.Dispose()
        If Not DtInput Is Nothing Then DtInput.Dispose()
        If Not DtUnSucInput Is Nothing Then DtUnSucInput.Dispose()
        If Not DtTemp Is Nothing Then DtTemp.Dispose()

        ObjBaseClass = Nothing
        DtValidation = Nothing
        DtInput = Nothing
        DtUnSucInput = Nothing
        DtTemp = Nothing

        GC.SuppressFinalize(Me)

    End Sub
#End Region

End Class
