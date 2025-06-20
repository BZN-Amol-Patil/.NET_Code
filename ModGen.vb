Option Explicit On

Module ModGen

    Public gstrInputFolder As String
    Public gstrInputFile As String
    Public gstrOutputFile As String
    Public gstrAdviceFolder As String
    Public gstrAdviceFile As String

    '-Standard-
    Public blnErrorLog As Boolean
    Public strSettingClientCode As String
    Public strSettingClientName As String

    '-General-
    Public strAuditFolderPath As String
    Public strErrorFolderPath As String
    Public strInputFolderPath As String
    Public strAdviceFilePath As String
    Public strRBIInputFolderPath As String

    Public strOutputFolderPath As String
    Public strReportFolderPath As String
    Public strValidationPath As String
    Public strTempFolderPath As String
    Public strAdviceFolderPath As String
    Public strArchivedFolderSuc As String
    Public strArchivedFolderUnSuc As String


    Public strSpCharValidation As String
    Public NoOfRecords As Double
    Public strProceed As String
    Public strRTGSLimit As String
    Public strBeneCode As String
    Public strPrintLocation As String
    Public strPaymentNo As String

    Public gstrOutputFileListing(0) As String
    Public gstrOutputFileCount As Integer

    '-Client Details-
    Public strClientCode As String
    Public strClientName As String
    Public strInputDateFormat As String
    Public strDomainID As String

    Public strAccounttype As String
    Public nmConventionSFTP As String = ""
    '-Encryption-
    Public strEncrypt As String
    Public strBatchFilePath As String
    Public strPICKDIRpath As String
    Public strDROPDIRPath As String

End Module
