Imports System
Imports System.Diagnostics
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet

Namespace EncryptionExample

    Friend Class Program

        Private Shared Property IsValid As Boolean

        Shared Sub Main(ByVal args As String())
            Dim workbook As Workbook = New Workbook()
            workbook.Options.Import.Password = "123"
            workbook.LoadDocument("Documents\encrypted.xlsx")
            AddHandler workbook.EncryptedFilePasswordRequest, AddressOf Workbook_EncryptedFilePasswordRequest
            AddHandler workbook.EncryptedFilePasswordCheckFailed, AddressOf Workbook_EncryptedFilePasswordCheckFailed
            AddHandler workbook.InvalidFormatException, AddressOf Workbook_InvalidFormatException
            Dim encryptionOptions As EncryptionSettings = New EncryptionSettings()
            encryptionOptions.Type = EncryptionType.Strong
            encryptionOptions.Password = "12345"
            Console.WriteLine("Select the file format: XLSX/XLS/XLSB")
            Dim answerFormat As String = Console.ReadLine()?.ToLower()
            Dim documentFormat As DocumentFormat = DocumentFormat.Undefined
            Select Case answerFormat
                Case "xlsx"
                    documentFormat = DocumentFormat.OpenXml
                Case "xls"
                    documentFormat = DocumentFormat.Xls
                Case "xlsb"
                    documentFormat = DocumentFormat.Xlsb
            End Select

            Dim fileName As String = String.Format("EncryptedwithNewPassword.{0}", answerFormat)
            workbook.SaveDocument(fileName, documentFormat, encryptionOptions)
            If IsValid = True Then
                workbook.SaveDocument(fileName, documentFormat)
                Call Process.Start(fileName)
            End If

            Console.WriteLine("The document is saved with new password. Continue? (y/n)")
            Dim answer As String = Console.ReadLine()?.ToLower()
            If Equals(answer, "y") Then
                Console.WriteLine("Re-opening the file...")
                workbook.LoadDocument(fileName)
            End If
        End Sub

        Private Shared Sub Workbook_InvalidFormatException(ByVal sender As Object, ByVal e As SpreadsheetInvalidFormatExceptionEventArgs)
            Console.WriteLine(e.Exception.Message.ToString() & " Press any key to close...")
            Console.ReadKey(True)
        End Sub

        Private Shared Sub Workbook_EncryptedFilePasswordRequest(ByVal sender As Object, ByVal e As EncryptedFilePasswordRequestEventArgs)
            Console.WriteLine("Enter password:")
            e.Password = Console.ReadLine()
            e.Handled = True
            IsValid = True
        End Sub

        Private Shared Sub Workbook_EncryptedFilePasswordCheckFailed(ByVal sender As Object, ByVal e As EncryptedFilePasswordCheckFailedEventArgs)
            Select Case e.Error
                Case SpreadsheetDecryptionError.PasswordRequired
                    Console.WriteLine("You did not enter the password!")
                    e.TryAgain = True
                    e.Handled = True
                Case SpreadsheetDecryptionError.WrongPassword
                    Console.WriteLine("The password is incorrect. Try Again? (y/n)")
                    Dim answer As String = Console.ReadLine()?.ToLower()
                    If Equals(answer, "y") Then
                        e.TryAgain = True
                        e.Handled = True
                    Else
                        IsValid = False
                    End If

            End Select

            IsValid = False
        End Sub
    End Class
End Namespace
