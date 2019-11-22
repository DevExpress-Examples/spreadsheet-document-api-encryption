using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;

namespace EncryptionExample
{
    class Program
    {
        static bool IsValid { get; set; }
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook();
            workbook.Options.Import.Password = "123";
            workbook.LoadDocument("Documents\\encrypted.xlsx");

            workbook.EncryptedFilePasswordRequest += Workbook_EncryptedFilePasswordRequest;
            workbook.EncryptedFilePasswordCheckFailed += Workbook_EncryptedFilePasswordCheckFailed;
            workbook.InvalidFormatException += Workbook_InvalidFormatException;

            EncryptionSettings encryptionOptions = new EncryptionSettings();
            encryptionOptions.Type = EncryptionType.Strong;
            encryptionOptions.Password = "12345";

            Console.WriteLine("Select the file format: XLSX/XLS/XLSB");
            string answerFormat = Console.ReadLine()?.ToLower();
            DocumentFormat documentFormat = DocumentFormat.Undefined;
            switch (answerFormat)
            {
                case "xlsx":
                    documentFormat = DocumentFormat.OpenXml;
                    break;
                case "xls":
                    documentFormat = DocumentFormat.Xls;
                    break;
                case "xlsb":
                    documentFormat = DocumentFormat.Xlsb;
                    break;
            }
            string fileName = String.Format("EncryptedwithNewPassword.{0}", answerFormat);
            workbook.SaveDocument(fileName, documentFormat, encryptionOptions);

            if (IsValid == true)
            {
                workbook.SaveDocument(fileName, documentFormat);
                Process.Start(fileName);
            }

            Console.WriteLine("The document is saved with new password. Continue? (y/n)");
            string answer = Console.ReadLine()?.ToLower();

            if (answer == "y")
            {
                Console.WriteLine("Re-opening the file...");
                workbook.LoadDocument(fileName);
            }

        }

        private static void Workbook_InvalidFormatException(object sender, SpreadsheetInvalidFormatExceptionEventArgs e)
        {
            Console.WriteLine(e.Exception.Message.ToString() + " Press any key to close...");
            Console.ReadKey(true);
        }

        private static void Workbook_EncryptedFilePasswordRequest(object sender, EncryptedFilePasswordRequestEventArgs e)
        {
            Console.WriteLine("Enter password:");
            e.Password = Console.ReadLine();
            e.Handled = true;
            IsValid = true;
        }

        private static void Workbook_EncryptedFilePasswordCheckFailed(object sender, EncryptedFilePasswordCheckFailedEventArgs e)
        {
            switch (e.Error)
            {
                case SpreadsheetDecryptionError.PasswordRequired:
                    Console.WriteLine("You did not enter the password!");
                    e.TryAgain = true;
                    e.Handled = true;
                    break;
                case SpreadsheetDecryptionError.WrongPassword:
                    Console.WriteLine("The password is incorrect. Try Again? (y/n)");
                    string answer = Console.ReadLine()?.ToLower();
                    if (answer == "y")
                    {
                        e.TryAgain = true;
                        e.Handled = true;
                    }

                    else
                    {
                        IsValid = false;
                    }
                    break;
            }

            Program.IsValid = false;
        }
    }

}

