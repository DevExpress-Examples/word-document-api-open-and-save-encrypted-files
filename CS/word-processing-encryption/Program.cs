using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Diagnostics;

namespace word_processing_encryption
{
    class Program
    {
        static bool IsValid { get; set; }


        static void Main(string[] args)
        {
            RichEditDocumentServer server = new RichEditDocumentServer();
            server.EncryptedFilePasswordRequested += Server_EncryptedFilePasswordRequested;
            server.EncryptedFilePasswordCheckFailed += Server_EncryptedFilePasswordCheckFailed;
            server.InvalidFormatException += Server_InvalidFormatException;

            server.Options.Import.EncryptionPassword = "test";
            server.LoadDocument("Documents//testEncrypted.docx");

            EncryptionSettings encryptionOptions = new EncryptionSettings();
            encryptionOptions.Type = EncryptionType.Strong;
            encryptionOptions.Password = "12345";

            Console.WriteLine("Select the file format: DOCX/DOC");
            string answerFormat = Console.ReadLine()?.ToLower();
            DocumentFormat documentFormat;
            if (answerFormat == "docx")
                documentFormat = DocumentFormat.OpenXml;
            else
                documentFormat = DocumentFormat.Doc;

            string fileName = String.Format("EncryptedwithNewPassword.{0}", answerFormat);

            server.SaveDocument(fileName, documentFormat, encryptionOptions);

            Console.WriteLine("The document is saved with new password. Continue? (y/n)");
            string answer = Console.ReadLine()?.ToLower();
            if (answer == "y")
            {
                Console.WriteLine("Re-opening the file...");
                server.LoadDocument(fileName);
            }
            if (IsValid == true)
                server.SaveDocument(fileName, documentFormat);
                Process.Start(fileName);
        }


        private static void Server_InvalidFormatException(object sender, RichEditInvalidFormatExceptionEventArgs e)
        {
            RichEditDocumentServer server = (RichEditDocumentServer)sender;
            server.SaveDocument("EmptyFile.docx", DocumentFormat.OpenXml);
            Process.Start("EmptyFile.docx");
        }

        private static void Server_EncryptedFilePasswordCheckFailed(object sender, EncryptedFilePasswordCheckFailedEventArgs e)
        {
            switch (e.Error)
            {
                case RichEditDecryptionError.PasswordRequired:
                    Console.WriteLine("You did not enter the password!");
                    e.TryAgain = true;
                    e.Handled = true;
                    break;
                case RichEditDecryptionError.WrongPassword:
                    Console.WriteLine("The password is incorrect. Try Again? (y/n)");
                    string answer = Console.ReadLine()?.ToLower();
                    if (answer == "y")
                    {
                        e.TryAgain = true;
                        e.Handled = true;
                    }

                    else
                    {
                        Console.WriteLine("Password check failed. Loading an empty file...");
                    }
                    break;
            }

            Program.IsValid = false;
        }

        private static void Server_EncryptedFilePasswordRequested(object sender, EncryptedFilePasswordRequestedEventArgs e)
        {
            Console.WriteLine("Enter password:");
            e.Password = Console.ReadLine();
            e.Handled = true;
            IsValid = true;
        }

    }
}
