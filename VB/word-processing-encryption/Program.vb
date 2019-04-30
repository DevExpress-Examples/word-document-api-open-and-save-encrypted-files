Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Diagnostics

Namespace word_processing_encryption
	Friend Class Program
		Private Shared Property IsValid() As Boolean


		Shared Sub Main(ByVal args() As String)
			Dim server As New RichEditDocumentServer()
			AddHandler server.EncryptedFilePasswordRequested, AddressOf Server_EncryptedFilePasswordRequested
			AddHandler server.EncryptedFilePasswordCheckFailed, AddressOf Server_EncryptedFilePasswordCheckFailed
			AddHandler server.InvalidFormatException, AddressOf Server_InvalidFormatException

			server.Options.Import.EncryptionPassword = "test"
			server.LoadDocument("Documents//testEncrypted.docx")

            Dim encryptionOptions As New EncryptionSettings()
            encryptionOptions.Type = EncryptionType.Strong
			encryptionOptions.Password = "12345"

			Console.WriteLine("Select the file format: DOCX/DOC")
			Dim answerFormat As String = Console.ReadLine()?.ToLower()
			Dim documentFormat As DocumentFormat
			If answerFormat = "docx" Then
				documentFormat = DocumentFormat.OpenXml
			Else
				documentFormat = DocumentFormat.Doc
			End If

			Dim fileName As String = String.Format("EncryptedwithNewPassword.{0}", answerFormat)

			server.SaveDocument(fileName, documentFormat, encryptionOptions)

			Console.WriteLine("The document is saved with new password. Continue? (y/n)")
			Dim answer As String = Console.ReadLine()?.ToLower()
			If answer = "y" Then
				Console.WriteLine("Re-opening the file...")
				server.LoadDocument(fileName)
			End If
			If IsValid = True Then
				server.SaveDocument(fileName, documentFormat)
			End If
				Process.Start(fileName)
		End Sub


		Private Shared Sub Server_InvalidFormatException(ByVal sender As Object, ByVal e As RichEditInvalidFormatExceptionEventArgs)
			Dim server As RichEditDocumentServer = DirectCast(sender, RichEditDocumentServer)
			server.SaveDocument("EmptyFile.docx", DocumentFormat.OpenXml)
			Process.Start("EmptyFile.docx")
		End Sub

		Private Shared Sub Server_EncryptedFilePasswordCheckFailed(ByVal sender As Object, ByVal e As EncryptedFilePasswordCheckFailedEventArgs)
			Select Case e.Error
				Case RichEditDecryptionError.PasswordRequired
					Console.WriteLine("You did not enter the password!")
					e.TryAgain = True
					e.Handled = True
				Case RichEditDecryptionError.WrongPassword
					Console.WriteLine("The password is incorrect. Try Again? (y/n)")
					Dim answer As String = Console.ReadLine()?.ToLower()
					If answer = "y" Then
						e.TryAgain = True
						e.Handled = True

					Else
						Console.WriteLine("Password check failed. Loading an empty file...")
					End If
			End Select

			Program.IsValid = False
		End Sub

		Private Shared Sub Server_EncryptedFilePasswordRequested(ByVal sender As Object, ByVal e As EncryptedFilePasswordRequestedEventArgs)
			Console.WriteLine("Enter password:")
			e.Password = Console.ReadLine()
			e.Handled = True
			IsValid = True
		End Sub

	End Class
End Namespace
