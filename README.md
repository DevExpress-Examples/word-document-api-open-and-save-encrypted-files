<!-- default file list -->
*Files to look at*:

* [Program.cs](./CS/word-processing-encryption/Program.cs) (VB: [Program.vb](./VB/word-processing-encryption/Program.vb))
<!-- default file list end -->

# How To: Open And Save Encrypted Files

The following sample project shows how to use a Word Processing File API to load and save password-encrypted files. The loaded file is decrypted using the EncryptionSettings properties.When a user re-opens the file with a new password, the RichEditDocumentServer.EncryptedFilePasswordRequested and RichEditDocumentServer.EncryptedFilePasswordCheckFailed events occur. If the user cancels the operation or exceeds the number of attempts to enter the password, RichEditDocumentServer loads an empty file.