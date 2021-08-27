<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/184084298/19.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830413)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Program.cs](./CS/word-processing-encryption/Program.cs) (VB: [Program.vb](./VB/word-processing-encryption/Program.vb))
<!-- default file list end -->

# How To: Open And Save Encrypted Files

The following sample project shows how to use a Word Processing File API to load and save password-encrypted files. The loaded file is decrypted using the EncryptionSettings properties. When a user re-opens the file with a new password, the [RichEditDocumentServer.EncryptedFilePasswordRequested](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer.EncryptedFilePasswordRequested?v=19.1) and [RichEditDocumentServer.EncryptedFilePasswordCheckFailed](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer.EncryptedFilePasswordCheckFailed?v=19.1) events occur. If the user cancels the operation or exceeds the number of attempts to enter the password, [RichEditDocumentServer](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer?v=19.1) shows an exception message.
