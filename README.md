<!-- default badges list -->
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830413)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# Word Processing Document API - Open and Save Encrypted Files

The following sample project shows how to use a Word Processing File API to load and save password-encrypted files. The loaded file is decrypted using the `EncryptionSettings` properties. When a user re-opens the file with a new password, the [RichEditDocumentServer.EncryptedFilePasswordRequested](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer.EncryptedFilePasswordRequested) and [RichEditDocumentServer.EncryptedFilePasswordCheckFailed](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer.EncryptedFilePasswordCheckFailed) events occur. If the user cancels the operation or exceeds the number of attempts to enter the password, [RichEditDocumentServer](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer) shows an exception message.

> [!Important]  
> The Universal Subscription or an additional Office File API Subscription is required to use this example in production code. For pricing information, please refer to the [DevExpress Subscription](https://www.devexpress.com/Subscriptions/) page.

## Files to Review

* [Program.cs](./CS/Program.cs) (VB: [Program.vb](./VB/Program.vb))

## Documentation

* [Protection in Word Documents](https://docs.devexpress.com/OfficeFileAPI/120152/word-processing-document-api/document-protection)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=word-document-api-open-and-save-encrypted-files&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=word-document-api-open-and-save-encrypted-files&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
