function doGet() {
  return HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Save Emails as PDF');
}


function saveEmailsAsPDF(labelName, folderName) {
  try {
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      throw 'Label "' + labelName + '" does not exist.';
    }

    var threads = label.getThreads();
    if (threads.length === 0) {
      throw 'No emails found in label "' + labelName + '".';
    }

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();

      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var pdf = messageToPdf(message);
        pdf.setName(formatDate(message, 'yyyy/MM/dd h:mm a - ') + pdf.getName());
        getFolder(folderName).createFile(pdf);
        var attachments = message.getAttachments();

        attached_PDF = messageGetPdfAttachment(message);
        if (attached_PDF){
          attached_PDF.setName(pdf.getName() + '| Attached File: ' + attached_PDF.getName());
          getFolder(folderName).createFile(attached_PDF);
        }  
      }
    }

    return 'Saved ' + threads.length + ' threads to folder "' + folderName + '".';
  } catch (error) {
    throw new Error('Error: ' + error);
  }
}



