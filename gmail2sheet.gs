function exportInboxEmailsToSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange("A1:E1").setValues([["From", "Subject", "Date", "Message", "Link"]]);
  var data = [];
  var pageSize = 100;
  var start = 0;
  var threads;
  do {
    threads = GmailApp.getInboxThreads(start, pageSize);
    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var row = [message.getFrom(), message.getSubject(), message.getDate(), message.getPlainBody()];
        var messageId = message.getId();
        var messageUrl = "https://mail.google.com/mail/u/0/#inbox/" + messageId;
        row.push(messageUrl);
        data.push(row);
      }
    }
    start += pageSize;
  } while (threads.length == pageSize && data.length < 2000);
  var range = sheet.getRange(2, 1, data.length, data[0].length);
  range.setValues(data);
}


