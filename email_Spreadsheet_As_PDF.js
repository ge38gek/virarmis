function emailSpreadsheetAsPDF() {
  DocumentApp.getActiveDocument();
  DriveApp.getFiles();

  // This is the link to my spreadsheet with the Form responses and the Invoice Template sheets
  // Add the link to your spreadsheet here 
  // or you can just replace the text in the link between "d/" and "/edit"
  // In my case is the text: 17I8-QDce0Nug7amrZeYTB3IYbGCGxvUj-XMt8uUUyvI
  const ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1uOeygWpKSqici3ipDEQAbVU3XwIQVVqGb_H7Y2Ccjsg/edit");

  // We are going to get the email address from the cell "B7" from the "Invoice" sheet
  // Change the reference of the cell or the name of the sheet if it is different
  const value = ss.getSheetByName("Order").getRange("I17").getValue();
  const email = value.toString();

  // Subject of the email message
  const subject = 'Your Order';

  // Email Text. You can add HTML code here - see ctrlq.org/html-mail
  const body = "Sent via Generate Invoice from Google Form and print/email it";

  // Again, the URL to your spreadsheet but now with "/export" at the end
  // Change it to the link of your spreadsheet, but leave the "/export"
  const url = 'https://docs.google.com/spreadsheets/d/1uOeygWpKSqici3ipDEQAbVU3XwIQVVqGb_H7Y2Ccjsg/export?';

  const exportOptions =
    'exportFormat=pdf&format=pdf' + // export as pdf
    '&size=letter' + // paper size letter / You can use A4 or legal
    '&portrait=true' + // orientation portal, use false for landscape
    '&fitw=false' + // fit to page width false, to get the actual size
    '&sheetnames=false&printtitle=false' + // hide optional headers and footers
    '&pagenumbers=false&gridlines=false' + // hide page numbers and gridlines
    '&fzr=false' + // do not repeat row headers (frozen rows) on each page
    '&gid=1648063207'; // the sheet's Id. Change it to your sheet ID.
  
  // You can find the sheet ID in the link bar. 
  // Select the sheet that you want to print and check the link,
  // the gid  number of the sheet is on the end of your link.
  
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  // Generate the PDF file
  var response = UrlFetchApp.fetch(url+exportOptions, params).getBlob();
  
  // Send the PDF file as an attachement 
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments: [{
            fileName: "Order" + ".pdf",
            content: response.getBytes(),
            mimeType: "application/pdf"
        }]
    });

  // Save the PDF to Drive. The name of the PDF is going to be the name of the Company (cell B5)
  const nameFile = ss.getSheetByName("Order").getRange("E21").getValue().toString() +".pdf"
  DriveApp.createFile(response.setName(nameFile));
} 