function sortEmails() {
    var table, tr, td, cell, i, j;

    table = document.getElementById("emailTable");
    tr = table.getElementsByTagName("tr");


    for (i = 0; i < tr.length; i++) {
        // Hide the row initially.
        tr[i].style.display = "none";

        td = tr[i].getElementsByTagName("td");


        for (var j = 0; j < td.length; j++) {
            cell = tr[i].getElementsByTagName("td")[j];
            if (cell) {
                for (var k = 0; k < emails.length; k++) {
                    if (cell.innerHTML.indexOf(emails[k]) > -1) {
                        tr[i].style.display = "";
                        break;
                    }
                }
            }

        }
    }
}

var _Item;

var emails = new Array();

Office.initialize = function () {
    _Item = Office.context.mailbox.item;

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.

        getAllRecipients();

        sortEmails();


    });

}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    //document.write(JSON.stringify(item));
    var toRecipients, ccRecipients, bccRecipients, sender2;
    // Verify if the composed item is an appointment or message.


    if (_Item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = _Item.to;
        ccRecipients = _Item.cc;
        bccRecipients = _Item.bcc;
        sender2 = _Item.sender;

        //document.write(JSON.stringify(sender2));



    }


    displayAddresses(toRecipients);
    if (ccRecipients) { displayAddresses(ccRecipients) };
    if (bccRecipients) { displayAddresses(bccRecipients) };
    displaySender(sender2);


}
function displaySender(asyncResult) {

    write(asyncResult.emailAddress);

}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses(asyncResult) {

    for (var i = 0; i < asyncResult.length; i++) {
        //document.write(JSON.stringify(asyncResult[i].emailAddress));
        write(asyncResult[i].emailAddress);
    }
}

// push to 'emails' array
function write(message) {

    emails.push(message);

}