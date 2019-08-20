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

var item;

var emails = new Array();

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
        sortEmails();
        noEmails();

    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    } else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients.

            displayAddresses(asyncResult);
        }
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        } else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.

            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                // Async call to get bcc-recipients of the item completed.
                // Display the email addresses of the bcc-recipients.

                displayAddresses(asyncResult);
            }

        }); // End getAsync for bcc-recipients.
    }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses(asyncResult) {
    for (var i = 0; i < asyncResult.value.length; i++)
        write(asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message) {
    emails.push(message);
}

function noEmails() {
    if (emails.length == 0) {
        document.getElementById("emptyEmails").innerHTML = "Add email addresses to see colors!";
    }
}