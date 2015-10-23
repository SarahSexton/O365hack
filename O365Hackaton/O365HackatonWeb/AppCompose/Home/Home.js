/// <reference path="../App.js" />

(function () {
    'use strict';

    var item;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {

        item = Office.context.mailbox.item;

        //var subject = item.subject.getAsync([test]);        

        //$("#lblStatus").text(subject);

        $(document).ready(function () {
            app.initialize();

            $('#set-subject').click(setSubject);
            $('#get-subject').click(getSubject);
            $('#add-to-recipients').click(addToRecipients);            
           
            $(document).ready(function () {
                // After the DOM is loaded, app-specific code can run.
                // Get all the recipients of the composed item.
                getAllRecipients();

                getSubject();
            });
        });
    };

  

    function setSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync("Hello world!");
    }

    function getSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function (result) {
            app.showNotification('The current subject is', result.value)

            // got the name of the event
            //$("#lblStatus").text(result.value);

            // update the link
            var newUrl = 'http://turnout.azurewebsites.net/User?eventid= ' + result.value;
            $('#lnkGoTo').attr('href',newUrl);

        });
    }

    function addToRecipients() {
        var item = Office.context.mailbox.item;
        var addressToAdd = {
            displayName: Office.context.mailbox.userProfile.displayName,
            emailAddress: Office.context.mailbox.userProfile.emailAddress
        };
 
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
        }
    }
     
   
  

    //Office.initialize = function () {
    //    item = Office.context.mailbox.item;
    //    // Checks for the DOM to load using the jQuery ready function.
    //    $(document).ready(function () {
    //        // After the DOM is loaded, app-specific code can run.
    //        // Get all the recipients of the composed item.
    //        getAllRecipients();
    //    });
    //}

    // Get the email addresses of all the recipients of the composed item.
    function getAllRecipients() {
        // Local objects to point to recipients of either
        // the appointment or message that is being composed.
        // bccRecipients applies to only messages, not appointments.
        //var toRecipients, ccRecipients, bccRecipients;
        // Verify if the composed item is an appointment or message.
        //if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
            //toRecipients = item.requiredAttendees;
            //ccRecipients = item.optionalAttendees;

           // $("#lblStatus").text(item.itemId);

        //  $("#lblStatus").text(item.itemId + '   test');


        //}
        //else {
        //    //toRecipients = item.to;
        //    //ccRecipients = item.cc;
        //    //bccRecipients = item.bcc;
        //}

       

      
        }
    
    // Recipients are in an array of EmailAddressDetails
    // objects passed in asyncResult.value.
    function displayAddresses(asyncResult) {
        for (var i = 0; i < asyncResult.value.length; i++)
            write(asyncResult.value[i].emailAddress);
    }

    // Writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }

    

   
   
    

})();