'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            sendGTFO(Office.context.mailbox.item);
        });
    });

    // a simple way to escape string in JS, mustache.js 
    // https://stackoverflow.com/questions/24816/escaping-html-strings-with-jquery
    function escapeHtml(string) {

        var entityMap = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;',
            '/': '&#x2F;',
            '`': '&#x60;',
            '=': '&#x3D;'
        };

        return String(string).replace(/[&<>"'`=\/]/g, function (s) {
            return entityMap[s];
        });
    }

    function sendGTFO(item) {

        // fetching message template
        var template = $.ajax({
            url: "ndr.template.txt",
            async: false
        }).responseText;        

        // exploding recipient email
        var recipient_email = Office.context.mailbox.userProfile.emailAddress;
        var recipient_alias = recipient_email.split("@")[0];
        var recipient_domain = recipient_email.split("@")[1];

        // preparing messsageId
        var message_id = item.internetMessageId.replace('<', '&#60;');
        message_id = message_id.replace('>', '&#62;');

        // replacing tokens in the template with actual values
        template = template.replace(/%recipient_email%/g, recipient_email);
        template = template.replace(/%recipient_alias%/g, recipient_alias);
        template = template.replace(/%recipient_domain%/g, recipient_domain);
        template = template.replace(/%sender_name%/g, item.from.displayName);
        template = template.replace(/%sender_email%/g, item.from.emailAddress);
        template = template.replace(/%subject%/g, item.subject);
        template = template.replace(/%message_id%/g, message_id);
        template = template.replace(/%email_timestamp%/g, item.dateTimeCreated);

        // we need to wrap this into EWS XML, so everything needs to be escaped
        template = escapeHtml(template);

        // EWS XML stub, https://stackoverflow.com/questions/57192921/is-there-a-way-to-send-a-mail-seamlessly-using-office-js
        var request = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
            '  <soap:Header><t:RequestServerVersion Version="Exchange2010" /></soap:Header>' +
            '  <soap:Body>' +
            '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
            '      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>' +
            '      <m:Items>' +
            '        <t:Message>' +
            '          <t:Subject>Undeliverable: ' + item.subject + '</t:Subject>' +
            '          <t:Body BodyType="HTML">' + template + '</t:Body>' +
            '          <t:ToRecipients>' +
            '            <t:Mailbox><t:EmailAddress>' + item.from.emailAddress + '</t:EmailAddress></t:Mailbox>' +
            '          </t:ToRecipients>' +
            '        </t:Message>' +
            '      </m:Items>' +
            '    </m:CreateItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        // submitting the message and displaying the result
        Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
            if (asyncResult.status == "failed") {
                $("#StatusResult").text("Send action failed with error: " + asyncResult.error.message);
            }
            else {
                $("#StatusResult").text("Fake NDR sent.");
                deleteMessage(item);
            }
        });

    }

    function deleteMessage(item) {

        // looks like deleteItem is not available for add-ins, we'll move the message instead
        var deleteRequest = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
            '<soap:Header><t:RequestServerVersion Version="Exchange2013" /></soap:Header>' +
            '<soap:Body><m:MoveItem><m:ToFolderId><t:DistinguishedFolderId Id="deleteditems" /></m:ToFolderId>' +
            '<m:ItemIds><t:ItemId Id="' + item.itemId + '" /></m:ItemIds></m:MoveItem></soap:Body></soap:Envelope>';

        // deleting the message, updating the panel
        // but if everything goes well, user won't see it
        Office.context.mailbox.makeEwsRequestAsync(deleteRequest, function (deleteAsyncResult) {
            if (deleteAsyncResult.status == "failed") {
                $("#StatusResult").text("<br /><br />Delete action failed with error: " + deleteAsyncResult.error.message);
            }
            else {
                $("#StatusResult").append("<br /><br />Message moved to Trash.");
            }
        });

    }

})();