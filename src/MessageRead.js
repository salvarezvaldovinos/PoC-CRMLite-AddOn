(function () {
    "use strict";

    var messageBanner;

    // La función initialize de Office debe ejecutarse cada vez que se carga una página nueva.
    //Office.initialize = function (reason) {
    //  $(document).ready(function () {
    //    var element = document.querySelector('.MessageBanner');
    //    messageBanner = new components.MessageBanner(element);
    //    messageBanner.hideBanner();
    //    loadProps();
    //  });
    //};

    Office.initialize = function (reason) {
        $(document).ready(function () {
            var appId = "7d86de0a-b1ac-4bb6-8bdc-2383e359f44e";
            var item = Office.context.mailbox.item;
            var url = "https://apps.powerapps.com/play/e/ec9ac751-4e0a-ecf6-bc33-e35e80f6c8e5/a/7d86de0a-b1ac-4bb6-8bdc-2383e359f44e?tenantId=132f2abd-4df2-4ff5-9bba-ddd8d566411c?source=iframe";
            var parameters =
                "&mailid=" + item.itemId +
                "&from=" + item.from.emailAddress +
                "&fromname=" + item.from.displayName +
                "&subject=" + item.subject +
                "&dateTimeReceived" + item.dateTimeCreated;
            $('#canvas-iframe').attr("src", url + parameters);

            Office.context.mailbox.item.body.getAsync('text', function (result) {
                debugger;
                var data = { id: item.itemId, body: result.value };
                $.ajax({
                    type: 'POST',
                    url: 'https://prod-30.brazilsouth.logic.azure.com:443/workflows/cd538fd158824b1884640ffd66b1fe35/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hA-R47o40awZKLUatNyzXjOzq5a1p9MMzN2nEXTXmyo',
                    contentType: "application/json",
                    data: JSON.stringify(data),
                    success: function (response) {
                        debugger;
                        console.log(response);
                    },
                    error: function (ex) {
                        debugger;
                        console.log(ex);
                    }
                });
            });

        });
    };

    // Tome una matriz de objetos AttachmentDetails y cree una lista de nombres de datos adjuntos separados por saltos de línea.
    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            var returnString = "";

            for (var i = 0; i < attachments.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + attachments[i].name;
            }

            return returnString;
        }

        return "None";
    }

    // Dé formato a un objeto EmailAddressDetails como
    // GivenName Surname <emailaddress>v
    function buildEmailAddressString(address) {
        return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }

    // Tome una matriz de objetos EmailAddressDetails y
    // cree una lista de cadenas con formato separadas por saltos de línea.
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Cargue las propiedades del objeto base Item y cargue las
    // propiedades específicas del mensaje.
    function loadProps() {
        var item = Office.context.mailbox.item;

        $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
        $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
        $('#itemClass').text(item.itemClass);
        $('#itemId').text(item.itemId);
        $('#itemType').text(item.itemType);

        $('#message-props').show();

        $('#attachments').html(buildAttachmentsString(item.attachments));
        $('#cc').html(buildEmailAddressesString(item.cc));
        $('#conversationId').text(item.conversationId);
        $('#from').html(buildEmailAddressString(item.from));
        $('#internetMessageId').text(item.internetMessageId);
        $('#normalizedSubject').text(item.normalizedSubject);
        $('#sender').html(buildEmailAddressString(item.sender));
        $('#subject').text(item.subject);
        $('#to').html(buildEmailAddressesString(item.to));
    }

    // Función del asistente para mostrar notificaciones
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();