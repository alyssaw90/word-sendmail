
/*//////////////////////////////////////

Make following changes -

1. Replace clientid = "a8c2f862-5700-47ae-8d21-aac6681012de" with your own appId like clientid = "yy-yy-yy-yy"
2. Replace following emailId - "tarunc@microsoft.com" with Recipients emailid

///////////////////////////////////////*/


var clientId = "{AppID}";
var authenticator;
var graphToken;

(function () {
    "use strict";
    var messageBanner;

    Office.initialize = function (reason) {
        $(document).ready(function () {

            // STEP 2: This to inform the Authenticator to automatically close the authentication dialog once the authentication is complete.
            if (OfficeHelpers.Authenticator.isAuthDialog()) return;
            
            $('#sendOffer').click(
                sendEmailPrep);
        });
    };


    function sendEmailPrep() {


        // STEP 3: Create a new instance of Authenticator and register the endpoints
        authenticator = new OfficeHelpers.Authenticator();

        // Optional: Delete the cached Token.
        if (authenticator.tokens[OfficeHelpers.DefaultEndpoints.Microsoft])
            authenticator.tokens.remove(OfficeHelpers.DefaultEndpoints.Microsoft);


        authenticator.endpoints.registerMicrosoftAuth(clientId, {
            redirectUrl: 'https://localhost:44300/Home.html'/* , scope: 'User.Read.All Mail.Send'*/
        });

        // STEP 4: To authenticate against the registered endpoint, do the following
        authenticator
            .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
            .then(function (token) {
                if (!token) {
                    console.log("ADAL error occurred: " + error);
                    return;
                }
                else {
                    console.log(token);
                    clientCall(token);
                }
            })
            .catch(function (error) {
                console.log(error);
            });

    }

    function clientCall(token) {

        graphToken = token;
        var client = MicrosoftGraph.Client.init({
            authProvider: function (done) {
                done(null, graphToken.access_token); //first parameter takes an error if you can't get an access token
            }
        });



        const mail = {
            subject: "Email from Add-in",
            toRecipients: [{
                emailAddress: {
                    address: "tarunc@microsoft.com"
                }
            }],
            body: {
                content: "Hello, You've Got an EMail",
                contentType: "html"
            }
        };

        client
            .api('/users/me/sendMail')
            .post({ message: mail }, function (err, res) {
                console.log(res); 

                // Run a batch operation against the Word object model.
                Word.run(function (context) {

                    // Create a proxy object for the paragraphs collection.
                    var paragraphs = context.document.body.paragraphs;

                    // Queue a commmand to load the style property for the top 2 paragraphs.
                    // We never perform an empty load. We always must request a property.
                    context.load(paragraphs, { select: 'style', top: 2 });

                    // Synchronize the document state by executing the queued commands, 
                    // and return a promise to indicate task completion.
                    return context.sync().then(function () {

                        // Queue a command to get the first paragraph.
                        var paragraph = paragraphs.items[0];

                        // Queue a command to insert the paragraph after the current paragraph.
                        paragraph.insertParagraph('Hurray!! You sent an Email using Graph API', Word.InsertLocation.after);

                        // Synchronize the document state by executing the queued commands, 
                        // and return a promise to indicate task completion.
                        return context.sync().then(function () {
                            console.log('Inserted a new paragraph at the end of the first paragraph.');
                        });
                    });
                })
                .catch(function (error) {
                    console.log('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });

            });
    }
})();
