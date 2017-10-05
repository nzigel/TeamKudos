
var restify = require('restify');
const builder = require('botbuilder');
const teams = require("botbuilder-teams");
require('dotenv-extended').load();
var documentClient = require('documentdb').DocumentClient;

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});

var config = {}
config.endpoint = process.env.DOCUMENTDB_HOST;
config.primaryKey = process.env.DOCUMENTDB_KEY;

config.database = {
    "id": process.env.DOCUMENTDB_DATABASE
};

config.collection = {
    "id": process.env.DOCUMENTDB_COLLECTION
};

var client = new documentClient(config.endpoint, { "masterKey": config.primaryKey });
var databaseUrl = `dbs/${config.database.id}`;
var collectionUrl = `${databaseUrl}/colls/${config.collection.id}`;

var appName = process.env.MICROSOFT_APP_NAME;
var appId = process.env.MICROSOFT_APP_ID;

// Create chat connector for communicating with the Bot Framework Service
var connector = new teams.TeamsChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

var bot = new builder.UniversalBot(connector, [
    function (session) {
        //session.userData.profile = {};
        var isChannelConversation = session.message.address.conversation.isGroup;
        
        if (isChannelConversation) {
            session.beginDialog('setUserList', session.userData.profile);
        }
        else {
            // we have a 1:1 conversation with a selected user
            //session.endDialog('%s', "Ready to provide Kudos for "+session.userData.profile.selecteduser.name);

            var giveKudosCard = new teams.O365ConnectorCardActionCard(session)
            .id("cardKudos")
            .name("giveKudos")
            .inputs([
                new teams.O365ConnectorCardMultichoiceInput(session)
                    .id("kudoslist")
                    .title("Kudos for "+session.userData.profile.selecteduser.name)
                    .isMultiSelect(true)
                    .style('expanded')
                    .choices([
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Was a great member of the team").value("0"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Kept the team organized and on track").value("1"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Really understood and communicated the customer requirements").value("2"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Knew or learned to use the tools and technology needed for the hack").value("3"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Would definately invite to another hack").value("4"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("I really learned a lot from "+session.userData.profile.selecteduser.name).value("5"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Went above and beyond the call of duty").value("6"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("A technical genius").value("7"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Essential to the success of the project").value("8")
                    ]),
                new teams.O365ConnectorCardTextInput(session)
                        .id("textkudos")
                        .title("In your own words:")
                        .isMultiline(true)
            ])
            .actions([
                new teams.O365ConnectorCardHttpPOST(session)
                    .id("cardKudosBtn")
                    .name("Give Kudos")
                    .body(JSON.stringify({
                    kudoslist: '{{kudoslist.value}}',
                    textkudos: '{{textkudos.value}}',
                    }))
            ]);

            var card = new teams.O365ConnectorCard(session)
                    .themeColor("#E67A9E")
                    .title("Give Kudos to "+session.userData.profile.selecteduser.name+":")
                    .potentialAction([
                    giveKudosCard
                    ]);
            
                var msg = new teams.TeamsMessage(session)
                    .attachments([card]);
            
                session.send(msg);
                session.endDialog();

        }
    },
    function (session, results) {
        session.userData.profile.users = results.response; // Save user profile.
        
        var btns = [];
        session.userData.profile.users.forEach(function(element) {
            btns.push(builder.CardAction.imBack(session, "Give "+element.name+" Kudos", element.name));
        }, this);

        var msg = new builder.Message(session);
        msg.attachmentLayout(builder.AttachmentLayout.carousel)
        msg.attachments([
            new builder.HeroCard(session)
                .title('Select a person to give Kudos')
                .buttons(btns)
        ]);
        session.send(msg).endDialog();
        
    }
]);

var o365CardActionHandler = function (event, query, callback) {
    var userId = event.address.user.id;
    var body = JSON.parse(query.body);
    var kudosList = body.kudoslist;
    var kudosText = body.textkudos;
    
    callback(null, null, 200);
};
connector.onO365ConnectorCardAction(o365CardActionHandler);

bot.dialog('setUserList', [
    function (session, args, next) {
        session.dialogData.profile = args || {}; // Set the profile or create the object.
        if (!session.dialogData.profile.users) {
            connector.fetchMembers(session.message.address.serviceUrl, session.message.address.conversation.id, function (err, result) {
                if (err) {
                    session.endDialog('There was an error collecting the list of users.');
                }
                else {
                    // load the userlist into the profile as it wasn't already there
                    session.dialogData.profile.users = result;
                    next();
                }
            });
        } 
        else {
            next();
        }
        
    },
    function (session) {
        session.endDialogWithResult({ response: session.dialogData.profile.users });
    }
]);     

bot.dialog('personButtonClick', [
    function (session, args, next) {
        // Get color and optional size from users utterance
        var utterance = args.intent.matched[0].replace("Give ", "").replace(" <at>Kudos", "");
        session.userData.profile.users.forEach(function(element) {
            if(utterance==element.name) {
                // match the user
                session.userData.profile.selecteduser = element;
            }
        }, this);
        session.beginDialog('Start1to1Chat');
    }
]).triggerAction({ matches: /(Give)\s.*Kudos/i });


bot.dialog('Start1to1Chat', function (session) {
    var userId = session.message.address.user.id;
    var tenantId = teams.TeamsMessage.getTenantId(session.message);

    var address = {
        channelId: 'msteams',
        user: { id: userId },
        channelData: {
            tenant: {
                id: tenantId
            }
        },
        bot: {
            id: appId,
            name: appName
        },
        serviceUrl: session.message.address.serviceUrl,
        useAuth: true
    };

    bot.beginDialog(address, '/');

});
