
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
        //session.userData.profile = {}; // uncomment to reset userData
        var isChannelConversation = session.message.address.conversation.isGroup;
        
        if (isChannelConversation) {
            // set the user session and kick off the process 

            setusers(session) 
                .then(setchannel(session))
                .then(function(good) {
                    var btns = [];
                    session.userData.profile.users.forEach(function(element) {
                        if (session.message.address.user.id==element.id) {
                            // matched the requesting user
                            session.userData.profile.requser = element;
                        }
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
                });
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
                    .title("Kudos for "+session.userData.profile.selecteduser.name+" working on '"+session.userData.profile.channel.name+"'")
                    .isMultiSelect(true)
                    .style('expanded')
                    .choices([
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Was a great member of the team").value("1"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Kept the team organized and on track").value("2"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Really understood and communicated the customer requirements").value("3"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Knew or learned to use the tools and technology needed for the hack").value("4"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Would definately invite to another hack").value("5"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("I really learned a lot from "+session.userData.profile.selecteduser.name).value("6"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Went above and beyond the call of duty").value("7"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("A technical genius").value("8"),
                    new teams.O365ConnectorCardMultichoiceInputChoice(session).display("Essential to the success of the project").value("9")
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
                    toName: session.userData.profile.selecteduser.name,
                    toEmail: session.userData.profile.selecteduser.email,
                    toId: session.userData.profile.selecteduser.id,
                    toAlias: session.userData.profile.selecteduser.userPrincipalName,
                    frmName: session.userData.profile.requser.name,
                    frmEmail: session.userData.profile.requser.email,
                    frmId: session.userData.profile.requser.id,
                    frmAlias: session.userData.profile.requser.userPrincipalName,
                    projectId: session.userData.profile.channel.id,
                    projectName: session.userData.profile.channel.name
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

                session.userData.profile = {};  // reset userdata for future kudos loops
        }
    }
]);



function setusers (session) {
    return new Promise((good, bad)=>{
        if (!session.userData.profile) session.userData.profile = {};
        if (!session.userData.profile.users) {
            connector.fetchMembers(session.message.address.serviceUrl, session.message.address.conversation.id, function (err, members) {
                if (err) {
                    bad(err);
                    return;
                }
                else {
                    // load the userlist into the profile as it wasn't already there
                    session.userData.profile.users = members;

                    good(members);
                    return;
                }
            });
        } 
        else {
            good(session.userData.profile.users);
            return;
        }
    });
}

function setchannel (session) {
    return new Promise((good, bad)=>{
        if (!session.userData.profile) session.userData.profile = {};
        if (!session.userData.profile.channel) {
            connector.fetchChannelList(session.message.address.serviceUrl, session.message.sourceEvent.team.id, function (err, channels) {
                if (err) {
                    bad(err);
                    return;
                }
                else {
                    channels.forEach(function(channel) {
                        if (channel.id == session.message.sourceEvent.channel.id) {
                            session.userData.profile.channel = channel;
                            good(session.userData.profile.channel);
                            return;
                        }
                    }, this);
                }
            });
        } 
        else {
            good(session.userData.profile.channel);
            return;
        }
    });
}


var o365CardActionHandler = function (event, query, callback) {

    var body = JSON.parse(query.body);
    var kudosList = body.kudoslist.split("; ");

    var docObj = {
        "fromName": body.frmName,
        "fromEmail": body.frmEmail,
        "fromId": body.frmId,
        "fromAlias": body.frmAlias,
        "toName": body.toName,
        "toEmail": body.toEmail,
        "toId": body.toId,
        "toAlias": body.toAlias,
        "projectId": body.projectId,
        "projectName": body.projectName,
        "Q1": false,
        "Q2": false,
        "Q3": false,
        "Q4": false,
        "Q5": false,
        "Q6": false,
        "Q7": false,
        "Q8": false,
        "Q9": false,
        "verbatim": body.textkudos
    }

    kudosList.forEach(function(element) {
        switch (element) {
            case "1":
            docObj.Q1=true;
            break;
            case "2":
            docObj.Q2=true;
            break;
            case "3":
            docObj.Q3=true;
            break;
            case "4":
            docObj.Q4=true;
            break;
            case "5":
            docObj.Q5=true;
            break;
            case "6":
            docObj.Q6=true;
            break;
            case "7":
            docObj.Q7=true;
            break;
            case "8":
            docObj.Q8=true;
            break;
            case "9":
            docObj.Q9=true;
            break;
        }
    }, this);

    client.queryDocuments(
        collectionUrl,
        'SELECT * FROM c where c.fromId = "'+docObj.fromId+'" and c.toId ="'+docObj.toId+'" and c.projectId = "'+docObj.projectId+'"'
        ).toArray((err, results) => {
            if (err) reject(err)
            else {
                if (results.length==0){
                    // there isn't a document for this person review on this project so create it

                    client.createDocument(collectionUrl, docObj, (err, created) => {
                        if (err) reject(err)
                    });
                }
                else {
                    for (var queryResult of results) {
                        replaceDocument(queryResult, docObj);
                    }
                }

            }
        });

    callback(null, null, 200);
};
connector.onO365ConnectorCardAction(o365CardActionHandler);
   

bot.dialog('personButtonClick', [
    function (session, args, next) {
        // Get color and optional size from users utterance
        var utterance = args.intent.matched[0].replace("Give ", "").replace(" <at>Kudos", "");
        if (session.message.address.user.name.replace(" (Guest)","")==utterance) {
            // you can't give kudos to yourself
            session.endDialog("Nice try, you can't give kudos to yourself");
        }
        else {

            setusers(session) 
            .then(setchannel(session))
            .then(function(good) {
                session.userData.profile.users.forEach(function(element) {
                    if(utterance==element.name) {
                        // match the user
                        session.userData.profile.selecteduser = element;
                    }
                    else if (session.message.address.user.id==element.id) {
                        // matched the requesting user
                        session.userData.profile.requser = element;
                    }

                }, this);
                
                session.beginDialog('Start1to1Chat');
            });
        }
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


function replaceDocument(document,docObj) {
    let documentUrl = `${collectionUrl}/docs/${document.id}`;

    docObj.id = document.id;
    return new Promise((resolve, reject) => {
        client.replaceDocument(documentUrl, docObj, (err, result) => {
            if (err) reject(err);
        });
    });
};