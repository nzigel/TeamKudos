
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

// Create chat connector for communicating with the Bot Framework Service
var connector = new teams.TeamsChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

var bot = new builder.UniversalBot(connector, [
    function (session) {
        var conversationId = session.message.address.conversation.id;
        connector.fetchMembers(session.message.address.serviceUrl, conversationId, function (err, result) {
            if (err) {
                session.endDialog('There was an error collecting the list of users.');
            }
            else {
                // list the users in teams
                session.endDialog('%s', JSON.stringify(result));
            }
        });
    }
]);
