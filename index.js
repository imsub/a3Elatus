const express = require('express');
const app = express();
const { Client, GatewayIntentBits ,Collection, REST,Routes, SlashCommandBuilder ,IntentsBitField } = require('discord.js');
const myIntents = new IntentsBitField();
myIntents.add(IntentsBitField.Flags.GuildPresences, IntentsBitField.Flags.GuildMembers);
const dotenv = require('dotenv');
const moongoose = require("mongoose");
const {Attendance} = require("./attendance.modal.js");
const {spreadSheet} = require("./googleSheet.modal.js");
const {DB,TOKEN,CLIENT_ID,client_email,private_key,GOOGLE_API_KEY,PORT} = process?.env?.ENV === "production" ?  process.env : dotenv.config().parsed;
const { google } = require('googleapis');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const serviceAccountKeyFile = "./googleSheetAPI.json";
const { JWT } = require('google-auth-library');
const sheetId = '';
const tabName = 'Sheet1'
const range = 'A:E';
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/drive',
];
const jwt = new JWT({
    email: client_email,
    key: private_key.replace(/\\n/g, "\n"),
    scopes: SCOPES,
  });
const client = new Client({ 
    intents: [
        GatewayIntentBits.Guilds , 
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent,
    ]    
});
client.commands = new Collection();

// client.on("message",(message )=>{
//     const roleName = message.member.roles.cache.find(r => r.name === "Owner")
//             if (roleName) {
//                 return message.reply("You can use this command.")
//             } else {
//                 return message.reply("Sorry, an error occured.")
//             }
// });
const data = new SlashCommandBuilder()
	.setName('delete')
	.setDescription('admin/moderator can delete record from tracker.')
	.addStringOption(option =>
		option.setName('displayname')
		.setDescription('enter name shown in discord channel').setRequired(true))
	    .addBooleanOption(option =>
		option.setName('username')
		.setDescription('enter userid/username of user'));
const commands = [
    {
        name : "lock",
        description : "Lock Attendance Sheet"
    },
    {
        name : "present",
        description : "record attendance"
    },
    {
        name : "yes",
        description : "record attendance"
    },
    {
        name : "all",
        description : "get attendance sheet"
    },
    data
]

const rest = new REST({version:"10"}).setToken(TOKEN);
(async ()=>{
    try {
        await moongoose.connect(DB,{  serverSelectionTimeoutMS: 5000 });
        console.log("connected to mongo db");
        console.log('Started refreshing application (/) commands.');
      
        await rest.put(Routes.applicationCommands(CLIENT_ID), { body: commands });//command cannot be longer than 15 charecters
      
        console.log('Successfully reloaded application (/) commands.');
        app.listen(PORT, () =>{
            console.log(`App listening at http://localhost:${PORT}`);
        });
      } catch (error) {
        console.error(error);
      }
})();
const getRoleIdBasedonRole = async (interaction,roleName)=>{
    const guildRole = await interaction.guild.fetch();
    const fetchRole = await guildRole.roles.fetch();
    const roleId = await fetchRole.find( r => roleName.includes(r.name));
    return roleId;
}
const checkUserRole = (interaction,roleId)=>{
    return interaction.member.roles.cache.has(roleId);
}
// app.get('/', (request, response) => {
// 	response.send("Connection to endpoint successfull");
// });
client.on('interactionCreate', async interaction => {
    switch(interaction.commandName){
        case "present":
        case "yes": 
        let username = '',globalName='',id='',nickName='',guildRole='',fetchRole='',roleId='',isModerator=false;
        try{
            if(interaction.guild){
                const interactionUser = await interaction.guild.members.fetch(interaction.user.id);
                //const userRole = await interaction.guild.roles.fetch(interaction.user.id);
                //const getrole = await interaction.guild.members.fetch(interaction.user.id);
                nickName = interactionUser.nickname;
                username = interactionUser.user.username;
                globalName = interactionUser.user.globalName;
                id = interactionUser.user.id;
                guildRole = await interaction.guild.fetch();
                fetchRole = await guildRole.roles.fetch();
                roleId = await fetchRole.find(r => r.name === "Moderator");
                isModerator = roleId ? interaction.member.roles.cache.has(roleId.id) : false;
            }
            else{
                username = interaction.member.user.username;
                globalName = interaction.member.user.global_name;
                id = interaction.member.user.id;
                nickName = interaction.member.nick;
                guildRole =  await interaction.options.client.guilds.fetch();
                fetchRole = await guildRole.roles.fetch();
                roleId = await fetchRole.find(r => r.name === "moderator");
            }
            //const interactionUser = await interaction.guild.members.fetch(interaction.user.id);
            //const userRole = await interaction.guild.roles.fetch(interaction.user.id);
            //const getrole = await interaction.guild.members.fetch(interaction.user.id);
             //nickName = interactionUser.nickname;
            //{username  , globalName , id} = interactionUser.user;
            const dateTime = new Date().toLocaleString("en-US", {timeZone: 'Asia/Kolkata'});
            const date = dateTime.split(',')[0];
            const time = dateTime.split(',')[1];
            const checkRecord = await Attendance.findOne({username});
            if(!checkRecord || checkRecord.date !== date){ // checkRecord is null for new user
                const query = {  userId : id};
                const update = { $set: { username , globalName , userId : id , nickName,  date ,  attendance : "present" , time}};
                const options = { upsert: true };
                await Attendance.updateOne(query, update, options);
                await interaction.reply(`Hello, ${globalName} your attendance is captured in our records.`);
                //interaction.channel.send({content : `Hello, ${globalName} your attendance is captured in our records.`});
            }else{
                await interaction.reply(`Hello, ${globalName} your attendance is already captured in our records, please try again tomorrow.`);
                //interaction.channel.send({content : `Hello, ${globalName} your attendance is already captured in our records, please try again tomorrow.`});
            }
        }catch(error){
            //interaction.reply(error.message);
            await interaction.channel.send({content : error.message});
        }
            break;
        case "delete":
            try{
                const moderatorRoleId = await getRoleIdBasedonRole(interaction,["Moderator","moderator","Mod","mod","moderators","Moderators"]);
                const isModerator = checkUserRole(interaction,moderatorRoleId.id);
                if(isModerator && !!interaction.options.getString("displayname")){
                    await Attendance.deleteOne( { globalName: interaction.options.getString("displayname") } );
                    await interaction.reply("Attendance deleted.");
                    //interaction.channel.send({content : "Attendance deleted."});
                }
                else{
                    await interaction.reply("You are not authorized to delete records from database.");
                    //interaction.channel.send({content : "You are not authorized to delete records from database."});
                }
            }catch(error){
                await interaction.reply(error.message);
                //interaction.channel.send({content : error.message});
            }
            break;
        case "all":
            try{
                const moderatorRoleId = await getRoleIdBasedonRole(interaction,["Moderator","moderator","Mod","mod","moderators","Moderators"]);
                const isModerator = checkUserRole(interaction,moderatorRoleId?.id);
                if(isModerator){
                    await interaction.reply("Generating URL for Excel Spread Sheet.");
                    //await interaction.channel.send({content : "Generating URL for Excel Spread Sheet."});
                    const dateTime = new Date().toLocaleString("en-US", {timeZone: 'Asia/Kolkata'});
                    const date = dateTime.split(',')[0];
                    const data = await Attendance.find( { date: date } );
                    const url = await main(data,date);
                    //console.log(`url---> ${url}`);
                    await interaction.channel.send({content : url})
                    // data.forEach(async element => {
                    //     await interaction.channel.send({ content: element.toString()});
                    // });
                }
                else{
                    await interaction.reply("You are not authorized to fetch attendance sheet.");
                    //interaction.channel.send({content : "You are not authorized to fetch attendance sheet."});
                }
            }catch(error){
                await interaction.channel.send({content : error.message})
            }
            break;
        default : 
            await interaction.reply("invalid command. Please try again!");
            //interaction.channel.send({content : "invalid command. Please try again!"});
    }
    // Making sure the interaction is a command
    if (!interaction.isCommand()) {
        return false;
    }
})
client.login(TOKEN);

async function _writeGoogleSheet(googleSheetClient, sheetId, tabName, range, data) {
    let doc ='';
    const dateTime = new Date().toLocaleString("en-US", {timeZone: 'Asia/Kolkata'});
    const date = dateTime.split(',')[0];
    const time = dateTime.split(',')[1];
    const existingSpreadsheet = await spreadSheet.find( { date: date } );
    if(existingSpreadsheet?.length){
        const object = existingSpreadsheet[0].toJSON();
        doc = {spreadsheetId: object.url.split("/")[5] , _spreadsheetUrl:object.url};
    }
    else{
        doc = await create();
    }
    //const permission =  await shareFile(newSpreadsheetId);
      const resource1 = {
          values : data,
      };
      const auth = new google.auth.GoogleAuth({
        keyFile: serviceAccountKeyFile,
        scopes: SCOPES,
      });
      const service = google.sheets({version: 'v4', auth});
      
          await googleSheetClient.spreadsheets.values.clear({
              spreadsheetId: doc.spreadsheetId,
          range: `${tabName}!${range}`,
          })
        //   await googleSheetClient.spreadsheets.values.append({
        //   spreadsheetId: doc.spreadsheetId,
        //   range: `${tabName}!${range}`,
        //   valueInputOption: 'USER_ENTERED',
        //   insertDataOption: 'INSERT_ROWS',
        //   resource: {
        //       "majorDimension": "ROWS",
        //       "values": data
        //   },
        //   });
        const result = await service.spreadsheets.values.update({
            spreadsheetId : doc.spreadsheetId,
            range: `${tabName}!${range}`,
            valueInputOption: 'USER_ENTERED',
            resource: resource1,
          });
        //   console.log('%d cells updated.', result.data.updatedCells);
    const query = {  url : doc._spreadsheetUrl};
    const update = { $set: { url : doc._spreadsheetUrl ,  date , time}};
    const options = { upsert: true };
    await spreadSheet.updateOne(query, update, options);
    return doc._spreadsheetUrl;
  }
  async function create() {
    try {
      const dateTime = new Date().toLocaleString("en-US", {timeZone: 'Asia/Kolkata'});
      const date = dateTime.split(',')[0];
      const doc = await GoogleSpreadsheet.createNewSpreadsheetDocument(jwt, { title: `attendance sheet ${date}` });
      await doc.loadInfo();
      //const sheet1 = doc.sheetsByIndex[0];
      //const spreadsheetGenerated = new GoogleSpreadsheet(doc.spreadsheetId, jwt);
      //const permissions = await doc.listPermissions();
      await doc.setPublicAccessLevel('writer');
    //   console.log(`url ${doc._spreadsheetUrl}`);
    //   console.log(doc._spreadsheetUrl);
      return doc;
    } catch (err) {
        // TODO (developer) - Handle exception
        throw err;
    }
  }
  async function main(data,date) {
    const googleSheetClient = await _getGoogleSheetClient();
    // Reading Google Sheet from a specific range
    //const data = await _readGoogleSheet(googleSheetClient, sheetId, tabName, range);
    //console.log(data);
    // Adding a new row to Google Sheet
    const dataToBeInserted = [  
      ['Global Name' , 'User Name', 'Date', 'Time','Share Given'],
    ];
    data.forEach(element =>{
        const obj = element.toJSON();
        dataToBeInserted.push([obj.globalName,obj.username,obj.date,obj.time,obj.shareGiven]);
    })
    const url =  await _writeGoogleSheet(googleSheetClient, sheetId, tabName, range, dataToBeInserted);
    return url;
  }

  async function _getGoogleSheetClient() {
    const auth = new google.auth.GoogleAuth({
      keyFile: serviceAccountKeyFile,
      scopes: SCOPES,
    });
    const authClient = await auth.getClient();
    return await google.sheets({
      version: 'v4',
      auth: authClient,
    });
  }