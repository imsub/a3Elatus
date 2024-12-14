const express = require('express');
const app = express();
const { Client, GatewayIntentBits ,Collection, REST,Routes, SlashCommandBuilder ,IntentsBitField } = require('discord.js');
const myIntents = new IntentsBitField();
myIntents.add(IntentsBitField.Flags.GuildPresences, IntentsBitField.Flags.GuildMembers);
const dotenv = require('dotenv');
const moongoose = require("mongoose");
const voice = require("@discordjs/voice");
const {Attendance} = require("./attendance.modal.js");
const {spreadSheet} = require("./googleSheet.modal.js");
const {DB,TOKEN,CLIENT_ID,client_email,private_key,PORT} = process?.env?.ENV === "production" ?  process.env : dotenv.config().parsed;
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
const MUTEDMEMBER = new Set();
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
        GatewayIntentBits.GuildMembers,
        GatewayIntentBits.GuildVoiceStates,
    ]    
});
client.commands = new Collection();

const deleteCommand = new SlashCommandBuilder().setName('delete').setDescription('admin/moderator can delete record from tracker.')
                    .addStringOption(option => option.setName('displayname').setDescription('enter name shown in discord channel').setRequired(true));

const muteCommand = new SlashCommandBuilder().setName('mute').setDescription('admin/moderator can mute any fucker present in server.')
                    .addStringOption(option => option.setName('rolename').setDescription('enter player role'))
                    .addStringOption(option =>option.setName('playername').setDescription('Enter Player Name'))
                    .addIntegerOption(option =>option.setName('duration').setDescription('enter time in minutes'));

const unmuteCommand = new SlashCommandBuilder().setName('unmute').setDescription('admin/moderator can unmute any fucker present in server.')
                    .addStringOption(option => option.setName('rolename').setDescription('enter player role'))
                    .addStringOption(option =>option.setName('playername').setDescription('Enter Player Name'))
                    .addIntegerOption(option =>option.setName('duration').setDescription('enter time in minutes'));
const commands = [
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
    deleteCommand,
    muteCommand,
    unmuteCommand
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
    try{
        await interaction.deferReply();
        const moderatorRoleId = await getRoleIdBasedonRole(interaction,["Moderator","moderator","Mod","mod","moderators","Moderators","@moderator","Owner","owner","@Moderator"]);
        const isModerator = checkUserRole(interaction,moderatorRoleId.id);
        switch(interaction.commandName){
            case "present":
            case "yes": 
            let username = '',globalName='',id='',nickName='',guildRole='',fetchRole='',roleId='';
                if(interaction.guild){
                    const interactionUser = await interaction.guild.members.fetch(interaction.user.id);;
                    nickName = interactionUser.nickname;
                    username = interactionUser.user.username;
                    globalName = interactionUser.user.globalName;
                    id = interactionUser.user.id;
                    guildRole = await interaction.guild.fetch();
                    fetchRole = await guildRole.roles.fetch();
                    roleId = await fetchRole.find(r => r.name === "Moderator");
                }
                const dateTime = new Date().toLocaleString("en-US", {timeZone: 'Asia/Kolkata'});
                const date = dateTime.split(',')[0];
                const time = dateTime.split(',')[1];
                const checkRecord = await Attendance.findOne({username});
                if(!checkRecord || checkRecord.date !== date){ // checkRecord is null for new user
                    const query = {  userId : id};
                    const update = { $set: { username , globalName , userId : id , nickName,  date ,  attendance : "present" , time}};
                    const options = { upsert: true };
                    await Attendance.updateOne(query, update, options);
                    await interaction.followUp(`Hello, ${globalName} your attendance is captured in our records.`);
                    const data = await Attendance.find( { date: date } );
                    const url = await main(data,date);
                }else{
                    await interaction.followUp(`Hello, ${globalName} your attendance is already captured in our records, please try again tomorrow.`);
                }
                break;
            case "delete":
                    if(isModerator && !!interaction.options.getString("displayname")){
                        const regex = new RegExp(["^", interaction.options.getString("displayname"), "$"].join(""), "i");
                        await Attendance.deleteOne( { globalName: regex } );
                        await interaction.followUp("Attendance deleted.");
                        const dateTime = new Date().toLocaleString("en-US", {timeZone: 'Asia/Kolkata'});
                        const date = dateTime.split(',')[0];
                        const data = await Attendance.find( { date: date } );
                        const url = await main(data,date);
                        await interaction.followUp("Attendance deleted.");
                    }
                    else{
                        await interaction.followUp("You are not authorized to delete records from database.");
                    }
                break;
            case "all":
                    if(isModerator){
                        const dateTime = new Date().toLocaleString("en-US", {timeZone: 'Asia/Kolkata'});
                        const date = dateTime.split(',')[0];
                        const data = await Attendance.find( { date: date } );
                        const url = await main(data,date);
                        await interaction.followUp(url);
                    }
                    else{
                        await interaction.followUp("You are not authorized to fetch attendance sheet.");
                    }
                break;
            case "mute":
                    if(isModerator){
                        const payload = getParams(interaction);
                        if(payload.playername)
                            await muteMemberByName(interaction,payload.playername,payload.duration);
                        else if(payload.rolename)
                            await muteRoleMembers(interaction,payload.rolename,payload.duration);
                        else
                            await interaction.followUp("Please provide either role or player name.");
                    }
                    else{
                        await interaction.followUp("You are not authorized to Mute player.");
                    }
                break;
            case "unmute":
                    if(isModerator){
                        const payload = getParams(interaction);
                        if(payload.playername)
                            await unmuteMemberByName(interaction,payload.playername,payload.duration);
                        else if(payload.rolename)
                            await unmuteRoleMembers(interaction,payload.rolename,payload.duration);
                        else
                            await interaction.followUp("Please provide either role or player name.");
                    }
                    else{
                        await interaction.followUp("You are not authorized to Unmute player.");
                    }
                break;
            default : 
                await interaction.followUp("invalid command. Please try again!");
        }
        if (!interaction.isCommand()) {
            return false;
        }
    }catch(error){
        await interaction.followUp(error.message);
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
    const resource1 = {
        values : data,
    };
    const auth = new google.auth.GoogleAuth({
        keyFile: serviceAccountKeyFile,
        scopes: SCOPES,
    });
    const service = google.sheets({version: 'v4', auth});
    const existingData =  await googleSheetClient.spreadsheets.values.get({
        spreadsheetId: doc.spreadsheetId,
    range: `${tabName}!${range}`,
    });
    if(existingData?.data?.values?.length > 0){
        for(let i=1 ;i<resource1.values.length;i++){ 
            for(let j=1;j<existingData?.data.values.length;j++){
                if(resource1.values[i][0] === existingData.data.values[j][0] && resource1.values[i][2] === existingData.data.values[j][2]){
                    resource1.values[i][4] = existingData.data.values[j][4]
                }
            }
        }
    }
    await googleSheetClient.spreadsheets.values.clear({
        spreadsheetId: doc.spreadsheetId,
    range: `${tabName}!${range}`,
    })
    const result = await service.spreadsheets.values.update({
        spreadsheetId : doc.spreadsheetId,
        range: `${tabName}!${range}`,
        valueInputOption: 'USER_ENTERED',
        resource: resource1,
    });
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
    return doc;
    } catch (err) {
        throw err;
    }
}
async function main(data,date) {
    const googleSheetClient = await _getGoogleSheetClient();
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
function getParams(interaction){
    return {
        rolename : interaction.options.getString("rolename"),
        duration : interaction.options.getInteger("duration"),
        playername : interaction.options.getString("playername")
    }
}
async function muteMemberByName(interaction,name) {
    try {
        const channels = await interaction.guild.channels.fetch();
        const voiceChannelCollection =  channels.filter(c => c.type === 2 || c.type === 'voice');
        let count = 0,memberFound=false;
        for (const [id, voiceChannel] of voiceChannelCollection) {
            count += voiceChannel.members.size;
            for(const [id1,member] of voiceChannel.members) {
                if(member.user.username.toLowerCase() === name.toLowerCase() ||
                (member.user.nickname && member.user.nickname.toLowerCase() === name.toLowerCase()) ||
                (member.user.globalName  && member.user.globalName.toLowerCase().includes(name.toLowerCase()))) {
                    memberFound = true;
                    if(member.voice.mute === false || !MUTEDMEMBER.has(name.toLowerCase())){
                        await member.voice.setMute(true);
                        await interaction.followUp(`Muted ${name} successfully.`);
                        MUTEDMEMBER.add(name.toLowerCase());
                    }
                    else{
                        await interaction.followUp(`${name} is already muted in a voice channel.`);
                    }
                    break;
                }
            }
            if(memberFound) break;
        }
        if(!memberFound){
            await interaction.followUp(`Either ${name} is not in a voice channel or ${name} does not exist in server`);
            return;
        }      
    }catch (error) {
        throw error;
    }
}
async function unmuteMemberByName(interaction, name) {
    try {
        const channels = await interaction.guild.channels.fetch();
        const voiceChannelCollection =  channels.filter(c => c.type === 2 || c.type === 'voice');
        let count = 0,memberFound=false;
        for (const [id, voiceChannel] of voiceChannelCollection) {
            count += voiceChannel.members.size;
            for(const [id1,member] of voiceChannel.members) {
                if(member.user.username.toLowerCase() === name.toLowerCase() ||
                (member.user.nickname && member.user.nickname.toLowerCase() === name.toLowerCase()) ||
                (member.user.globalName  && member.user.globalName.toLowerCase().includes(name.toLowerCase()))) {
                    memberFound = true;
                    if(member.voice.mute === true || MUTEDMEMBER.has(name.toLowerCase())){
                        await member.voice.setMute(false);
                        await interaction.followUp(`Unmuted ${name} successfully.`);
                        MUTEDMEMBER.delete(name.toLowerCase());
                    }
                    else{
                        await interaction.followUp(`${name} is already unmuted in a voice channel.`);
                    }
                    break;
                }
            }
            if(memberFound) break;
        }
        if(!memberFound){
            await interaction.followUp(`Either ${name} is not in a voice channel or ${name} does not exist in server`);
            return;
        }      
    }catch (error) {
        throw error;
    }
}
async function muteRoleMembers(message, roleName) {
    const role = message.guild.roles.cache.find(r => r.name.toLowerCase() === roleName.toLowerCase());

    if (!role) {
        await message.followUp(`Role "${roleName}" not found.`);
        return;
    }
    const members = role.members;

    if (members.size === 0) {
        await message.followUp(`No members found with the role "${roleName}".`);
        return;
    }

    let successCount = 0;
    let failureCount = 0;

    // Iterate through the members
    for (const [memberId, member] of members) {
        try {
            if (member.voice.channel) { // Check if the member is in a voice channel
                await member.voice.setMute(true);
                successCount++;
            } else {
                failureCount++;
            }
        } catch (error) {
            console.error(`Failed to mute ${member.user.tag}:`, error);
            failureCount++;
        }
    }

    await message.followUp(`Muted ${successCount} members with the role "${roleName}". ${failureCount} members were not muted (not in voice channels or other issues).`);
}
async function unmuteRoleMembers(message, roleName) {
    const role = message.guild.roles.cache.find(r => r.name.toLowerCase() === roleName.toLowerCase());
    if (!role) {
        await message.followUp(`Role "${roleName}" not found.`);
        return;
    }

    const members = role.members;

    if (members.size === 0) {
        await message.followUp(`No members found with the role "${roleName}".`);
        return;
    }

    let successCount = 0;
    let failureCount = 0;

    for (const [memberId, member] of members) {
        try {
            if (member.voice.channel) {
                await member.voice.setMute(false);
                successCount++;
            } else {
                failureCount++;
            }
        } catch (error) {
            console.error(`Failed to unmute ${member.user.tag}:`, error);
            failureCount++;
        }
    }

    await message.followUp(`Unmuted ${successCount} members with the role "${roleName}". ${failureCount} members were not unmuted.`);
}
client.on('ready', () => {
    console.log(`Logged in as ${client.user.tag}`);
});
client.on('error',(err)=>{
    console.log(`Error Log ${err.message}`);
})
