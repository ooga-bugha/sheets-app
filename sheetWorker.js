const {google} = require('googleapis');
const externalObjs = require('./sheetObject');
const { workerData, parentPort } = require('worker_threads');

let testSheet = new externalObjs.Sheets(workerData.sheet_ID,workerData.sheet_email,workerData.sheet_private_key);

const client = new google.auth.JWT(
    testSheet.creds.client_email, null, testSheet.creds.private_key,['https://www.googleapis.com/auth/spreadsheets']
);

console.log("Worker Started");

client.authorize(async function(err, tokens){
    if(err){
        console.log(err);
        return;
    }
    else{
        await testSheet.getPrimarySheet(client);
        await testSheet.getHeaderRow(client);
        if (testSheet.headerRow.includes("Response Status") & testSheet.headerRow.includes("Response Error")){
            console.log("It's there");
        }
        else{
            console.log("Nope, adding header values");
            await testSheet.setHeaderRow(client);
        }
        console.log("Adding values");
        await testSheet.addValues(client);
    }
});