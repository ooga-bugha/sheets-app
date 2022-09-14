const {google} = require('googleapis');
const externalObjs = require('./sheetObject');
const { workerData, parentPort } = require('worker_threads');
const unirest = require('unirest');

// Initialize sheet object

let testSheet = new externalObjs.Sheets(workerData.sheet_ID,workerData.sheet_email,workerData.sheet_private_key);

// Initialize gsapi client

const client = new google.auth.JWT(
    testSheet.creds.client_email, null, testSheet.creds.private_key,['https://www.googleapis.com/auth/spreadsheets']
);

console.log("Worker Started");

// Authorize Client

client.authorize(async function(err, tokens){
    if(err){
        console.log(err);
        return;
    }
    else{

        // Get header Row of sheet

        await testSheet.getPrimarySheet(client);
        await testSheet.getHeaderRow(client);

        // Perform Bulk Close
        
        if(workerData.task == "bulk_close"){
            await testSheet.bulkClose(client, workerData.ticket_id_column, workerData.response_code, workerData.error_message, workerData.api_key, workerData.fd_url).catch(err =>{
                console.log(err);
                
                // Send email on failure

                unirest.post(`https://smartstation.freshdesk.com/api/v2/tickets/outbound_email`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ZER3dXNTb2hOa1QzRm5yMHg0eg==`
                }).send({
                    "email":"nikhiloogabugha@gmail.com",
                    "subject":`${ workerData.task } job failed for ${ workerData.fd_url }!!`,
                    "description":`Task failed error details below \n ${ err }`,
                    "email_config_id":82000032722,
                    "priority":2,
                    "status":2
                })
            });

            // Send email on success

            await unirest.post(`https://smartstation.freshdesk.com/api/v2/tickets/outbound_email`).headers({
                'Content-Type': 'application/json',
                'Authorization':`Basic ZER3dXNTb2hOa1QzRm5yMHg0eg==`
            }).send({
                "email":"nikhiloogabugha@gmail.com",
                "subject":`${ workerData.task } job successful for ${ workerData.fd_url }!!`,
                "description":`Task successful!`,
                "email_config_id":82000032722,
                "priority":2,
                "status":2
            });
        }
        
        // Perform Bulk Reopen

        if(workerData.task == "bulk_reopen"){
            await testSheet.bulkReopen(client, workerData.ticket_id_column, workerData.response_code, workerData.error_message, workerData.api_key, workerData.fd_url).catch(err =>{
                console.log(err);

                // Send email on failure

                unirest.post(`https://smartstation.freshdesk.com/api/v2/tickets/outbound_email`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ZER3dXNTb2hOa1QzRm5yMHg0eg==`
                }).send({
                    "email":"nikhiloogabugha@gmail.com",
                    "subject":`${ workerData.task } job failed for ${ workerData.fd_url }!!`,
                    "description":`Task failed error details below \n ${ err }`,
                    "email_config_id":82000032722,
                    "priority":2,
                    "status":2
                })
            });

            // Send email on success
            
            await unirest.post(`https://smartstation.freshdesk.com/api/v2/tickets/outbound_email`).headers({
                'Content-Type': 'application/json',
                'Authorization':`Basic ZER3dXNTb2hOa1QzRm5yMHg0eg==`
            }).send({
                "email":"nikhiloogabugha@gmail.com",
                "subject":`${ workerData.task } job successful for ${ workerData.fd_url }!!`,
                "description":`Task successful!`,
                "email_config_id":82000032722,
                "priority":2,
                "status":2
            });
        }
    }
});