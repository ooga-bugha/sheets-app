const { get } = require('express/lib/request');
const {google} = require('googleapis');
const unirest= require('unirest');

// Function to get A1 notation

const getA1Notation = (row, column) => {
    const a1Notation = [`${row + 1}`];
    const totalAlphabets = 'Z'.charCodeAt() - 'A'.charCodeAt() + 1;
    let block = column;
    while (block >= 0) {
      a1Notation.unshift(String.fromCharCode((block % totalAlphabets) + 'A'.charCodeAt()));
      block = Math.floor(block / totalAlphabets) - 1;
    }
    return a1Notation.join('');
};

// Sheets Object

function Sheets(sheetId, client_email, private_key){
    this.sheetId = sheetId;
    this.creds = {'client_email':client_email,'private_key':private_key};
}

// Get existing header row and set property

Sheets.prototype.getHeaderRow = async function(cl){
    
    const gsapi = google.sheets({version :'v4', auth : cl});
    const opt = {
        spreadsheetId: this.sheetId,
        range: this.primarySheetName
    };

    let data= await gsapi.spreadsheets.values.get(opt).then();
    this.headerRow = data.data.values[0];
}

// Get primary sheet name (sheet containing the list of tickets to be actioned on, assumed to be first sheet in the doc)

Sheets.prototype.getPrimarySheet = async function(cl){

    const gsapi = google.sheets({version :'v4', auth : cl});

    const opt = {
        spreadsheetId: this.sheetId 
    };

    let data = await gsapi.spreadsheets.get(opt).catch(error => {
        console.log(error.response.status);  
    });
    
    this.primarySheetName = data.data.sheets[0].properties.title;
}

// Update Header Row in sheet and object (Not in use)

Sheets.prototype.setHeaderRow = async function(cl){
    const gsapi = google.sheets({version :'v4', auth : cl});
    var updatedHeader = this.headerRow;
    console.log(updatedHeader);
    if(!updatedHeader.includes("Response Status")){
        updatedHeader.push("Response Status");        
    }
    console.log(updatedHeader);
    if(!updatedHeader.includes("Response Error")){
        updatedHeader.push("Response Error");        
    }
    this.headerRow = updatedHeader;
    console.log(this.headerRow);
    var updateOpt = {
        spreadsheetId : this.sheetId,
        range : `${this.primarySheetName}!A1`,
        valueInputOption : 'USER_ENTERED',
        resource : {values: [updatedHeader]}
    };
    let data = await gsapi.spreadsheets.values.update(updateOpt).catch(error => {
        console.log(error);        
    });
    console.log(this.headerRow);
}

// Test function

Sheets.prototype.addValues = async function(cl, resCodeCol, errMsgCol){
    console.log("Started");
    const gsapi = google.sheets({version :'v4', auth : cl});
    var resCodeColA1 = getA1Notation(0,this.headerRow.indexOf(resCodeCol)).slice(0,-1);
    var errMsgColA1 = getA1Notation(0,this.headerRow.indexOf(errMsgCol)).slice(0,-1);
    var responseCodes = [];
    var errorMessages = [];
    let i = 0;
    while(i<=5000){
        //console.log("Looping #" + i);
        if(i%100==0 && i>0){
            var responseCodeOpt = {
                spreadsheetId : this.sheetId,
                range : this.primarySheetName + "!" + resCodeColA1 + String(i-198),
                valueInputOption : 'USER_ENTERED',
                resource : {values: responseCodes}
            };
            var errorMessageOpt = {
                spreadsheetId : this.sheetId,
                range : this.primarySheetName + "!" + errMsgColA1 + String(i-198),
                valueInputOption : 'USER_ENTERED',
                resource : {values: errorMessages}
            }
            let resCodeUpdate = await gsapi.spreadsheets.values.update(responseCodeOpt).catch(error => {
                console.log(error);        
            });
            let errMsgUpdate = await gsapi.spreadsheets.values.update(errorMessageOpt).catch(error => {
                console.log(error);        
            });
            await new Promise(resolve => setTimeout(resolve, 3000)); 
            responseCodes= [];
            errorMessages = [];
        }
        responseCodes.push(["200"]);
        errorMessages.push(["Success"]);
        i++;
    }

}

// Bulk close method

Sheets.prototype.bulkClose = async function(cl, ticketsCol, resCodeCol, errMsgCol, apiKey, url){

    // Convert API key to B64

    apiKey = Buffer.from(apiKey).toString('base64');
    const gsapi = google.sheets({version :'v4', auth : cl});

    // Get column A1 notation

    var ticketsColA1 = getA1Notation(0,this.headerRow.indexOf(ticketsCol)).slice(0,-1);
    var resCodeColA1 = getA1Notation(0,this.headerRow.indexOf(resCodeCol)).slice(0,-1);
    var errMsgColA1 = getA1Notation(0,this.headerRow.indexOf(errMsgCol)).slice(0,-1);

    var ticketIDs = [];
    var responseCodes = [];
    var errorMessages = [];
    var ticketIDsOpt = {
        spreadsheetId : this.sheetId,
        range : null
    };
    var responseCodeOpt = {
        spreadsheetId : this.sheetId,
        range : null,
        valueInputOption : 'USER_ENTERED',
        resource : {values: responseCodes}
    };
    var errorMessageOpt = {
        spreadsheetId : this.sheetId,
        range : null,
        valueInputOption : 'USER_ENTERED',
        resource : {values: errorMessages}
    }
    let i = 0;
    while(i==0 || ticketIDs.length == 100){
        ticketIDs = [];
        errorMessages = [];
        responseCodes=[];
        if(i%100==0){
            console.log("Started");

            // Get list of tickets (max 100)

            ticketIDsOpt.range = this.primarySheetName + "!" + ticketsColA1 + String(i+2) + ":" + ticketsColA1 +String(i+101);
            console.log(ticketIDsOpt.range);
            let ticketIDsRes = await gsapi.spreadsheets.values.get(ticketIDsOpt);
            ticketIDs = (ticketIDsRes.data.values).map(function(x){
                x = String(x);
                return parseInt(x,10);
            });

            console.log(ticketIDs);
            if(ticketIDs.length == 0){
                continue;
            }

            // Make Bulk Update call

            var response = await unirest.post(`https://${ url }/api/v2/tickets/bulk_update`).headers({
                'Content-Type': 'application/json',
                'Authorization':`Basic ${ apiKey }`
            }).send({
                "bulk_action": {
                    "ids": ticketIDs,
                    "properties":{
                        "status": 5
                    }
                }
            });
            if(response.status >= 300){

                // If rate limits breached, try again after 60 seconds

                await new Promise(resolve => setTimeout(resolve, 60000));
                response = await unirest.post(`https://${ url }/api/v2/tickets/bulk_update`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ${ apiKey }`
                }).send({
                    "bulk_action": {
                        "ids": ticketIDs,
                        "properties":{
                            "status": 5
                        }
                    }
                });

                // End task if breached again

                if(response.status>=300){break;}

                else{console.log(response.body);}
            }

            console.log(response.body);

            // Get Job ID

            var jobID = String(response.body["job_id"]);
            var jobResponse = await unirest.get(`https://${ url }/api/v2/jobs/${ jobID }`).headers({
                'Content-Type': 'application/json',
                'Authorization':`Basic ${ apiKey }`
            });

            if(jobResponse.status >=300){
                await new Promise(resolve => setTimeout(resolve, 60000));
                var jobResponse = await unirest.get(`https://${ url }/api/v2/jobs/${ jobID }`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ${ apiKey }`
                });
                if(jobResponse.status >=300){break;}
            }
            var jobStatus = jobResponse.body["status"];

            // Start polling 10 times or till completion at intervals of 10 secs

            var count = 0;
            while (jobStatus == "IN_PROGRESS" && count < 10){
                console.log("Checking");
                await new Promise(resolve => setTimeout(resolve, 10000));
                jobResponse = await unirest.get(`https://${ url }/api/v2/jobs/${ jobID }`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ${ apiKey }`
                });
                if(jobResponse.status>=300){
                    await new Promise(resolve => setTimeout(resolve, 50000));
                    count+=1
                    continue;
                }
                jobStatus = String(jobResponse.body["status"]);
                console.log(jobStatus);
                count+=1;
            }
            count = 0;

            // Generate status/error message array and job ID payload

            if (jobStatus != "IN_PROGRESS" && jobResponse.body["data"]){
                console.log("Successful update");
                for(let ticketID of ticketIDs){
                    for(let status of jobResponse.body["data"]){
                        if(status.id == ticketID){
                            if(status.success == false){
                                errorMessages.push([JSON.stringify(status.error)]);
                                responseCodes.push([jobID]);                                                    
                            }
                            else{
                                errorMessages.push(["Success"]);
                                responseCodes.push([jobID]); 
                            }
                            break;
                        }
                        continue;
                    }
                }
                console.log(errorMessages);
            }
            else{
                console.log("Unsuccessful update");
                for(let ticketID of ticketIDs){
                    responseCodes.push([jobID]); 
                    errorMessages.push(["Something went wrong"]);
                }
            }
            
            // Send errorMessages list to sheet

            errorMessageOpt.range = this.primarySheetName + "!" + errMsgColA1 + String(i+2);
            responseCodeOpt.range = this.primarySheetName + "!" + resCodeColA1 + String(i+2);
            errorMessageOpt.resource.values = errorMessages;
            responseCodeOpt.resource.values = responseCodes;
            let errMsgUpdate = await gsapi.spreadsheets.values.update(errorMessageOpt).catch(error => {
                console.log(error);        
            });
            let resCodeUpdate = await gsapi.spreadsheets.values.update(responseCodeOpt).catch(error => {
                console.log(error);        
            });

        }
        i+=100;
    }

}

Sheets.prototype.bulkReopen = async function(cl, ticketsCol, resCodeCol, errMsgCol, apiKey, url){
    apiKey = Buffer.from(apiKey).toString('base64');
    const gsapi = google.sheets({version :'v4', auth : cl});
    var ticketsColA1 = getA1Notation(0,this.headerRow.indexOf(ticketsCol)).slice(0,-1);
    var resCodeColA1 = getA1Notation(0,this.headerRow.indexOf(resCodeCol)).slice(0,-1);
    var errMsgColA1 = getA1Notation(0,this.headerRow.indexOf(errMsgCol)).slice(0,-1);
    var rateLimitBreach = 0;
    var ticketIDs = [];
    var responseCodes = [];
    var errorMessages = [];
    var ticketIDsOpt = {
        spreadsheetId : this.sheetId,
        range : null
    };
    var responseCodeOpt = {
        spreadsheetId : this.sheetId,
        range : null,
        valueInputOption : 'USER_ENTERED',
        resource : {values: responseCodes}
    };
    var errorMessageOpt = {
        spreadsheetId : this.sheetId,
        range : null,
        valueInputOption : 'USER_ENTERED',
        resource : {values: errorMessages}
    }
    let i = 0;
    while(i==0 || ticketIDs.length == 100){
        ticketIDs = [];
        errorMessages = [];
        responseCodes=[];
        if(i%100==0){
            console.log("Started");

            // Get list of tickets (max 100)

            ticketIDsOpt.range = this.primarySheetName + "!" + ticketsColA1 + String(i+2) + ":" + ticketsColA1 +String(i+101);
            console.log(ticketIDsOpt.range);
            let ticketIDsRes = await gsapi.spreadsheets.values.get(ticketIDsOpt);
            ticketIDs = (ticketIDsRes.data.values).map(function(x){
                x = String(x);
                return parseInt(x,10);
            });

            console.log(ticketIDs);

            // Make Bulk Update call
            if(ticketIDs.length == 0){
                continue;
            }
            var response = await unirest.post(`https://${ url }/api/v2/tickets/bulk_update`).headers({
                'Content-Type': 'application/json',
                'Authorization':`Basic ${ apiKey }`
            }).send({
                "bulk_action": {
                    "ids": ticketIDs,
                    "properties":{
                        "status": 2
                    }
                }
            });
            if(response.status >= 300){

                // If rate limits breached, try again after 60 seconds

                await new Promise(resolve => setTimeout(resolve, 60000));
                response = await unirest.post(`https://${ url }/api/v2/tickets/bulk_update`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ${ apiKey }`
                }).send({
                    "bulk_action": {
                        "ids": ticketIDs,
                        "properties":{
                            "status": 2
                        }
                    }
                });

                // End task if breached again

                if(response.status>=300){break;}

                else{console.log(response.body);}
            }
            console.log(response.body);

            // Get Job ID

            var jobID = String(response.body["job_id"]);
            var jobResponse = await unirest.get(`https://${ url }/api/v2/jobs/${ jobID }`).headers({
                'Content-Type': 'application/json',
                'Authorization':`Basic ${ apiKey }`
            });

            if(jobResponse.status >=300){
                await new Promise(resolve => setTimeout(resolve, 60000));
                var jobResponse = await unirest.get(`https://${ url }/api/v2/jobs/${ jobID }`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ${ apiKey }`
                });
                if(jobResponse.status >=300){break;}
            }

            var jobStatus = jobResponse.body["status"];

            // Start polling 10 times or till completion at intervals of 10 secs

            var count = 0;
            while (jobStatus == "IN_PROGRESS" && count < 10){
                console.log("Checking");
                await new Promise(resolve => setTimeout(resolve, 10000));
                jobResponse = await unirest.get(`https://${ url }/api/v2/jobs/${ jobID }`).headers({
                    'Content-Type': 'application/json',
                    'Authorization':`Basic ${ apiKey }`
                });
                if(jobResponse.status>=300){
                    await new Promise(resolve => setTimeout(resolve, 50000));
                    count+=1
                    continue;
                }
                jobStatus = String(jobResponse.body["status"]);
                console.log(jobStatus);
                count+=1;
            }
            count = 0;

            // Generate status/error message array and job ID payload

            if (jobStatus != "IN_PROGRESS" && jobResponse.body["data"]){
                console.log("Successful update");
                for(let ticketID of ticketIDs){
                    for(let status of jobResponse.body["data"]){
                        if(status.id == ticketID){
                            if(status.success == false){
                                errorMessages.push([JSON.stringify(status.error)]);
                                responseCodes.push([jobID]);                                                    
                            }
                            else{
                                errorMessages.push(["Success"]);
                                responseCodes.push([jobID]); 
                            }
                            break;
                        }
                        continue;
                    }
                }
                console.log(errorMessages);
            }
            else{
                console.log("Unsuccessful update");
                for(let ticketID of ticketIDs){
                    responseCodes.push([jobID]); 
                    errorMessages.push(["Something went wrong"]);
                }
            }
            
            // Send errorMessages list to sheet

            errorMessageOpt.range = this.primarySheetName + "!" + errMsgColA1 + String(i+2);
            responseCodeOpt.range = this.primarySheetName + "!" + resCodeColA1 + String(i+2);
            errorMessageOpt.resource.values = errorMessages;
            responseCodeOpt.resource.values = responseCodes;
            let errMsgUpdate = await gsapi.spreadsheets.values.update(errorMessageOpt).catch(error => {
                console.log(error);        
            });
            let resCodeUpdate = await gsapi.spreadsheets.values.update(responseCodeOpt).catch(error => {
                console.log(error);        
            });

        }
        i+=100;
    }

}


module.exports = {
    Sheets, getA1Notation
}