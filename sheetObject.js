const {google} = require('googleapis');

function Sheets(sheetId, client_email, private_key){
    this.sheetId = sheetId;
    this.creds = {'client_email':client_email,'private_key':private_key};
}
Sheets.prototype.getHeaderRow = async function(cl){
    
    const gsapi = google.sheets({version :'v4', auth : cl});
    const opt = {
        spreadsheetId: this.sheetId,
        range: this.primarySheetName
    };

    let data= await gsapi.spreadsheets.values.get(opt).then();
    this.headerRow = data.data.values[0];
}

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

Sheets.prototype.addValues = async function(cl){
    console.log("Started");
    const gsapi = google.sheets({version :'v4', auth : cl});
    updatedHeader = [];
    let i = 0;
    while(i<5000){
        //console.log("Looping #" + i);
        if(i%200==0 && i>0){
            var range = this.primarySheetName + "!E" + String(i-198);
            var updateOpt = {
                spreadsheetId : this.sheetId,
                range : range,
                valueInputOption : 'USER_ENTERED',
                resource : {values: updatedHeader}
            };
            console.log(updateOpt.range);
            console.log(updateOpt.resource.values.length);
            let data = await gsapi.spreadsheets.values.update(updateOpt).catch(error => {
                console.log(error);        
            });
            await new Promise(resolve => setTimeout(resolve, 3000)); 
            console.log(`Push # ${i/200}`);
            updatedHeader= [];
        }
        updatedHeader.push(["200","Success"]);
        i++;
    }

}


module.exports = {
    Sheets
}