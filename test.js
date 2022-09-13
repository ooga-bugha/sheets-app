const {google} = require('googleapis');
const keys = {"type": "service_account","project_id": "my-project-to-fetch-tickets","private_key_id": "e252adb349a91e551f7ea8f0cafe95126f90c8b1","private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDgbDCkkSCmI+LC\nzkuGH7x3RS3Frxi8jeRosn1RIxeULECVmrUQCKU6EFNauCm8Hs3W+N7AlaN0z5nj\nqQMSE/WimWDuEmzichWoXXIqxTGyog/V7Y1MjuSIZ7RBMgKk/RZIPvxuPCadz5Sy\nIII+biVNxzSrW7SgqMg5hLeNA2Ma7Jm2e+7absCpQQNZEkFsAJoDojt6iZ8q9pZE\na2wfAgxzBKeGzxaJFWu5GZvNz62IYv4vK139RALQqfXuWTw+htTXZL1h2ZQvGbYF\nlI0m0i02idRlXsVM+PLbw5PjdG1JtJ6JVVFM966ryZc8SQbNyu/8kEwjYCs81DER\nK48de9m/AgMBAAECggEADuHdz+v32Eyk6vo3M+vC2b3yrRtRbp+SOAcuHEVRePf+\nSG17+FY6bsFKZce0rM72I31ZeDf0IPjrYZeBpp9AOMonDWKXaeTa3tFjksaE5y2s\nEymvpxYKvajy4Sfp2PsXkS7sGntOrERpve499NnlC2VcbsikD8thi77rVSSWW+CY\n+wJG6qWoxisSDjToY9yNodBbywnxsRgpJdKH46BeEOjpFu/e+h5mWUqePL714dnv\nLE9DlQxWpaHqblHY7TDk5AO6xben/hYxIIyBxZ6E6i7GCZjxNYxoTixGc57K1ZjN\nTnqD+SD6CIM9WaOIQR1OdMlFPe6ul89kkiPzTOEpcQKBgQD28xi4Q4rgEmiCulm0\n19l2thFKBQBcK8xEYYKp+SZ4+bI6fkrJcMD3M+V1u6dcPr6710EhtzNWkaX4ThLx\nysf74FbYXHaa6R5kd2Bh/7LHMiXmQ1Us8+J5Vfoqc0NerQrbbnt5PlXeItrPvE7g\nCdznYHMwOmqX+dRL15JIyoupHQKBgQDopb5Iy6KG8kZnzgYIWeZBNHkrOKTnZck7\nCEYUmvH6XTnNw4qq/WNjsRb1sE4Y7UoxPyY083xAN3FQbqkPQCfPmsXd3gNUWvDh\nPUpMP/aa30x5IFSXnLum9n8JjCgx0GclhKeDQVoeM5BN3miLknnH7UzwIb94QR2P\nmuDvkTtziwKBgHuX7ydJppe+ns/OtFbuMMhZFw4Usrlusi0HIH4xVC/3yFu+GW/4\nHpuaPZ1O7dQdExiwAsj9B5SsEITVjmW1N6G4Bb8Dh9jAE5X0qShi8PcBAjbcPCTj\natWPUkUsqusXb/eis+laaV9j0l9lv5QhW43xl7Trh63IO5g5q90CgiOBAoGAKCGX\nHmWKJq8aOAPRBJXFY1AS6sK9p3Dmcnlt5VYJEcANHZJylCZbg7HjnQQJpMEiADa9\nd3rc3xLxSAeewBO4ClbPdQM8HcwGK0RwUZDjEDoerfJGxVRzBk0VAuebc2RYtp8Y\nakrWqckJRnVsIU9mFHe5wt5/cdYBrGjyDkFGORsCgYEA6tM6vlx+LAuM865oqkN6\ncQ5rKTSVe8J4itEp8DUkzK5WzbC9XH0Tk5wETu3a0MkK4kzEd50XnwZJDEO1RTur\nz4C+qJLxh6EHOTcq/HdPjQb47FFUr0a2LXmVN5Zb+N1wyZF7aCVY/L75WUSpQ2aZ\nA4ns5rVk7prIxje0bo65wFo=\n-----END PRIVATE KEY-----\n","client_email": "sheets@my-project-to-fetch-tickets.iam.gserviceaccount.com","client_id": "104507883129542484945","auth_uri": "https://accounts.google.com/o/oauth2/auth","token_uri": "https://oauth2.googleapis.com/token","auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs","client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/sheets%40my-project-to-fetch-tickets.iam.gserviceaccount.com"};
const externalObjs = require('./sheetObject');
let testSheet = new externalObjs.Sheets('14ANEmi826WqgzqPSHgOx46dwrs4sRiS0anjshrcNttQ',keys.client_email,keys.private_key);
const client = new google.auth.JWT(
    testSheet.creds.client_email, null, testSheet.creds.private_key,['https://www.googleapis.com/auth/spreadsheets']
);

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
        console.log(testSheet.headerRow.indexOf("Response Status"));
        await testSheet.bulkClose(client,"Ticket ID","Response Status","Response Error","B2VUFrZyXLx9UKtIWx5","testgames.freshdesk.com");

    }
});



// async function gsrun(cl){

//     const gsapi = google.sheets({version :'v4', auth : cl});

//     const opt = {
//         //spreadsheetId: '1QGaSfBlUerGKJOQlflCGeypjJbPwzSYANt9J1Q0I5aM',
//         spreadsheetId: '1DuP1l04wZ8AVUt-vCVzg5LiMNtXyXMereQ7JlVxD50M',
//         range:'Sheet1'
//     };

//     let data= await gsapi.spreadsheets.values.get(opt);
//     console.log(data.data.values[0].length);
//     console.log(data.data.values.length-1);
// }

