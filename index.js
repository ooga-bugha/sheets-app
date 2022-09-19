const express = require('express');
const exphbs = require('express-handlebars');
const bodyParser = require('body-parser');
const methodOverride = require('method-override');
const { Worker } = require("worker_threads");
const path = require('path');
const multer = require('multer');
var storage = multer.memoryStorage();
var upload = multer({ storage: storage });
const { Readable } = require('stream');
const sheetModule = require('./sheetObject');
const { response } = require('express');
const unirest = require('unirest');
const {google} = require('googleapis');



// Init app

const app = express();

// Set engine

app.engine('handlebars', exphbs.engine({ defaultLayout:'main'}));
app.set('view engine', 'handlebars');
app.set('port', process.env.PORT || 3000);

// body-parser

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended : false}));

// use public directory

app.use(express.static(path.join(__dirname, '/public')));

// method-override

app.use(methodOverride('_method'));

// listen on port

app.listen(process.env.PORT || 3000, function(){
    console.log('Server started' );
});


// Main page

app.get('/', function(req, res, next){
    res.render('sheetDetails');
});

// Validation pages

app.get('/validate', function(req, res, next){
    res.render('validated');
});

app.get('/viewSheet', function(req, res, next){
    res.render('viewSheet');
});


//Render Sheet

app.post('/jobDetails', upload.single('creds'), function(req, res, next){

    console.log(req.body);
    
    // Initialize a sheet

    var sheet = new sheetModule.Sheets(req.body.sheetID, JSON.parse(req.file.buffer).client_email, JSON.parse(req.file.buffer).private_key);

    // Validate FD creds

    b64Key = Buffer.from(req.body.api_key).toString('base64');  
    unirest.get(`https://${ req.body.fd_url }/api/v2/agents/me`).headers({
        'Content-Type': 'application/json',
        'Authorization':`Basic ${ b64Key }`
    }).then(response =>{
        if(response.status < 300){
            console.log('Valid API key');

            // Validate Sheet creds

            var client = new google.auth.JWT(
                sheet.creds.client_email, null, sheet.creds.private_key,['https://www.googleapis.com/auth/spreadsheets']
            );
            client.authorize(async function(err,tokens){
                if(err){
                    res.render('sheetDetails',{error: 'Invalid google API creds'});
                    console.log(`Invalid Google API creds`);
                }
                else{
                
                    // Get Primary sheet name

                    await sheet.getPrimarySheet(client).then(async function(){

                        // Get Header Row

                        await sheet.getHeaderRow(client);
                        console.log(`Awesome`);
                        console.log(sheet.headerRow);

                        // Render Job details page with columns and task to be selected

                        res.render('jobDetails', {
                            header:sheet.headerRow, 
                            sheetID:sheet.sheetId, 
                            sheetPrivateKey:sheet.creds.private_key, 
                            sheetClientEmail:sheet.creds.client_email, 
                            apiKey:req.body.api_key, 
                            url:req.body.fd_url});
                    })
                    .catch(err => {
                        
                        // Catch invalid sheet error

                        res.render('sheetDetails',{error: 'Invalid sheet'});
                        console.log("Invalid sheet");
                    });
                }
            });
        }

        // Catch Invalid FD creds
        else{
            res.render('sheetDetails',{error: 'Invalid FD API creds'});
            console.log('Invalid API key');
        }
    }).catch(err => {
        res.render('sheetDetails',{error: 'Something went wrong while checking FD creds'});
    });
});


// Display sheet details and start Worker
app.post('/submitted', upload.none(), function(req, res, next){

    // Check if unique columns are submitted

    if(req.body.ticket_id_column == req.body.response_code || req.body.ticket_id_column== req.body.error_message || req.body.error_message==req.body.response_code){
        res.render('jobDetails', {
            error:'Please select different columns',
            header:req.body.header, 
            sheetID:req.body.sheet_ID, 
            sheetPrivateKey:req.body.sheet_private_key, 
            sheetClientEmail:req.body.sheet_email, 
            apiKey:req.body.api_key,
            url:req.body.fd_url
        });
    }
    else if(req.body.task == "bulk_delete"){
        res.render('jobDetails', {
            error:'Bulk deletion is in development',
            header:req.body.header, 
            sheetID:req.body.sheet_ID, 
            sheetPrivateKey:req.body.sheet_private_key, 
            sheetClientEmail:req.body.sheet_email, 
            apiKey:req.body.api_key,
            url:req.body.fd_url
        });

    }
    else{
        const worker = new Worker('./sheetWorker.js',{workerData : req.body});
        res.render('sheetDetails',{success: 'Job Submitted'});
    }
});