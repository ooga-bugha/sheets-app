const unirest = require('unirest');

module.exports = {
    validateAPIKey : async function(key, url, callback){

        b64Key = Buffer.from(key).toString('base64');  
        var response = await unirest.get(`https://${ url }/api/v2/agents/me`).headers({
            'Content-Type': 'application/json',
            'Authorization':`Basic ${ b64Key }`
        })
        if(response.status < 300){
            console.log("valid key");
            return 0;
        }
        else{
            console.log("invalid key");
            callback();
        }
    }
}