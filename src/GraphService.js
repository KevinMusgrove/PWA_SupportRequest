import { async } from 'q';
var graph = require('@microsoft/microsoft-graph-client');

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken.accessToken);
    }
  });
  return client;
}

export async function sendEmail(accessToken,firstName,lastName,location,email,phone,device,priority,issueDesc,managerEmail,files){
    const client = getAuthenticatedClient(accessToken);
    var t = managerEmail;
    debugger;
    let sendMail;
    if(managerEmail)
    {
       sendMail = {      
        message: {
          subject: "New Support Request",
          attachments: files,         
           body: {
            contentType: "Html",
            content: 
            `<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding:5px;">
            <span style="font-weight: bold">User Having Issue: </span>${firstName} ${lastName}
            <br/>
            <span style="font-weight: bold">User's Location: </span>${location}
            <br/>
            <span style="font-weight: bold">User's Email: </span>${email}<a href="mailto:${email}"></a>
            <br/>
            <span style="font-weight: bold">User's Phone Number: </span><a href="tel:${phone}">${phone}</a>
            <br/>
            <span style="font-weight: bold">Type of Device with Issue:</span> ${device}
            <br/>
            <span style="font-weight: bold">Priority Level:</span> ${priority}
            <br/><br/>
            <span style="font-weight: bold">Issue Description:</span> ${issueDesc}
            </body>`
          },
          toRecipients: [
            {
              emailAddress: {
                address: "Kevin.Musgrove@atb-tech.com"
              }
            }
          ],          
          ccRecipients:[
            {
              emailAddress:{
                address:managerEmail
              }
            }
          ]                    
        },        
        saveToSentItems: "true"
      };
    }
    else{
        sendMail = {      
        message: {
          subject: "New Support Request",
          attachments: files,         
           body: {
            contentType: "Html",
            content: 
            `<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding:5px;">
            <span style="font-weight: bold">User Having Issue: </span>${firstName} ${lastName}
            <br/>
            <span style="font-weight: bold">User's Location: </span>${location}
            <br/>
            <span style="font-weight: bold">User's Email: </span>${email}<a href="mailto:${email}"></a>
            <br/>
            <span style="font-weight: bold">User's Phone Number: </span><a href="tel:${phone}">${phone}</a>
            <br/>
            <span style="font-weight: bold">Type of Device with Issue:</span> ${device}
            <br/>
            <span style="font-weight: bold">Priority Level:</span> ${priority}
            <br/><br/>
            <span style="font-weight: bold">Issue Description:</span> ${issueDesc}
            </body>`
          },
          toRecipients: [
            {
              emailAddress: {
                address: "Kevin.Musgrove@atb-tech.com"
              }
            }
          ]          
        },        
        saveToSentItems: "true"
      };
    }          
      let res = await client.api('/me/sendMail').post(sendMail);
      return res;
}

export async function getUserBlob(accessToken){
  const blob = await fetch('https://graph.microsoft.com/beta/me/photo/$value', {
    headers: {
      'Authorization': "Bearer " + accessToken
    }
  })
  .then(async response => {
    if(response.status == 200)
    {
      const blob = await response.blob();
      return blob
    }
    else{
      return ""
    }
    
  }).catch(err =>{
    debugger;
  });
  return blob;
}

export async function getUserManager(accessToken){
  try{
    const client = getAuthenticatedClient(accessToken);
    const manager = await client.api('/me/manager').get();    
    if(manager){
      return manager.mail;
    }              
  }
  catch(err){
    return "";
  }  
}

export async function getUserDetails(accessToken) {
  const client = getAuthenticatedClient(accessToken);
  const user = await client.api('/me').get();    
  return user;
}