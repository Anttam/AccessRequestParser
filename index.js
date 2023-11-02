//this app is used to read an access request word document
// it takes the data and console logs it in JSON format to the console


// for this public upload, I removed the filepath and changed a few names with REDACTED


//this is an app that I have written to help me with my current role

const officeParser = require('officeparser');
let file = "FILEPATH" 

//REGEX used to parse the data from the form
const firstName = /(?<=First Name\W+)\w+/
const lastName = /(?<=Last Name\W+)\w+/
const email = /(?:(?!)|(?<=User E-Mail +))(?:(?! User).)*/;
const phone = /(?<=User Phone #\W+)\d+-?\d+-?\d+/;
const description = /(?<=Division\/Company\W+)\w+/;
const roles = /(?:(?!)|(?<=REDACTED \d\W+))[^\WChoose Role](?:(?! Add).)*/g;
const roleEnv = /(?<=(Add)\s+)\w+/

//formatting the roles so they are the exact same as they are entered in Active Directory
let  formatRoles = function (user){
  //tailoring each role to the environment the access is for; PROD,TEST,MODEL etc.
  const env = `ICE_${user.roleEnv.toUpperCase()}_REDACTED_`
  for(let role in user.roles){
  user.roles[role] = env.concat(user.roles[role].toUpperCase().replaceAll(" ","_"))
  }
  //standard roles given to every user
  user.roles.push('ManageEngineGroup');
  user.roles.push('REDACTED');
  user.roles.push(`ICE_${env}_PAYER_ALL_0`);
  user.roles.push('rlink_REDACTED')
  return user;
}
//the officeparser NPM package returns a promise

// I chose to use async as opposed to a callback for readability 

async function getUser(){try {
    const data = await officeParser.parseOfficeAsync(file);
    return {firstName: data.match(firstName)[0],
           lastName:data.match(lastName)[0],
           email: data.match(email)[0],
           phone: data.match(phone)[0],
           description: data.match(description)[0],
           roles: data.match(roles),
           roleEnv: data.match(roleEnv)[0],
           }
} catch (err) {
    console.log(err);
}
                        }
// anon aysnc function to call all the functions in the proper order                        
(async () => {
  return await getUser()
})().then((user) => formatRoles(user))
    .then((user) => console.log(user))
