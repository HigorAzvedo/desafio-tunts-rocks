const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const { calendar } = require('googleapis/build/src/apis/calendar');


const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets.readonly',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/spreadsheets.readonly'
];

const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}


async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
async function listMajors(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: '1MD1Yced-DJLQ-ryzDrohyAL5M8W5ynmL7njRhhSurG8',
    range: 'engenharia_de_software!A4:H27',
  });
  
  const data = res.data.values;
  const situacao = [];
  
  data.forEach(d => {
    const [, , faltas, p1, p2, p3] = d
    const media = (Number(p1) + Number(p2) + Number(p3)) / 3
    const aulasNoSemestre = 60;
    const limiteDeFaltas = (aulasNoSemestre * 25) / 100;
    const faltasNum = Number(faltas)
    
    if (faltasNum > limiteDeFaltas) {
      situacao.push(['Reprovado por faltas', 0])
    }
    else if (media >= 70 && faltasNum < limiteDeFaltas) {
      situacao.push(['Aprovado', 0])
    } else if (media >= 50 && faltasNum < limiteDeFaltas){
      const naf = Math.ceil(100 - media)
      situacao.push(['Exame Final' , naf])
      
    } else{
      situacao.push(['Reprovado', 0])
    }
  })

  try {
      await sheets.spreadsheets.values.update({
      spreadsheetId: '1MD1Yced-DJLQ-ryzDrohyAL5M8W5ynmL7njRhhSurG8',
      range: 'engenharia_de_software!G4:H27',
      valueInputOption: 'RAW',
      resource: {
        values: situacao
      },
    })
  }
  
  catch (e){
    console.log(e.errors)
  }

}
authorize().then(listMajors).catch(console.error);