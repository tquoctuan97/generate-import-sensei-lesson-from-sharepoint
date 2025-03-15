const fs = require("fs");
const readline = require('readline');
const dotenv = require('dotenv');

// Load environment variables from .env file
dotenv.config();

const appSettings = {
  tenantId: process.env.TENANT_ID,
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
};

async function main(shareLinkUrl) {
  const accessToken = await fetchAccessToken();

  const shareUrlBase64 = convertShareLinkToBase64(shareLinkUrl);

  const childrenDriveItemList = await fetchChildrenDriveItemList(accessToken, shareUrlBase64);

  console.log(childrenDriveItemList);

  const videoDriveItemList = childrenDriveItemList.filter((item) =>
    item.file.mimeType.includes("video/")
  );

  const dataToWrite = videoDriveItemList.map((item) => {
    const title = item.name.replace(".mp4", "");
    return {
      Lesson: title,
      Description: getEmbedCode({ sharepointIds: item.sharepointIds, title }),
      Status: "publish",
      Prerequisite: "",
    };
  });

  writeToCSV(dataToWrite);
}

async function fetchAccessToken() {
  const url = `https://login.microsoftonline.com/${appSettings.tenantId}/oauth2/v2.0/token`;
  const options = {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: appSettings.clientId,
      client_secret: appSettings.clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  };

  try {
    const response = await fetch(url, options);
    const data = await response.json();
    return data.access_token;
  } catch (error) {
    console.error(error);
  }
}

function convertShareLinkToBase64(shareLinkUrl) {
  return Buffer.from(shareLinkUrl).toString("base64");
}

async function fetchChildrenDriveItemList(accessToken, shareUrlBase64) {
  const url = `https://graph.microsoft.com/v1.0/shares/u!${shareUrlBase64}/driveItem/children?%24select=sharepointids%2Cname%2Cfile`;

  const options = {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: undefined,
  };

  try {
    const response = await fetch(url, options);
    const data = await response.json();
    return data.value;
  } catch (error) {
    console.error(error);
  }
}

function writeToCSV(data) {
  const headerData = Object.keys(data[0]).join(",") + "\n";

  // Convert the array of objects to a CSV string
  const rowData = data
    .map((item) => {
      return Object.values(item).join(",");
    })
    .join("\n");

  const csv = headerData + rowData;

  // Create output directory if it doesn't exist
  const outputDir = 'output';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }

  const now = new Date();
  const formattedDate = `${now.getFullYear()}-${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}_${now.getHours().toString().padStart(2, '0')}-${now.getMinutes().toString().padStart(2, '0')}`;
  const filePath = `${outputDir}/data-${formattedDate}.csv`;
  fs.writeFileSync(filePath, csv);

  console.log(`CSV file written successfully to ${filePath}`);
}

function getEmbedCode({ title, sharepointIds }) {
  const { listItemUniqueId, siteUrl } = sharepointIds;

  return `<div style="max-width: 1280px"><div style="position: relative; padding-bottom: 56.25%; height: 0; overflow: hidden;"><iframe src="${siteUrl}/_layouts/15/embed.aspx?UniqueId=${listItemUniqueId}&embed=%7B%22hvm%22%3Atrue%2C%22ust%22%3Atrue%7D&referrer=StreamWebApp&referrerScenario=EmbedDialog.Create" width="1280" height="720" frameborder="0" scrolling="no" allowfullscreen title="${title}" style="border:none; position: absolute; top: 0; left: 0; right: 0; bottom: 0; height: 100%; max-width: 100%;"></iframe></div></div>`;
}

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

rl.question("Input your sharepoint url: ", (shareLinkUrl) => {
  main(shareLinkUrl);
  rl.close();
});