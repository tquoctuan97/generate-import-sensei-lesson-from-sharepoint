const fs = require("fs");
const readline = require('readline');
const dotenv = require('dotenv');
const { exec } = require('child_process');

// Load environment variables from .env file
dotenv.config();

const appSettings = {
  tenantId: process.env.TENANT_ID,
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
};

async function main(shareLinkUrl) {
  try {
    // Check input
    if (!shareLinkUrl || typeof shareLinkUrl !== 'string' || !shareLinkUrl.trim()) {
      throw new Error('Please provide a valid SharePoint URL');
    }
    
    const accessToken = await fetchAccessToken();
    const shareUrlBase64 = convertShareLinkToBase64(shareLinkUrl);
    const childrenDriveItemList = await fetchChildrenDriveItemList(accessToken, shareUrlBase64);
    
    if (!childrenDriveItemList || childrenDriveItemList.length === 0) {
      console.log("No items found in the shared folder");
      return;
    }

    console.log(`Found ${childrenDriveItemList.length} items in the shared folder`);

    const videoDriveItemList = childrenDriveItemList.filter((item) =>
      item.file && item.file.mimeType && item.file.mimeType.includes("video/")
    );
    
    console.log(`Found ${videoDriveItemList.length} video items`);

    if (videoDriveItemList.length === 0) {
      console.log("No video files found");
      return;
    }

    const dataToWrite = videoDriveItemList.map((item) => {
      const title = item.name.replace(/\.mp4$/i, "");
      return {
        Lesson: title,
        Description: getEmbedCode({ sharepointIds: item.sharepointIds, title }),
        Status: "publish",
        Prerequisite: "",
      };
    });

    writeToCSV(dataToWrite);
  } catch (error) {
    console.error("An error occurred:", error.message);
    process.exit(1);
  }
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

async function fetchAllItemsRecursively(accessToken, items, shareUrlBase64) {
  // Ensure items is an array
  if (!Array.isArray(items)) {
    console.warn('Received non-array items:', items);
    return [];
  }

  let allItems = [...items];
  
  for (const item of items) {
    if (item?.folder?.childCount > 0) {

      const { siteId, listItemUniqueId: itemId } = item.sharepointIds;

      const folderUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/children?%24select=sharepointids%2Cname%2Cfile%2Cfolder`;
      const options = {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        body: undefined,
      };

      try {
        const response = await fetch(folderUrl, options);
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        
        // Ensure data.value is an array before processing
        if (!data.value || !Array.isArray(data.value)) {
          console.warn(`Invalid response format for folder ${item.name}:`, data);
          continue;
        }
        
        const nestedItems = await fetchAllItemsRecursively(accessToken, data.value, shareUrlBase64);
        allItems = [...allItems, ...nestedItems];
      } catch (error) {
        console.error(`Error fetching items from folder ${item.name}:`, error);
      }
    }
  }

  return allItems;
}

async function fetchChildrenDriveItemList(accessToken, shareUrlBase64) {
  const url = `https://graph.microsoft.com/v1.0/shares/u!${shareUrlBase64}/driveItem/children?%24select=sharepointids%2Cname%2Cfile%2Cfolder`;

  const options = {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: undefined,
  };

  try {
    const response = await fetch(url, options);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();

    let normalItem = [], folderItem = [];

    data.value.forEach((item) => {
      if(item?.folder?.childCount > 0) {
        folderItem.push(item)
      } else {
        normalItem.push(item)
      }
    })

    const itemInside = await fetchAllItemsRecursively(accessToken, folderItem, shareUrlBase64);
    
    return [...normalItem, ...itemInside];
  } catch (error) {
    console.error(error);
    return [];
  }
}

function writeToCSV(data) {
  if (!data || !Array.isArray(data) || data.length === 0) {
    console.warn("No data to write to CSV");
    return;
  }

  // Escape CSV fields properly to handle commas, quotes, etc.
  const escapeCSV = (field) => {
    if (field === null || field === undefined) return '';
    const stringField = String(field);
    // If the field contains a comma, double quote, or newline, wrap it in double quotes
    if (stringField.includes(',') || stringField.includes('"') || stringField.includes('\n')) {
      // Replace double quotes with two double quotes (standard CSV)
      return `"${stringField.replace(/"/g, '""')}"`;
    }
    return stringField;
  };

  const headerData = Object.keys(data[0]).join(",") + "\n";

  // Convert the array of objects to a CSV string
  const rowData = data
    .map((item) => {
      return Object.values(item).map(escapeCSV).join(",");
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
  
  const platform = process.platform;
  
  const isWSL = process.platform === 'linux' && require('fs').existsSync('/proc/version') && 
                require('fs').readFileSync('/proc/version', 'utf-8').toLowerCase().includes('microsoft');
  
  const openCommand = isWSL || platform === 'win32' ? 'explorer.exe' :
                     platform === 'darwin' ? 'open' :
                     'xdg-open';
                     
  exec(`${openCommand} ${outputDir}`, (error) => {
    if (error) {
      console.error('Cannot open output directory:', error);
    }
  });
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
  main(shareLinkUrl.trim());
  rl.close();
});