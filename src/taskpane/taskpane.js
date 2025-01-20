/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { initializeAPI } from '../openalgoApi';
import { Settings } from '../settings';

// The initialize function must be run each time a new page is loaded
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Load saved settings
    const settings = Settings.getSettings();
    document.getElementById('api-host').value = settings.host;
    document.getElementById('api-key').value = settings.apikey;
    
    // Add event listener for the save button
    document.getElementById('save-settings').onclick = saveSettings;
    
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function saveSettings() {
    const host = document.getElementById('api-host').value || 'http://127.0.0.1:5000';
    const apikey = document.getElementById('api-key').value;

    try {
        const success = await Settings.saveSettings(host, apikey);
        if (success) {
            document.getElementById('status').textContent = 'Settings saved successfully!';
            document.getElementById('status').style.color = 'green';
        } else {
            throw new Error('Failed to save settings');
        }
    } catch (error) {
        document.getElementById('status').textContent = 'Error saving settings: ' + error.message;
        document.getElementById('status').style.color = 'red';
    }
}
