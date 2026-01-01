/**

Office Script: Sync Availability to Web Dashboard

Automates data push from Excel to external monitoring systems */ function main(workbook: ExcelScript.Workbook) { let sheet = workbook.getWorksheet("Availability_Summary"); let dataRange = sheet.getUsedRange(); let values = dataRange.getValues();

// Logic to format data as JSON for web endpoints let payload = values.map(row => { return { staffName: row[0], qualification: row[1], availabilityStatus: row[2] }; });

console.log("Availability Payload Ready for API Sync:"); console.log(JSON.stringify(payload));

// Example Webhook integration would go here }