{/* GEt the data from  excel sheet */}


// const { google } = require("googleapis");}

// const { google } = require("googleapis");

// async function fetchAndUploadToSlack() {
//   try {
//     const auth = new google.auth.GoogleAuth({
//       keyFile: "credentials.json",
//       scopes: ["https://www.googleapis.com/auth/spreadsheets"],
//     });

//     const authClientObject = await auth.getClient();
//     const googleSheetInstance = google.sheets({ version: "v4", auth: authClientObject });

//     const spreadsheetId = "19kUCaanHSUVIG1CxKLSbjVj0X9VfE0FsuQsnTtcW3e4";

//     // Test if the service account has access
//     const sheetMetadata = await googleSheetInstance.spreadsheets.get({
//       spreadsheetId,
//     });
//     console.log("Sheet metadata fetched successfully:", sheetMetadata.data);

//     // get the data from the google sheet
//     const response = await googleSheetInstance.spreadsheets.values.get({
//       spreadsheetId,
//       range: "Sheet1",
//       valueInputOption: "USER_ENTERED",
//       resource: {
//         values: [["marmeto", "marmeto@.com", "hello"]],
//       },
//     });

//     console.log("Sheet updated successfully:", response.body);

//     // console.log("Sheet updated successfully:", response.data);
//   } catch (error) {
//     console.error("Error updating the sheet:", error.message);
//     console.error("Full error details:", error);
//   }
// }

// fetchAndUploadToSlack();

// const { google } = require("googleapis");

// async function fetchAndUploadToSlack() {
//   try {
//     // Authenticate using the credentials file
//     const auth = new google.auth.GoogleAuth({
//       keyFile: "credentials.json", // Path to your credentials file
//       scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"], // Use readonly scope if fetching data only
//     });

//     // Create an authenticated client
//     const authClientObject = await auth.getClient();

//     // Initialize the Google Sheets API instance
//     const googleSheetInstance = google.sheets({ version: "v4", auth: authClientObject });

//     // Google Sheet ID
//     const spreadsheetId = "19kUCaanHSUVIG1CxKLSbjVj0X9VfE0FsuQsnTtcW3e4";

//     // Fetch metadata (optional, for debugging)
//     const sheetMetadata = await googleSheetInstance.spreadsheets.get({
//       spreadsheetId,
//     });
//     console.log("Sheet metadata fetched successfully:", sheetMetadata.data);

//     // Get data from the sheet
//     const response = await googleSheetInstance.spreadsheets.values.get({
//       spreadsheetId,
//       range: "Sheet1", // Specify the range to fetch
//     });

//     // Log the fetched data
//     const rows = response.data.values; // The rows of the sheet
//     if (rows.length) {
//       console.log("Data fetched successfully from the sheet:");
//       rows.forEach((row, index) => {
//         console.log(`Row ${index + 1}:`, row);
//       });
//     } else {
//       console.log("No data found in the sheet.");
//     }
//   } catch (error) {
//     console.error("Error fetching data from the sheet:", error.message);
//     console.error("Full error details:", error);
//   }
// }

// fetchAndUploadToSlack();

{/* GEt the data from  excel sheet convert into pdf */}

// const { google } = require("googleapis");
// const { PDFDocument, rgb } = require("pdf-lib");
// const fs = require("fs");

// async function fetchAndConvertToPDF() {
//   try {
//     // Authenticate using the credentials file
//     const auth = new google.auth.GoogleAuth({
//       keyFile: "credentials.json", // Path to your credentials file
//       scopes: ["https://www.googleapis.com/auth/spreadsheets"], // Use readonly scope if fetching data only
//     });

//     // Create an authenticated client
//     const authClientObject = await auth.getClient();

//     // Initialize the Google Sheets API instance
//     const googleSheetInstance = google.sheets({ version: "v4", auth: authClientObject });

//     // Google Sheet ID
//     const spreadsheetId = "19kUCaanHSUVIG1CxKLSbjVj0X9VfE0FsuQsnTtcW3e4";

//     // Get data from the sheet
//     const response = await googleSheetInstance.spreadsheets.values.get({
//       spreadsheetId,
//       range: "Sheet1", // Specify the range to fetch
//     });

//     const rows = response.data.values; // The rows of the sheet
//     if (rows.length) {
//       console.log("Data fetched successfully from the sheet:");
//       rows.forEach((row, index) => {
//         console.log(`Row ${index + 1}:`, row);
//       });

//       // Convert data to a PDF
//       const pdfBytes = await createPDF(rows);

//       // Save the PDF
//       fs.writeFileSync("output.pdf", pdfBytes);
//       console.log("PDF created successfully: output.pdf");
//     } else {
//       console.log("No data found in the sheet.");
//     }
//   } catch (error) {
//     console.error("Error fetching data from the sheet:", error.message);
//     console.error("Full error details:", error);
//   }
// }

// async function createPDF(rows) {
//   // Create a new PDF document
//   const pdfDoc = await PDFDocument.create();

//   // Add a page to the document
//   const page = pdfDoc.addPage();

//   // Define text properties
//   const fontSize = 12;
//   const margin = 50;
//   let yPosition = page.getHeight() - margin; // Start at the top of the page

//   // Add title
//   page.drawText("Google Sheet Data", {
//     x: margin,
//     y: yPosition,
//     size: 16,
//     color: rgb(0, 0, 0),
//   });
//   yPosition -= 20; // Move down after the title

//   // Add table headers (assumes the first row is the header)
//   const headers = rows[0];
//   page.drawText(headers.join(" | "), {
//     x: margin,
//     y: yPosition,
//     size: fontSize,
//     color: rgb(0, 0.2, 0.8),
//   });
//   yPosition -= 20; // Move down after the headers

//   // Add table rows
//   rows.slice(1).forEach((row) => {
//     if (yPosition < margin) {
//       // Add a new page if the current page is full
//       page = pdfDoc.addPage();
//       yPosition = page.getHeight() - margin;
//     }
//     page.drawText(row.join(" | "), {
//       x: margin,
//       y: yPosition,
//       size: fontSize,
//       color: rgb(0, 0, 0),
//     });
//     yPosition -= 20; // Move down after each row
//   });

//   // Serialize the document to bytes
//   return pdfDoc.save();
// }

// // Run the script
// fetchAndConvertToPDF();



{/* GEt the data from  excel sheet convert into pdf then send to the email*/}

const { google } = require("googleapis");
const { PDFDocument, rgb } = require("pdf-lib");
const fs = require("fs");
const nodemailer = require("nodemailer");
const cron = require("node-cron");

async function fetchAndConvertToPDF() {
  try {
    // Authenticate using the credentials file
    const auth = new google.auth.GoogleAuth({
      keyFile: "credentials.json", // Path to your credentials file
      scopes: ["https://www.googleapis.com/auth/spreadsheets"], // Use readonly scope if fetching data only
    });

    // Create an authenticated client
    const authClientObject = await auth.getClient();

    // Initialize the Google Sheets API instance
    const googleSheetInstance = google.sheets({ version: "v4", auth: authClientObject });

    // Google Sheet ID
    const spreadsheetId = process.env.SPREADSHEET_ID;

    // Get data from the sheet
    const response = await googleSheetInstance.spreadsheets.values.get({
      spreadsheetId,
      range: "Sheet1", // Specify the range to fetch
    });

    const rows = response.data.values; // The rows of the sheet
    if (rows.length) {
      console.log("Data fetched successfully from the sheet:");
      rows.forEach((row, index) => {
        console.log(`Row ${index + 1}:`, row);
      });

      // Convert data to a PDF
      const pdfBytes = await createPDF(rows);

      // Save the PDF
      fs.writeFileSync("output.pdf", pdfBytes);
      console.log("PDF created successfully: output.pdf");

      // Set up cron job to send the PDF via email
      scheduleEmailTask();
    } else {
      console.log("No data found in the sheet.");
    }
  } catch (error) {
    console.error("Error fetching data from the sheet:", error.message);
    console.error("Full error details:", error);
  }
}

async function createPDF(rows) {
  // Create a new PDF document
  const pdfDoc = await PDFDocument.create();

  // Define text properties
  const fontSize = 12;
  const margin = 50;
  let yPosition;

  // Add a page to the document
  let page = pdfDoc.addPage();
  yPosition = page.getHeight() - margin; // Start at the top of the page

  // Add title
  page.drawText("Google Sheet Data", {
    x: margin,
    y: yPosition,
    size: 16,
    color: rgb(0, 0, 0),
  });
  yPosition -= 20; // Move down after the title

  // Add table headers (assumes the first row is the header)
  const headers = rows[0];
  page.drawText(headers.join(" | "), {
    x: margin,
    y: yPosition,
    size: fontSize,
    color: rgb(0, 0.2, 0.8),
  });
  yPosition -= 20; // Move down after the headers

  // Add table rows
  rows.slice(1).forEach((row) => {
    if (yPosition < margin) {
      // Add a new page if the current page is full
      page = pdfDoc.addPage();
      yPosition = page.getHeight() - margin;
    }
    page.drawText(row.join(" | "), {
      x: margin,
      y: yPosition,
      size: fontSize,
      color: rgb(0, 0, 0),
    });
    yPosition -= 20; // Move down after each row
  });

  // Serialize the document to bytes
  return pdfDoc.save();
}

function scheduleEmailTask() {
  // Run the task every second using setInterval
  setInterval(async () => {
    console.log("Sending email with PDF...");

    // Read the generated PDF file
    const pdfPath = "output.pdf";

    // Set up the email transporter
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: "anoops1909@gmail.com", // Your email
        pass: "efxd dcbf epnf veyi", // Your email password or app-specific password
      },
    });

    // Prepare the email options
    const mailOptions = {
      from: "anoops1909@gmail.com",
      to: "anoop@marmeto.com", // Target email address
      subject: "Google Sheet Data as PDF",
      text: "Please find the attached PDF containing the data from the Google Sheet.",
      attachments: [
        {
          filename: "output.pdf",
          path: pdfPath,
        },
      ],
    };

    // Send the email
    try {
      await transporter.sendMail(mailOptions);
      console.log("Email sent successfully.");
    } catch (error) {
      console.error("Error sending email:", error.message);
    }
  }, 1000); // Every second (1000 ms)
  console.log("Task scheduled to run every second.");
}

// Run the fetch and convert process
fetchAndConvertToPDF();
