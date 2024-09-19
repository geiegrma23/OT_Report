const nodemailer = require('nodemailer');
const fs = require('fs');
const axios = require('axios');
const ExcelJS = require('exceljs');

const url = 'https://ksp.shipstoresoftware.com/Prod/api/CustomReport/GenerateReport';

function formatDate(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
}

// Get the current date
const currentDate = new Date();

// Calculate the date range
const twoDaysAgo = new Date(currentDate.getTime() - (2 * 24 * 60 * 60 * 1000));
const fiveDaysBefore = new Date(twoDaysAgo.getTime() - (5 * 24 * 60 * 60 * 1000));

// Format the dates as mm/dd/yyyy
const twoDaysAgoFormatted = formatDate(twoDaysAgo);
const fiveDaysBeforeFormatted = formatDate(fiveDaysBefore);

const data = {
  NumberOfResults: 2500,
  ReportKey: 'e8e3ee2a-b2d0-49a4-bc7e-5c62e187f096',
  CustomParameters: [
    {
      Value: fiveDaysBeforeFormatted,
      Conjunction: 1,
      FieldName: 'Shipments.ShipDate',
      Comparer: 9,
    },
    {
      Value: twoDaysAgoFormatted,
      Conjunction: 1,
      FieldName: 'Shipments.ShipDate',
      Comparer: 11,
    },
  ],
};

axios
  .post(url, data)
  .then((response) => {
    const payload = response.data;
    const results = JSON.parse(payload.Results[0]);

    // Create a new workbook
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Results');

    // Define the header row
    const headers = [
      // Header definitions...
      { id: 'Id', title: 'ID' },
      { id: 'OrderNumber', title: 'Order Number' },
      { id: 'ref1', title: 'Ref' },
      { id: 'CarrierName', title: 'Carrier Name' },
      { id: 'HostServiceCode_Out', title: 'Host Service Code Out' },
      { id: 'TrackingNumber', title: 'Tracking Number' },
      { id: 'ShipDate', title: 'Ship Date' },
      { id: 'Contact', title: 'Contact' },
      { id: 'Company', title: 'Company' },
      { id: 'Address1', title: 'Address 1' },
      { id: 'City', title: 'City' },
      { id: 'State', title: 'State' },
      { id: 'Zip', title: 'Zip' },
      { id: 'Country', title: 'Country' },
      { id: 'ListTotalCharge', title:'ListTotalCharge'},
      { id: 'TotalPackageCount', title:'TotalPackageCount'},
      { id: 'TotalPalletCount', title: 'TotalPalletCount'},
      { id: 'Weight', title:'Weight'},
      { id: 'Dimensions', title:'Dimensions'},
      { id: 'udf1', title:'udf1'},
      { id: 'udf10', title:'udf10'},
      { id: 'CarrierName1', title:'CarrierName1'},
      { id: 'TrackingStatus', title:'TrackingStatus'}
    ];

    // Write the headers to the first row
    sheet.getRow(1).values = headers.map((header) => header.title);

    // Write the data rows
    results.forEach((result) => {
      const stringResult = {};

      // Convert numbers to strings with fixed-point notation
      for (const [key, value] of Object.entries(result)) {
        if (typeof value === 'number') {
          stringResult[key] = key === 'TrackingNumber' ? `"${value.toString()}"` : Number(value).toFixed(2);
        } else {
          stringResult[key] = value;
        }
      }

      // Check criteria for writing to the "Tracking Status" column
      if (
        stringResult['TrackingStatus'] === 'Shipment information sent to FedEx' ||
        stringResult['TrackingStatus'] === 'Shipment exception' ||
        stringResult['TrackingStatus'] === 'Shipper created a label, UPS has not received the package yet.' ||
        stringResult['TrackingStatus'] === 'Error: The package identifier value is missing or invalid.' ||
        stringResult['TrackingStatus'] === 'Shipping Label Created, USPS Awaiting Item' ||
        stringResult['TrackingStatus'] === null ||
        stringResult['TrackingStatus'] === undefined ||
        stringResult['TrackingStatus'] === '' ||
        stringResult['TrackingStatus'].includes('a shipping label has been prepared')
      ) {
        // Write the values to a new row
        sheet.addRow(Object.values(stringResult));
      }
    });

    // Save the workbook to a file
    const filePath = 'results.xlsx';
    return workbook.xlsx.writeFile(filePath)
      .then(() => {
        console.log('Result saved to', filePath);

        // Send email with attachment
        const smtpConfig = {
          host: 'smtp.office365.com',
          port: 587,
          secure: false,
          requireTLS: true,
          auth: {
            user: 'KSP_NO_REPLY@ksp3pl.com',
            pass: 'Kitting23!',
          },
        };

        const transporter = nodemailer.createTransport(smtpConfig);

        const mailOptions = {
          from: 'KSP_NO_REPLY@ksp3pl.com',
          to: 'mike.geiger@ksp3pl.com',
          subject: 'At Risk Shipment Report',
          text: 'Do not reply.  This report is for information purposes only.  Please identify your at risk shipments and take appropriate action.',
          attachments: [
            {
              filename: 'results.xlsx',
              path: filePath,
            },
          ],
        };

        transporter.sendMail(mailOptions, (error, info) => {
          if (error) {
            console.error('Error sending email:', error);
          } else {
            console.log('Email sent:', info.response);
          }
        });
      });
  })
  .catch((error) => {
    console.error(error);
  });
