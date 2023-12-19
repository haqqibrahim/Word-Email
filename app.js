const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs').promises; // Use the promises version of fs
const path = require('path');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const nodemailer = require('nodemailer');

const app = express();
const PORT = process.env.PORT || 5000;

let transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'dev.omari.ai@gmail.com',
    pass: 'pvag lwwq csgo qena',
  },
});

const htmlContent = `
  <!DOCTYPE html>
  <html>
  <head>
    <title>This is an example of how the Email will look</title>
    <style>
      /* CSS styles */
    </style>
  </head>
  <body>
    <div class="container">
      <h1>This is the document</h1>
      <p>Below is a Word document with the data</p>
    </div>
  </body>
  </html>
`;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));

app.post('/submit', async (req, res) => {
  const formData = req.body;
  console.log('Form Data:', formData);

  try {
    const outputDocPath = await createWordDocument(formData);

    const mailOptions = {
      from: 'dev.omari.ai@gmail.com',
      to: formData.senderEmail,
      subject: 'Word document from the form',
      html: htmlContent,
      attachments: [
        {
          filename: 'docs.docx',
          content: await fs.readFile(outputDocPath),
        },
      ],
    };

    await sendEmail(mailOptions);
    res.send('Form submitted successfully!');
  } catch (error) {
    console.error('Error processing form:', error);
    res.status(500).send('Internal Server Error');
  }
});

async function createWordDocument(formData) {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun('This is a test email'),
              new TextRun({
                text: formData.firstName,
                bold: true,
              }),
              new TextRun({
                text: formData.lastName,
                bold: true,
              }),
              new TextRun({
                text: formData.senderEmail,
                bold: true,
              }),
            ],
          }),
        ],
      },
    ],
  });

  const outputDir = 'output';
  const outputDocFilename = 'docs.docx';
  const outputDocPath = path.join(outputDir, outputDocFilename);

  await fs.mkdir(outputDir, { recursive: true });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outputDocPath, buffer);

  console.log('Document saved to:', outputDocPath);
  return outputDocPath;
}

async function sendEmail(mailOptions) {
  try {
    const info = await transporter.sendMail(mailOptions);
    console.log('Email sent:', info.response);
  } catch (error) {
    console.error('Failed to send the email:', error);
    throw error; // Re-throw the error to be caught in the main try/catch block
  }
}

app.listen(PORT, () => {
  console.log(`Server is running at http://localhost:${PORT}`);
});
