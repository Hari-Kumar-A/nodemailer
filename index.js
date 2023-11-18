const nodemailer = require('nodemailer');
const axios = require('axios');
const xlsx = require('xlsx');
require('dotenv').config();
const fs = require('fs');


//setting mail and service
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: 'harikumar3868@gmail.com',
        pass: process.env.PASS,
    }
});

//axios to get files via stream
async function downloadFile(url, destination) {
    const response = await axios({
        method: 'get',
        url: url,
        responseType: 'stream',
    });

    //writing the stream data into a <filename>.pdf
    const writestream=fs.createWriteStream(destination);
    response.data.pipe(writestream);

    return new Promise((resolve, reject) => {
        response.data.on('end', () => resolve());
        response.data.on('error', (err) => reject(err));
    });
}

async function sendEmail(member) {
    const certificateUrl = member.certificate;
    const localFilePath = './downloaded-certificate.pdf';

    try {
        if (certificateUrl) {
            // Download the file if there is a certificate URL
            await downloadFile(certificateUrl, localFilePath);
        }

        const mailOptions = {
            from: 'harikumar3868@gmail.com',
            to: member.email,
            subject: 'Festember \'23 Participation Certificate',
            text: `✨Greetings from Festember! Here's your Participation Certificate for your active enthusiasm and efforts in bringing Festember'23 to life.`,
        };

        // Add attachment only if there is a certificate URL
        if (certificateUrl) {
            const certificateAttachment = fs.readFileSync(localFilePath);
            mailOptions.attachments = [
                {
                    filename: `${member.name} certificate.pdf`,
                    content: certificateAttachment,
                }
            ];
        }

        await transporter.sendMail(mailOptions);
        console.log(`Email sent to ${member.email}`);

    } catch (error) {
        console.error('Error sending email:', error);
    } finally {
        // Clean up: Remove the downloaded file if it exists
        if (fs.existsSync(localFilePath)) {
            fs.unlinkSync(localFilePath);
        }
    }
}


async function sendEmailsToMembers() {
    const workbook = xlsx.readFile('participants.xlsx');
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const membersList = xlsx.utils.sheet_to_json(sheet);

    for (const member of membersList) {
        await sendEmail(member);
    }

    transporter.close();
}

sendEmailsToMembers();