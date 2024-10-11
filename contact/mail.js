const http = require("http");
const nodemailer = require("nodemailer");
const { parse } = require("querystring");
const fs = require("fs");
const xlsx = require("xlsx"); // Importing the xlsx module

let server = http.createServer((req, res) => {
    if (req.method === 'POST') {
        let form_URLENCODED = "application/x-www-form-urlencoded";
        
        if (req.headers["content-type"] === form_URLENCODED) {
            let body = "";

            // Collect form data
            req.on('data', chunk => {
                body += chunk.toString();
            });

            req.on('end', () => {
                let formData = parse(body);
                let email = formData.email;  // Get email from form
                let name = formData.name;    // Get name from form
                let message = formData.message;  // Get message from form
                
                console.log("Name:", name);
                console.log("Email:", email);
                console.log("Message:", message);

                // Save form data to Excel
                saveToExcel({ Name: name, Email: email, Message: message });

                res.writeHead(200, { "Content-Type": "text/html" });
                res.end(`<h1>Thank you for contacting me, ${name}!</h1><p>Your message has been received.</p>`);

                // Send email using Nodemailer
                let transporter = nodemailer.createTransport({
                    service: "gmail",
                    auth: {
                        user: "hariharan98704@gmail.com",  
                        pass: "bhcaozpchzjguljr" // Replace with your email password or app-specific password
                    }
                });

                let mailOptions = {
                    from: "hariharan98704@gmail.com",
                    to: email,
                    subject: "Thank you for reaching out!",
                    html: `<h1>Hi ${name},</h1>
                           <p>Thank you for your message. I will get back to you soon.</p>
                           <p><strong>Your message:</strong> ${message}</p>`
                };

                transporter.sendMail(mailOptions, (err, info) => {
                    if (err) {
                        console.error("Error sending email:", err);
                    } else {
                        console.log("Email sent:", info.response);
                    }
                });
            });
        } else {
            res.writeHead(400, { "Content-Type": "text/html" });
            res.end("<h1>Invalid Content-Type</h1>");
        }
    } else {
        res.writeHead(200, { "Content-Type": "text/html" });
        fs.createReadStream("./contactme.html", "utf-8").pipe(res);
    }
});

server.listen(3000, (err) => {
    if (err) throw err;
    console.log("Server is running at http://localhost:3000");
});

// Function to save form data to Excel
function saveToExcel(data) {
    const filePath = './formData.xlsx'; // Path to the Excel file

    let workbook;
    let worksheet;

    // Check if the file exists
    if (fs.existsSync(filePath)) {
        // Read existing file
        workbook = xlsx.readFile(filePath);
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
        // Create a new workbook and worksheet
        workbook = xlsx.utils.book_new();
        worksheet = xlsx.utils.json_to_sheet([]);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    }

    // Append new data to the worksheet
    const worksheetData = xlsx.utils.sheet_to_json(worksheet);
    worksheetData.push(data);

    // Convert JSON back to worksheet and save
    const newWorksheet = xlsx.utils.json_to_sheet(worksheetData);
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

    xlsx.writeFile(workbook, filePath);
}
