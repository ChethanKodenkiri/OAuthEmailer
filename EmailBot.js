const xlsx = require("xlsx");
const nodemailer = require("nodemailer");
const { google } = require("googleapis");
const { exit } = require("process");
require('dotenv').config();

// Load your Excel file
const workbook = xlsx.readFile("./Senior Software Engineer.xlsx");
const sheetName = "Senior Software Engineer";
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet);

// OAuth2 Configuration
const client_id = process.env.CLIENT_ID;
const client_secret = process.env.CLIENT_SECRET;
const redirect_uri = process.env.REDIRECT_URI;
const refreshtoken = process.env.REFRESH_TOKEN;
const accessToken = process.env.ACCESSTOKEN;

// OAuth2 Client
const oAuth2Client = new google.auth.OAuth2(
  client_id,
  client_secret,
  redirect_uri
);

oAuth2Client.setCredentials({ refresh_token: refreshtoken });

// Create transporter
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    type: "OAuth2",
    user: "chethankodenkiri.career@gmail.com",
    clientId: client_id,
    clientSecret: client_secret,
    refreshToken: refreshtoken,
    accessToken: accessToken,
  },
});

const sendEmail = async (row) => {
  const { Name, Company, Email, Role, Link } = row; // Adjust column names accordingly
  const nameParts = Name.split(" ");
  const name = nameParts[0];
  // Mail options
  const mailOptions = {
    from: "Chethan Kodenkiri <chethankodenkiri.career@gmail.com>",
    to: Email,
    subject: `Request for an Interview Opportunity - ${Role} at ${Company}`,
    html: `
<p>Dear <b>${name}</b>,</p>

  <p>I hope this email finds you well.</p>

  <p>I am <b>Chethan Kodenkiri</b>, a Senior Software Engineer at <b>Persistent Systems</b>, and I came across your LinkedIn post indicating that <b>${Company}</b> is currently looking for a <b>${Role}</b>. I am writing to express my interest in this position and to share a brief overview of my qualifications:</p>

  <ul>
    <li><b>3+ years</b> of hands-on experience in the <b>Frontend domain</b>.</li>
    <li><b>3 years</b> as a <b>Frontend Developer</b> at <a href="https://www.persistent.com/">Persistent Systems</a>.</li>
    <li>Expertise in <b>JavaScript</b>, <b>TypeScript</b>, and <b>React.js</b>.</li>
    <li>Familiar with tools and technologies such as <b>REST</b>, <b>Jest</b>, <b>React Testing Library</b>, <b>Java</b>, <b>Selenium</b>, <b>Jenkins</b>, <b>Docker</b>, and <b>Git</b>.</li>
    <li>A Bachelor's degree in <b>Electronics and Communication</b> from <b>Visvesvaraya Technological University (Graduated in 2020)</b>.</li>
  </ul>

  <p>I am currently serving my notice period and can <b>join within 15 days</b> of receiving an offer. I believe I would be a valuable addition to your team, and I would greatly appreciate the opportunity for an interview to discuss how my skills align with the needs of <b>${Company}</b>.</p>

  <p>For your reference, I have attached my <b><a href="https://drive.google.com/file/d/1sVhUAy6kWqAcbi8oUN9n0QIgyAPEI36t/view">Resume</a></b> and provided links to my <b><a href="https://www.linkedin.com/in/chethan-kodenkiri">LinkedIn Profile</a></b> and <b><a href="https://github.com/ChethanKodenkiri">GitHub</a></b>. If applicable, I am also including a link to the <b><a href="${Link}">${Role}</a></b> Opening.</p>

  <p>Thank you for your time and consideration. I look forward to the possibility of connecting and exploring this opportunity further.</p>

  <p>Best regards,<br>
  <b>Chethan Kodenkiri</b><br>
  Senior Software Developer, Persistent Systems<br>
  Phone: +91 6363719787<br>
  <a href="https://www.linkedin.com/in/chethan-kodenkiri">LinkedIn</a> | <a href="https://github.com/ChethanKodenkiri">GitHub</a></p>`,
  };

  // Send the email
  try {
    await transporter.sendMail(mailOptions);
    console.log("Email sent to", Email);
  } catch (error) {
    console.error("Error sending email:", row.Email, error);
  }
};

const sendEmailsSynchronously = async () => {
  for (const row of data) {
    await sendEmail(row);
    await new Promise((resolve) => setTimeout(resolve, Math.random() * 90000)); // Pause for 1 minute (adjust the duration as needed)
  }
  console.log("Done Sending emails");
  exit();
};

// Call the function to send emails
sendEmailsSynchronously();
