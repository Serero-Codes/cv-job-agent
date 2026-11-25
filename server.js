require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cookieSession = require('cookie-session');
const { google } = require('googleapis');
const { GoogleGenAI, Type } = require('@google/genai');
const AdmZip = require('adm-zip');
const { execFile } = require('child_process');
const { promisify } = require('util');
const os = require('os');
const fs = require('fs').promises;
const path = require('path');

const execFileAsync = promisify(execFile);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY, apiVersion: 'v1beta' });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(
  cookieSession({
    name: 'session',
    keys: [process.env.SESSION_SECRET || 'dev-secret-key'],
    maxAge: 24 * 60 * 60 * 1000, // 24 hours
  })
);

// ============================================================================
// AVAILABLE TAILORED CV FILES IN ROOT FOLDER (__dirname)
// ============================================================================
const CV_TRACKS = [
  {
    key: 'software_ai',
    filename: 'Serero_Immanuel_Mokwena_CV_Software_AI.docx',
    focus: 'Software & AI posts (Python, C#, ASP.NET Core, React/Vue, RAG, OpenAI, Groq, Vector Stores, AWS, SQL/MongoDB)',
  },
  {
    key: 'software_engineering',
    filename: 'Serero_Immanuel_Mokwena_CV_Software_Engineering.docx',
    focus: 'Software Engineering posts (C#, ASP.NET Core MVC, Python, Flask, JavaScript, React, Vue.js, REST APIs, OOP, TDD, SQL/MongoDB)',
  },
  {
    key: 'cloud_engineering',
    filename: 'Serero_Immanuel_Mokwena_CV_Cloud_Engineering.docx',
    focus: 'Cloud Engineering & DevOps posts (AWS VPC/EC2/S3/RDS/Lambda/ALB/IAM/CloudWatch, Terraform IaC, Docker, CI/CD, Azure AZ-900, CCNAv7)',
  },
  {
    key: 'database_administrator',
    filename: 'Serero_Immanuel_Mokwena_CV_Database_Administrator.docx',
    focus: 'Database Administrator & Data posts (PostgreSQL, SQL Server, MySQL, MongoDB, Amazon RDS, DynamoDB, Supabase, schema design, query optimization, XSLT/MXML)',
  },
];

// ============================================================================
// MASTER CV KNOWLEDGE BASE (Single Source of Truth)
// ============================================================================
const MASTER_CV = `
SERERO IMMANUEL MOKWENA
Location: Johannesburg, Gauteng, South Africa | Phone: (+27) 76 533 4475 | Email: sereroemmanuel4@gmail.com
Links: LinkedIn (Serero Mokwena) | GitHub (Serero-Codes) | Portfolio (Serero Mokwena Portfolio)

Technical Skills:
- Programming & Backend: Python, C#, JavaScript, SQL, HTML5, CSS3, Flask, ASP.NET Core MVC, RESTful APIs, LINQ
- Frontend & Frameworks: React.js, Vue.js, Bootstrap, Vite, responsive UI integration
- Cloud & DevOps: AWS (VPC, EC2, S3, RDS, Lambda, ALB, IAM, CloudWatch, CloudTrail, SSM), Azure Fundamentals (AZ-900), Terraform (IaC), Docker, GitHub Actions, CI/CD
- Databases & Systems: PostgreSQL, Microsoft SQL Server, MySQL, MongoDB, Supabase, DynamoDB, SQLAlchemy, schema design, query optimization, Unix/Linux, XML/MXML, XSLT
- AI & Advanced Technologies: Retrieval-Augmented Generation (RAG), OpenAI API, Groq API, Vector Stores, Prompt Engineering, Agentic AI workflows, NLP basics
- Practices & Methodologies: Agile/Scrum, MVC Architecture, Test-Driven Development (TDD), Git/GitHub code reviews, OOP, Data Structures & Algorithms

Professional Experience:
1. Backend & Cloud Developer | TNG Solutions (Host from Capaciti) | 2026 (September – Present) | Johannesburg, South Africa
  - Lead a 3-developer team maintaining a client event management system (virtual, hybrid, and in-person events) across frontend/backend repositories and MongoDB.
  - Implement backend services and APIs to ensure system reliability, MongoDB data integrity, and seamless repository integration.
  - Review and approve GitHub pull requests, facilitate team sync meetings, and deliver progress feedback to senior leadership.
2. Cloud Engineer Intern | UVU Africa (Capaciti) | 2026 (April – Present) | Johannesburg, South Africa
  - Architect and provision multi-tier AWS infrastructure (VPC, EC2, S3, RDS, Lambda, ALB) automated with Terraform (IaC) and Docker.
  - Enforce least-privilege IAM roles, VPC network segmentation, SSM Parameter Store secrets, and CloudWatch Auto Scaling.
  - Engineered an agentic AI solution achieving 3rd place in the 2026 AWS Summit Johannesburg GameDay challenge.
3. Software Integration Intern (Murex Analyst) | Kion Consulting | 2025 (June–July & Nov–Dec) | Johannesburg (Hybrid)
  - Integrated backend services, SQL databases, RESTful APIs, and XSLT/MXML data pipelines with Unix shell scripts and SQL query optimizations.
  - Built an internal Python/Flask AI assistant using OpenAI APIs to automate debugging and authoring of trade-booking formulas.
4. Assistant Manager, Tech & Innovations | GeekKulchaNMU | 2025 | Nelson Mandela University
  - Facilitated student developer workshops in C# and OOP principles; organized hackathons and mentored aspiring developers.

Education:
- National Diploma in Information Technology (Software Development) | Nelson Mandela University | Feb 2023 – Dec 2025
- Higher Certificate in Information Technology | Nelson Mandela University | 2022

Selected Technical Projects:
- 3-Tier Containerized AWS Architecture (AWS, Terraform, Docker, PostgreSQL)
- Contact Book API (Python, Flask, Vue.js, Docker, EC2, ALB, S3, CI/CD, CloudWatch)
- Siyafunda Siyapassa – Multilingual AI Tutoring Platform (Python, JS, Flask, React, Groq API, NLP – 2nd Place Kion Hackathon)
- Murex Minimalizer / Mx3 AI Assistant (Python, Flask, Vue.js, OpenAI API, Vector Store, PostgreSQL)
- Hospital Management System (C#, ASP.NET Core MVC, SQL Server, ASP.NET Identity)

Certifications:
- Microsoft Certified: Azure Fundamentals (AZ-900)
- SFIA: AWS Cloud Solutions Architect Certification
- Cisco: CCNAv7 Introduction to Networks
- IBM: Python for Data Science & AI Development | IBM: Generative AI: Prompt Engineering Basics
- AWS / DeepLearning.AI: Generative AI with LLMs | Coursera: Architecting Solutions on AWS | Google Cloud: Introduction to Generative AI
`;

// ============================================================================
// DOCX XML HELPERS (Update ONLY Introduction & Convert DOCX -> PDF)
// ============================================================================
function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function unescapeXml(str) {
  return String(str)
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function findSummaryParagraphBounds(documentXml) {
  const headingIdx = documentXml.indexOf('PROFESSIONAL SUMMARY');
  if (headingIdx === -1) {
    throw new Error('Could not locate PROFESSIONAL SUMMARY section in DOCX.');
  }
  const headingEndIdx = documentXml.indexOf('</w:p>', headingIdx);
  if (headingEndIdx === -1) {
    throw new Error('Malformed heading paragraph in DOCX.');
  }
  const summaryStartIdx = documentXml.indexOf('<w:p>', headingEndIdx);
  const summaryEndIdx = documentXml.indexOf('</w:p>', summaryStartIdx);
  if (summaryStartIdx === -1 || summaryEndIdx === -1) {
    throw new Error('Could not locate introduction paragraph after PROFESSIONAL SUMMARY.');
  }
  return { summaryStartIdx, summaryEndIdx: summaryEndIdx + '</w:p>'.length };
}

async function extractSummaryFromDocx(docxPath) {
  const zip = new AdmZip(docxPath);
  const documentXml = zip.readAsText('word/document.xml');
  const { summaryStartIdx, summaryEndIdx } = findSummaryParagraphBounds(documentXml);
  const summaryParaXml = documentXml.slice(summaryStartIdx, summaryEndIdx);
  const textMatches = [...summaryParaXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)];
  return textMatches.map((m) => unescapeXml(m[1])).join('');
}

async function loadAvailableCvsFromRoot() {
  const available = [];
  for (const track of CV_TRACKS) {
    const fullPath = path.join(__dirname, 'templates', track.filename);
    try {
      await fs.access(fullPath);
      const currentSummary = await extractSummaryFromDocx(fullPath);
      available.push({ ...track, fullPath, currentSummary });
    } catch {
      // File not present in root folder yet
    }
  }
  if (available.length === 0) {
    throw new Error('No tailored .docx CV files found in the root folder.');
  }
  return available;
}

function resolveSelectedCvPath(selectedCvFile) {
  const match = CV_TRACKS.find((t) => t.filename === selectedCvFile);
  const chosenFilename = match ? match.filename : CV_TRACKS[0].filename;
  return {
    filename: chosenFilename,
    fullPath: path.join(__dirname, 'templates', chosenFilename),
  };
}

async function updateDocxIntroductionOnly(docxPath, updatedSummary) {
  const zip = new AdmZip(docxPath);
  const documentXml = zip.readAsText('word/document.xml');
  const { summaryStartIdx, summaryEndIdx } = findSummaryParagraphBounds(documentXml);
  const originalParaXml = documentXml.slice(summaryStartIdx, summaryEndIdx);

  // Preserve original paragraph properties (<w:pPr>...</w:pPr>)
  const pPrMatch = originalParaXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  const pPrXml = pPrMatch ? pPrMatch[0] : '';

  // Rebuild ONLY the summary text run inside the existing paragraph formatting
  const newRunXml =
    `<w:r><w:rPr>` +
    `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>` +
    `<w:b w:val="0"/><w:i w:val="0"/><w:color w:val="1A1A1A"/><w:sz w:val="22"/>` +
    `</w:rPr><w:t xml:space="preserve">${escapeXml(updatedSummary.trim())}</w:t></w:r>`;

  const newParaXml = `<w:p>${pPrXml}${newRunXml}</w:p>`;
  const updatedDocumentXml =
    documentXml.slice(0, summaryStartIdx) + newParaXml + documentXml.slice(summaryEndIdx);

  zip.updateFile('word/document.xml', Buffer.from(updatedDocumentXml, 'utf-8'));
  return zip.toBuffer();
}

async function convertDocxBufferToPdf(docxBuffer, baseName = 'cv') {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'cv-pdf-'));
  const loProfile = path.join(tmpDir, 'lo_profile');
  const loProfileUrl = 'file:///' + loProfile.replace(/\\/g, '/').replace(/^\//, '');
  const inputDocx = path.join(tmpDir, `${baseName}.docx`);
  const outputPdf = path.join(tmpDir, `${baseName}.pdf`);

  try {
    await fs.writeFile(inputDocx, docxBuffer);
    const loBin = process.env.LIBREOFFICE_BIN || 'libreoffice';
    await execFileAsync(
      loBin,
      [
        `-env:UserInstallation=${loProfileUrl}`,
        '--headless',
        '--nologo',
        '--nolockcheck',
        '--norestore',
        '--convert-to',
        'pdf',
        '--outdir',
        tmpDir,
        inputDocx,
      ],
      { timeout: 45000 }
    );
    return await fs.readFile(outputPdf);
  } finally {
    await fs.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function generateTailoredCV(selectedCvFile, approvedSummary) {
  const { filename, fullPath } = resolveSelectedCvPath(selectedCvFile);

  console.log('============================================================');
  console.log(`[CV SELECTED]: ${filename}`);
  console.log(`[SOURCE PATH]: ${fullPath}`);
  console.log('============================================================');

  const updatedDocxBuffer = await updateDocxIntroductionOnly(fullPath, approvedSummary);
  return await convertDocxBufferToPdf(updatedDocxBuffer, 'Serero_Mokwena_Immanuel_CV');
}

// ============================================================================
// GMAIL OAUTH 2.0 CLIENT & ROUTES
// ============================================================================
function getOAuth2Client() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
}

app.get('/auth/status', (req, res) => {
  res.json({ authenticated: Boolean(req.session && req.session.tokens) });
});

app.get('/auth/google', (req, res) => {
  const oauth2Client = getOAuth2Client();
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: ['https://www.googleapis.com/auth/gmail.compose'],
  });
  res.redirect(url);
});

app.get('/auth/google/callback', async (req, res) => {
  try {
    const { code } = req.query;
    const oauth2Client = getOAuth2Client();
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    req.session.tokens = tokens;
    res.redirect('/');
  } catch (err) {
    console.error('OAuth Callback Error:', err);
    res.status(500).send('Authentication failed. Check server logs.');
  }
});

app.post('/auth/logout', (req, res) => {
  req.session = null;
  res.json({ ok: true });
});

// ============================================================================
// GEM 1 ENDPOINT: SELECT MOST RELEVANT CV & ADJUST ONLY ITS INTRODUCTION
// ============================================================================
app.post('/api/analyze', upload.single('screenshot'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Please upload a job post screenshot.' });
    }

    const availableCvs = await loadAvailableCvsFromRoot();
    const cvOptionsText = availableCvs
      .map(
        (cv, i) =>
          `${i + 1}. File: ${cv.filename}\n   Focus: ${cv.focus}\n   Current Introduction: "${cv.currentSummary}"`
      )
      .join('\n\n');

    const gem1SystemInstruction = `
You are an expert Technical Recruiter and CV Optimization Specialist dedicated to tailoring Serero Immanuel Mokwena’s CV for specific job postings.

STRICT RULES:
1. Extract the core role title, company name, location, recipient email, required subject line, required attachments/documents, and key technical/qualification requirements from the job post image.
2. Select the MOST RELEVANT CV file from the available root-folder CVs listed below based on the job posting. Set "selectedCvFile" to that exact filename.
3. ONLY make targeted adjustments to the selected CV's "Current Introduction" (Professional Summary) so it aligns with the job post.
   - Do NOT replace the introduction wholesale or rewrite it from scratch.
   - Keep the existing structure, tone, and core phrasing intact, weaving in or emphasizing matching skills/keywords from the job post that Serero actually possesses.
   - Leave the CV header/leading title, skills, experience, projects, education, and certifications completely untouched.
4. Do NOT claim Serero possesses any tool, language, or certification not explicitly listed in his Master CV.
5. Keep the adjusted Professional Summary concise (3–4 impactful, ATS-friendly sentences).

AVAILABLE ROOT-FOLDER CVs:
${cvOptionsText}

MASTER CV:
${MASTER_CV}
`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        {
          inlineData: {
            mimeType: req.file.mimetype,
            data: req.file.buffer.toString('base64'),
          },
        },
        'Analyze this job posting screenshot, select the most relevant CV file, and make targeted adjustments to its existing introduction.',
      ],
      config: {
        systemInstruction: gem1SystemInstruction,
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            selectedCvFile: { type: Type.STRING },
            selectionReason: { type: Type.STRING },
            roleTitle: { type: Type.STRING },
            companyName: { type: Type.STRING },
            location: { type: Type.STRING },
            recipientEmail: { type: Type.STRING },
            requiredSubjectLine: { type: Type.STRING },
            requiredDocuments: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
            },
            matchedKeywords: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
            },
            updatedProfessionalSummary: { type: Type.STRING },
          },
          required: [
            'selectedCvFile',
            'selectionReason',
            'roleTitle',
            'companyName',
            'recipientEmail',
            'requiredSubjectLine',
            'requiredDocuments',
            'matchedKeywords',
            'updatedProfessionalSummary',
          ],
        },
      },
    });

    const data = JSON.parse(response.text);

    // Log selected CV to terminal for manual verification
    console.log('============================================================');
    console.log(`[CV SELECTED]: ${data.selectedCvFile}`);
    console.log(`[TARGET ROLE]: ${data.roleTitle} @ ${data.companyName}`);
    console.log(`[REASON]:      ${data.selectionReason}`);
    console.log('============================================================');

    res.json(data);
  } catch (err) {
    console.error('Model 1 Error:', err);
    res.status(500).json({ error: 'Failed to analyze job screenshot.' });
  }
});

// ============================================================================
// OPTIONAL PREVIEW ENDPOINT: TEST ADJUSTED DOCX -> PDF IN BROWSER
// ============================================================================
app.post('/api/preview-cv', async (req, res) => {
  try {
    const { selectedCvFile, approvedSummary } = req.body;
    const pdfBuffer = await generateTailoredCV(selectedCvFile, approvedSummary);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'inline; filename="Serero Mokwena Immanuel CV.pdf"');
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF Preview Error:', err);
    res.status(500).json({ error: 'Failed to generate CV preview. Check root .docx files and LibreOffice.' });
  }
});

// ============================================================================
// GEM 2 ENDPOINT: GENERATE EMAIL + CONVERTED PDF & PUSH DRAFT TO GMAIL
// ============================================================================
app.post('/api/draft-gmail', async (req, res) => {
  try {
    if (!req.session || !req.session.tokens) {
      return res.status(401).json({ error: 'Not authenticated with Gmail. Please connect Gmail first.' });
    }

    const { jobDetails, selectedCvFile, approvedSummary } = req.body;
    const chosenCv = selectedCvFile || (jobDetails && jobDetails.selectedCvFile);

    // 1. Update ONLY the selected .docx introduction & convert to PDF
    const tailoredPdfBuffer = await generateTailoredCV(chosenCv, approvedSummary);
    const pdfBase64 = tailoredPdfBuffer.toString('base64');
    const attachmentFilename = 'Serero Mokwena Immanuel CV.pdf';

    // 2. Generate tailored application email body
    const gem2SystemInstruction = `
You are an executive career communications specialist writing job application emails on behalf of Serero Immanuel Mokwena.

STRICT RULES:
1. Subject Line: If requiredSubjectLine is provided and non-empty, use it verbatim. Otherwise, ALWAYS use this exact format: "Application for [Role Title] - Serero Immanuel Mokwena".
2. Length: Keep the email body concise (120–160 words), direct, and professional. Never use clichés like "I hope this email finds you well."
3. Required Structure & Paragraph Spacing (ALWAYS separate every section with a blank line \\n\\n):
   - Greeting: Start on its own line with "Dear Hiring Team," (or "Dear [Company Name] Recruitment Team," if company name is known).
   - Opening Paragraph: 1–2 sentences stating the exact role applied for and Serero's core qualification.
   - Alignment Paragraph: 2–3 sentences connecting Serero's relevant experience (TNG Solutions, UVU Africa/Capaciti, Kion Consulting) and projects directly to the job requirements.
   - Attached Documents: Briefly list the attached CV (and any other requested documents using "- " bullets).
   - Conclusion: 1–2 polite closing sentences confirming availability, inviting an interview, and thanking the reader for their time and consideration.
   - Signature Block (always separated by a blank line before "Kind regards,"):

Kind regards,

Serero Immanuel Mokwena
Johannesburg, Gauteng
Phone: (+27) 76 533 4475
Email: sereroemmanuel4@gmail.com
LinkedIn: linkedin.com/in/serero-mokwena | GitHub: github.com/Serero-Codes

4. Output plain text only for the email body (no Markdown ** or ##). Ensure every paragraph and the conclusion are separated by \\n\\n.
`;

    const promptPayload = `
Job Details: ${JSON.stringify(jobDetails)}
Selected CV: ${chosenCv}
Approved CV Summary: ${approvedSummary}
`;

    const emailGenResponse = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [promptPayload],
      config: {
        systemInstruction: gem2SystemInstruction,
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            to: { type: Type.STRING },
            subject: { type: Type.STRING },
            body: { type: Type.STRING },
          },
          required: ['to', 'subject', 'body'],
        },
      },
    });

    const emailDraft = JSON.parse(emailGenResponse.text);
    const recipient = emailDraft.to || (jobDetails && jobDetails.recipientEmail) || '';
    const consistentSubject =
      (jobDetails && jobDetails.requiredSubjectLine && jobDetails.requiredSubjectLine.trim()) ||
      `Application for ${(jobDetails && jobDetails.roleTitle) || 'the Position'} - Serero Immanuel Mokwena`;
    const formattedBody = emailDraft.body.replace(/\r?\n/g, '\r\n');
    


    // 3. Build RFC 2822 Multipart MIME Message with PDF Attachment
    const boundary = `----=_Part_${Date.now()}_${Math.random().toString(36).substring(2)}`;

    const mimeParts = [
      `To: ${recipient}`,
      `Subject: ${consistentSubject}`,
      // ...
      formattedBody,
      'MIME-Version: 1.0',
      `Content-Type: multipart/mixed; boundary="${boundary}"`,
      '',
      `--${boundary}`,
      'Content-Type: text/plain; charset="UTF-8"',
      'Content-Transfer-Encoding: 7bit',
      '',
      emailDraft.body,
      '',
      `--${boundary}`,
      `Content-Type: application/pdf; name="${attachmentFilename}"`,
      'Content-Transfer-Encoding: base64',
      `Content-Disposition: attachment; filename="${attachmentFilename}"`,
      '',
      pdfBase64,
      '',
      `--${boundary}--`,
    ];

    const encodedMessage = Buffer.from(mimeParts.join('\r\n'))
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    // 4. Authenticate Gmail client & create draft with attachment
    const oauth2Client = getOAuth2Client();
    oauth2Client.setCredentials(req.session.tokens);
    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    const draftRes = await gmail.users.drafts.create({
      userId: 'me',
      requestBody: {
        message: {
          raw: encodedMessage,
        },
      },
    });

    res.json({
      success: true,
      draftId: draftRes.data.id,
      selectedCvFile: chosenCv,
      attachedFile: attachmentFilename,
      emailPreview: {
        to: recipient,
        subject: emailDraft.subject,
        body: emailDraft.body,
      },
      gmailDraftUrl: 'https://mail.google.com/mail/u/0/#drafts',
    });
  } catch (err) {
    console.error('Model 2 / Gmail Draft Error:', err);
    res.status(500).json({ error: 'Failed to generate email, attach CV, or create Gmail draft.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));