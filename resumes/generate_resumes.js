const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const { 
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, 
  BorderStyle
} = require('docx');

const resumesDir = __dirname;

const contactInfo = {
  name: "ABIMBOLA SAMUEL IBITAYO",
  phone: "+234 706 176 7012",
  email: "psalmstringss@gmail.com",
  location: "Lagos, Nigeria",
  github: "https://github.com/Psalmstrings",
  portfolio: "https://psalmstrings.vercel.app",
  education: [
    {
      school: "SQI College of ICT, Ibadan",
      credential: "Professional Diploma in Software Engineering",
      year: "2025"
    },
    {
      school: "Federal University of Technology, Akure (FUTA)",
      credential: "Bachelor of Technology (B.Tech)",
      year: "2024"
    }
  ],
  languages: "English (Fluent / Professional), Yoruba (Native)"
};

// ==========================================
// RESUME DATA DEFINITIONS (2 BALANCED PAGES)
// ==========================================

const resumesData = [
  // 1. FULL-STACK DEVELOPER
  {
    id: "fullstack",
    fileNamePrefix: "Abimbola_Samuel_Full_Stack_Developer_Resume",
    roleTitle: "FULL-STACK SOFTWARE ENGINEER (MERN & CLOUD ARCHITECTURE)",
    summary: "Performance-driven Full-Stack Software Engineer with extensive experience designing, engineering, and deploying end-to-end web applications, multi-tenant enterprise platforms, and transaction-critical digital systems. Expert across the MERN ecosystem (MongoDB, Express.js, React 19, Node.js), RESTful API design, Paystack financial workflows, cryptographic authentication (HMAC/JWT), and automated test suites (300+ assertions with 100% pass rates). Proven ability to convert complex operational problems into reliable, high-performance production software.",
    skillsGroups: [
      {
        title: "Frontend Engineering",
        skills: "React 19, JavaScript (ES6+), HTML5, CSS3, Tailwind CSS v4, Context API, React Hooks, Responsive UI/UX, Component Lifecycle, Vite"
      },
      {
        title: "Backend Engineering",
        skills: "Node.js (ESM), Express.js, RESTful API Architecture, JWT Authentication, Role-Based Access Control (RBAC), Middleware Chains, Error Pipelines"
      },
      {
        title: "Databases & Cloud",
        skills: "MongoDB Atlas, Mongoose ODM, Schema Architecture, Aggregation Pipelines, Indexing Optimization, Firebase Cloud Firestore, Cloudinary CDN"
      },
      {
        title: "Payments & Security",
        skills: "Paystack Gateway & Webhook LifeCycles, Escrow/Ledger Settlement, HMAC SHA-256 Signatures, VIN Masking, Opaque QR Verification"
      },
      {
        title: "Testing & DevOps",
        skills: "Vitest, Jest, Supertest (310+ Passing Automated Tests), Git, GitHub, Render, Vercel, Postman, Linux/Bash, Agile Methodologies"
      }
    ],
    // Page 1 Projects (Flagship platforms)
    page1Experience: [
      {
        title: "Tickza — Production Event Ticketing & Admission Infrastructure",
        role: "Lead Full-Stack Engineer",
        tech: "React 19, Node.js, Express.js, MongoDB Atlas, Paystack API, Tailwind CSS, Vitest",
        bullets: [
          "Architected and deployed an end-to-end event ticketing platform with multi-tier ticket sales, dynamic reservation locks, and native Paystack Naira checkout.",
          "Designed an escrow and virtual ledger settlement model with strict webhook idempotency, preventing duplicate transactions and race conditions.",
          "Engineered an offline cryptographic HMAC SHA-256 QR gate admission scanner validating passes at sub-500ms latency without database roundtrips.",
          "Authored an exhaustive automated test suite of 206 passing unit and integration assertions with a 100% pass rate, ensuring bulletproof financial safety."
        ]
      },
      {
        title: "CampusMart — Mobile-First Peer-to-Peer Student Marketplace",
        role: "Full-Stack Software Engineer",
        tech: "React 19, Node.js (v22), Express.js, MongoDB Atlas, Cloudinary CDN, Tailwind CSS",
        bullets: [
          "Developed a hyper-localized campus commerce platform connecting university students across UNILAG, YABATECH, and LASU, replacing disorganized chat-based trading.",
          "Implemented ReDoS-safe query regex escaping, student identity verification, administrative moderation queues, and one-tap WhatsApp seller handoff tracking.",
          "Wrote 104 automated tests covering catalog indexing, query sanitization, and seller lead analytics, achieving conditional production readiness."
        ]
      }
    ],
    // Page 2 Projects & Credentials
    page2Experience: [
      {
        title: "SchoolMeg — Multi-Tenant School Management ERP & Gate Scanner",
        role: "Full-Stack Software Engineer",
        tech: "Multi-Tenant MERN, QR Gate Scanner, ISO CR80 Studio, Integer-Kobo Bursary, Vitest",
        bullets: [
          "Engineered a multi-tenant school operating system isolating tenant databases with zero cross-organization leakage across 6 distinct user roles.",
          "Built printable ISO/IEC 7810 CR80 PVC student ID cards with cryptographic QR codes for instant arrival and departure gate attendance logging.",
          "Implemented Nigerian CA (40) + Exam (60) terminal result computation and an integer-kobo financial ledger eliminating decimal floating-point errors."
        ]
      },
      {
        title: "HireFit — Production AI CV Optimization Engine",
        role: "Full-Stack Engineer & AI Integrator",
        tech: "React.js, Node.js, Google Gemini API, pdf-parse, Diff Engine, Tailwind CSS",
        bullets: [
          "Developed an AI career platform that benchmarks resumes against target job descriptions with a strict Zero-Fabrication Guarantee.",
          "Engineered 6-category ATS scoring, structure-preserving resume parsing (PDF/DOCX), and an interactive before/after split-diff editing workspace."
        ]
      },
      {
        title: "SwiftExpress — Logistics & Consignment Tracking Infrastructure",
        role: "Full-Stack Software Engineer",
        tech: "React.js, Node.js, Express.js, MongoDB Atlas, Sub-Document Event Streams",
        bullets: [
          "Engineered a package tracking platform with a deterministic 13-stage consignment lifecycle finite state machine preventing invalid status transitions.",
          "Implemented sub-document timeline event streams capturing GPS checkpoints, operator timestamps, milestone remarks, and instant waybill lookup."
        ]
      },
      {
        title: "PBAT-KOH & BVDI — Civic Mobilisation & Voters Digitization Platforms",
        role: "Full-Stack & Civic Solutions Developer",
        tech: "React.js, Node.js, Express, MongoDB Atlas, ExcelJS, Cryptographic Masking",
        bullets: [
          "Built PBAT-KOH (batkohsa.org) mobilizing student volunteers across 14 tertiary institutions and 20 LGAs with captain coupon attribution and ExcelJS streaming.",
          "Engineered BVDI (bdgvotersinitiative.com) with consented voter registration, recruiter tracking, demographic telemetry, and cryptographic VIN masking across 10 official wards."
        ]
      }
    ]
  },

  // 2. FRONTEND DEVELOPER
  {
    id: "frontend",
    fileNamePrefix: "Abimbola_Samuel_Frontend_Developer_Resume",
    roleTitle: "FRONTEND SOFTWARE ENGINEER (REACT, MODERN JS & UI/UX ARCHITECTURE)",
    summary: "Detail-oriented, user-centric Frontend Software Engineer specializing in building accessible, high-performance, and responsive web applications using React (v18/19), Modern JavaScript (ES6+), and modern CSS architectures (Tailwind CSS, CSS Grid/Flexbox). Experienced in translating intricate design concepts into fluid component systems, optimizing Core Web Vitals, and implementing rich interactive interfaces—from dynamic e-commerce catalogs to real-time analytical dashboards and split-diff text editors.",
    skillsGroups: [
      {
        title: "Core Frontend & Frameworks",
        skills: "React 19, React 18, JavaScript (ES6+ / ESNext), HTML5 Semantic Markup, CSS3, Modern UI/UX Architecture"
      },
      {
        title: "Styling & Design Systems",
        skills: "Tailwind CSS v4, Modern CSS Grid & Flexbox, CSS Custom Properties (Design Tokens), Glassmorphism, Micro-Interactions"
      },
      {
        title: "State, Routing & UX",
        skills: "React Hooks (useState, useEffect, useMemo, useCallback), Context API, React Router, Custom Hooks, Form Validation"
      },
      {
        title: "Web Performance & Quality",
        skills: "Core Web Vitals Optimization, Lazy Loading & Code Splitting, Responsive Clamp Typography, Accessibility (a11y), Cross-Browser Compatibility"
      },
      {
        title: "Tooling & Integrations",
        skills: "Vite, Cloudinary Image CDN, REST API Consumption (Axios/Fetch), Git, GitHub, Browser DevTools, Responsive Mobile Testing"
      }
    ],
    page1Experience: [
      {
        title: "CampusMart — Mobile-First Peer-to-Peer Student Marketplace",
        role: "Lead Frontend Engineer",
        tech: "React 19, Tailwind CSS v4, Vite, Cloudinary CDN, Lucide Icons",
        bullets: [
          "Crafted an ultra-responsive, mobile-first marketplace interface serving university campuses with instant search, category filtering, and location badges.",
          "Built dynamic item listing workflows with client-side image compression, real-time preview, and automated Cloudinary CDN transformations.",
          "Designed accessible interactive modals, sliding navigation drawers, and one-tap WhatsApp seller handoff triggers maximizing conversion."
        ]
      },
      {
        title: "TOMÉA — Luxury Artisanal Fragrance E-Commerce Boutique",
        role: "Frontend UI Architect",
        tech: "Semantic HTML5, CSS3 Custom Properties, Vanilla ES6+, Glassmorphism, Firebase",
        bullets: [
          "Designed and built a cinematic, editorial e-commerce frontend delivering a sub-0.8s mobile time-to-interactive with zero layout shifts.",
          "Engineered an interactive Olfactory Pyramid studio dynamically decomposing fragrances into Top, Heart, and Base notes with longevity gauges.",
          "Built a glassmorphic sticky header, smooth IntersectionObserver scroll reveals, and a luxury concierge checkout pipeline directly into WhatsApp."
        ]
      }
    ],
    page2Experience: [
      {
        title: "HireFit — AI CV Diagnostic & Before/After Diff Interface",
        role: "Frontend Engineer",
        tech: "React.js, Tailwind CSS, Interactive Diff Engine, Score Gauges, Vite",
        bullets: [
          "Developed an interactive career diagnostic dashboard displaying 6-category visual ATS score meters, keyword gap indicators, and status chips.",
          "Engineered a split-view before/after diff editor allowing job seekers to review, accept, or reject AI-tailored bullet points with zero layout instability.",
          "Implemented drag-and-drop file upload zones with animated progress states, error boundary fallbacks, and real-time validation previews."
        ]
      },
      {
        title: "Tickza — Live Event Discovery & Admission Experience",
        role: "Frontend Engineer",
        tech: "React 19, Tailwind CSS, Dynamic QR Pass Rendering, Lucide React",
        bullets: [
          "Designed a mobile-first event discovery storefront featuring live date pickers, venue filters, and interactive multi-tier ticket quantity selectors.",
          "Created animated digital ticket passes with real-time admission verification indicators, status cards, and seamless mobile responsive layouts."
        ]
      },
      {
        title: "DX Sportz — Sports Media Hub & Match Center Interface",
        role: "Frontend Developer",
        tech: "React (Vite), CSS3 Modules, Responsive Grid/Flexbox, Multimedia Feeds",
        bullets: [
          "Built a live domestic NPFL match center with fixture schedules, live league standings tables, and multimedia sports journalism editorial layouts.",
          "Implemented responsive video and photo galleries with lazy loading, fluid image aspect ratios, and cross-browser touch controls."
        ]
      },
      {
        title: "SwiftExpress — Logistics & Consignment Tracking Dashboard",
        role: "Frontend Developer",
        tech: "React.js, CSS3 Custom Properties, Timeline Milestones, Real-Time Feedback",
        bullets: [
          "Engineered a public package tracking interface rendering live multi-checkpoint milestone timelines and status indicator badges.",
          "Created an administrative dispatch view allowing courier managers to log checkpoint transitions with instant visual validation."
        ]
      }
    ]
  },

  // 3. BACKEND DEVELOPER
  {
    id: "backend",
    fileNamePrefix: "Abimbola_Samuel_Backend_Developer_Resume",
    roleTitle: "BACKEND SOFTWARE ENGINEER (NODE.JS, REST APIS & DATABASE INFRASTRUCTURE)",
    summary: "Systematic, security-focused Backend Software Engineer specializing in scalable RESTful API design, relational and document database modeling, and transaction-critical server logic. Expert in Node.js, Express.js, MongoDB Atlas (Mongoose ODM), Firebase Firestore, and automated API testing (Vitest, Jest, Supertest). Experienced in building webhook life cycles (Paystack), cryptographic tokens (HMAC SHA-256 / JWT), multi-tenant data isolation, rate limiting, and memory-efficient streaming exports.",
    skillsGroups: [
      {
        title: "Server & Runtime Environments",
        skills: "Node.js (ESM / CommonJS), Express.js, RESTful API Architecture, Middleware Pipelines, Asynchronous Event Loops, Error Handling"
      },
      {
        title: "Databases & Data Modeling",
        skills: "MongoDB Atlas, Mongoose ODM, Schema Architecture, Complex Aggregations, Indexing & Query Optimization, Firebase Firestore"
      },
      {
        title: "Security & Cryptography",
        skills: "HMAC SHA-256 Signatures, JWT Authentication, Role-Based Access Control (RBAC), VIN Masking, Input Sanitization (NoSQL Injection Guards)"
      },
      {
        title: "Payments & Financial Engineering",
        skills: "Paystack Payment Gateway, Webhook Handlers, Idempotency Guards, Escrow/Virtual Ledgers, Integer-Kobo Financial Accounting"
      },
      {
        title: "Testing, DevOps & Tooling",
        skills: "Vitest, Jest, Supertest (310+ Passing Automated Tests), Git, GitHub, Render, Vercel, Postman, Thunder Client, Linux CLI, ExcelJS"
      }
    ],
    page1Experience: [
      {
        title: "Tickza — Production Event Ticketing & Financial Ledger Backend",
        role: "Lead Backend Engineer",
        tech: "Node.js (ESM), Express.js, MongoDB Atlas, Paystack API, HMAC SHA-256, Vitest",
        bullets: [
          "Engineered a high-concurrency Node.js/Express backend managing multi-tier ticket reservations, Paystack payment sessions, and webhook lifecycles.",
          "Built an escrow and virtual ledger settlement model with strict idempotency keys preventing double-spending and duplicate ticket allocations.",
          "Designed an offline cryptographic admission verification engine utilizing dynamic-salt HMAC SHA-256 signatures for sub-second gate validation.",
          "Wrote 206 automated unit and integration tests achieving a 100% pass rate across ticket quotas, webhook processing, and door check-in states."
        ]
      },
      {
        title: "SchoolMeg — Multi-Tenant School Management ERP Backend",
        role: "Backend Architect",
        tech: "Multi-Tenant Node/Express, MongoDB Mongoose, Cryptographic QR Tokens, Vitest",
        bullets: [
          "Architected strict multi-tenant database isolation patterns ensuring zero cross-tenant query leakage across 6 administrative user roles.",
          "Engineered an integer-kobo bursary accounting engine storing all currency values as discrete integers, eliminating floating-point rounding errors.",
          "Designed opaque cryptographic QR attendance verification tokens with time-stamped gate logs, tamper-evident audit trails, and automated grading."
        ]
      }
    ],
    page2Experience: [
      {
        title: "SwiftExpress — Logistics & Consignment Tracking State Machine",
        role: "Backend Engineer",
        tech: "Node.js, Express.js, MongoDB Atlas, Sub-Document Event Streams, JWT Auth",
        bullets: [
          "Developed a deterministic 13-stage consignment lifecycle state machine with strict transition guards preventing invalid tracking leaps.",
          "Implemented sub-document timeline event streams capturing GPS checkpoints, operator timestamps, milestone remarks, and courier dispatch logs.",
          "Built secure JWT-authenticated dispatch management APIs with granular permission enforcement and high-speed public waybill indexing."
        ]
      },
      {
        title: "PBAT-KOH — Campus Mobilisation & Referral Attribution Backend",
        role: "Backend Engineer",
        tech: "Node.js, Express.js, MongoDB Atlas, ExcelJS Streaming, Render",
        bullets: [
          "Architected high-throughput registration endpoints coordinating volunteer intake across 14 tertiary campuses and 20 Lagos LGAs.",
          "Implemented referral coupon attribution logic, duplicate detection algorithms, and ExcelJS streaming memory-efficient report exports."
        ]
      },
      {
        title: "BVDI — Civic Data Governance & Audit Log Stream Engine",
        role: "Backend & Data Privacy Engineer",
        tech: "Node.js, Express.js, MongoDB Atlas, SHA-256 Masking, Audit Logs",
        bullets: [
          "Built secure voter data ingestion pipelines featuring cryptographic VIN hashing and client-side masking to safeguard citizen privacy.",
          "Designed an immutable administrative audit log stream tracking all record access, data modifications, and recruiter performance metrics."
        ]
      },
      {
        title: "HireFit — AI Integration & Resume Parsing Pipeline",
        role: "Backend Developer",
        tech: "Node.js, Express, Google Gemini REST API, pdf-parse, Structured JSON",
        bullets: [
          "Engineered server-side PDF and Word document extraction pipelines converting raw CV text into structured candidate skill schemas.",
          "Integrated Google Gemini LLM with system prompt guardrails enforcing strict zero-fabrication of applicant credentials."
        ]
      }
    ]
  }
];

// ==========================================
// 1. DOCX GENERATION FUNCTION
// ==========================================

async function generateDocx(data) {
  const primaryColor = "0F172A"; // Slate 900
  const accentColor = "0284C7";  // Tech blue
  const mutedColor = "475569";   // Slate 600

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Arial",
            size: 19 // 9.5pt
          }
        }
      }
    },
    sections: [
      // SECTION 1: PAGE 1
      {
        properties: {
          page: {
            margin: { top: 600, right: 600, bottom: 600, left: 600 }
          }
        },
        children: [
          // Header Name
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 40 },
            children: [
              new TextRun({
                text: contactInfo.name,
                bold: true,
                size: 30, // 15pt
                color: primaryColor,
                font: "Arial"
              })
            ]
          }),
          // Role Title
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: data.roleTitle,
                bold: true,
                size: 21, // 10.5pt
                color: accentColor,
                font: "Arial"
              })
            ]
          }),
          // Contact Info Bar
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 140 },
            children: [
              new TextRun({ text: `${contactInfo.phone}  •  ${contactInfo.email}  •  ${contactInfo.location}\n`, size: 17, color: mutedColor }),
              new TextRun({ text: `GitHub: github.com/Psalmstrings  •  Portfolio: psalmstrings.vercel.app`, size: 17, color: accentColor })
            ]
          }),

          // PROFESSIONAL SUMMARY
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 120, after: 60 },
            border: {
              bottom: { color: accentColor, space: 3, value: BorderStyle.SINGLE, size: 8 }
            },
            children: [
              new TextRun({ text: "PROFESSIONAL SUMMARY", bold: true, size: 20, color: primaryColor })
            ]
          }),
          new Paragraph({
            spacing: { after: 120 },
            children: [
              new TextRun({ text: data.summary, size: 18, color: "1F2937" })
            ]
          }),

          // TECHNICAL SKILLS
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 120, after: 60 },
            border: {
              bottom: { color: accentColor, space: 3, value: BorderStyle.SINGLE, size: 8 }
            },
            children: [
              new TextRun({ text: "TECHNICAL & ENGINEERING SKILLS", bold: true, size: 20, color: primaryColor })
            ]
          }),
          ...data.skillsGroups.map(group => 
            new Paragraph({
              spacing: { after: 40 },
              bullet: { level: 0 },
              children: [
                new TextRun({ text: `${group.title}: `, bold: true, size: 18, color: primaryColor }),
                new TextRun({ text: group.skills, size: 18, color: "334155" })
              ]
            })
          ),

          // KEY PROJECTS (PAGE 1)
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 140, after: 60 },
            border: {
              bottom: { color: accentColor, space: 3, value: BorderStyle.SINGLE, size: 8 }
            },
            children: [
              new TextRun({ text: "FEATURED SOFTWARE ENGINEERING PROJECTS", bold: true, size: 20, color: primaryColor })
            ]
          }),
          ...data.page1Experience.flatMap(exp => [
            new Paragraph({
              spacing: { before: 90, after: 30 },
              children: [
                new TextRun({ text: exp.title, bold: true, size: 19, color: primaryColor }),
                new TextRun({ text: `  —  ${exp.role}`, bold: true, size: 18, color: accentColor })
              ]
            }),
            new Paragraph({
              spacing: { after: 40 },
              children: [
                new TextRun({ text: "Core Technologies: ", bold: true, size: 16, color: mutedColor }),
                new TextRun({ text: exp.tech, italics: true, size: 16, color: mutedColor })
              ]
            }),
            ...exp.bullets.map(bullet => 
              new Paragraph({
                spacing: { after: 30 },
                bullet: { level: 0 },
                children: [
                  new TextRun({ text: bullet, size: 17, color: "1F2937" })
                ]
              })
            )
          ])
        ]
      },

      // SECTION 2: PAGE 2
      {
        properties: {
          page: {
            margin: { top: 600, right: 600, bottom: 600, left: 600 }
          }
        },
        children: [
          // PROJECTS CONTINUED (PAGE 2)
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 40, after: 60 },
            border: {
              bottom: { color: accentColor, space: 3, value: BorderStyle.SINGLE, size: 8 }
            },
            children: [
              new TextRun({ text: "PROJECTS & SYSTEM IMPLEMENTATIONS (CONTINUED)", bold: true, size: 20, color: primaryColor })
            ]
          }),
          ...data.page2Experience.flatMap(exp => [
            new Paragraph({
              spacing: { before: 90, after: 30 },
              children: [
                new TextRun({ text: exp.title, bold: true, size: 19, color: primaryColor }),
                new TextRun({ text: `  —  ${exp.role}`, bold: true, size: 18, color: accentColor })
              ]
            }),
            new Paragraph({
              spacing: { after: 40 },
              children: [
                new TextRun({ text: "Core Technologies: ", bold: true, size: 16, color: mutedColor }),
                new TextRun({ text: exp.tech, italics: true, size: 16, color: mutedColor })
              ]
            }),
            ...exp.bullets.map(bullet => 
              new Paragraph({
                spacing: { after: 30 },
                bullet: { level: 0 },
                children: [
                  new TextRun({ text: bullet, size: 17, color: "1F2937" })
                ]
              })
            )
          ]),

          // EDUCATION & CREDENTIALS
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 140, after: 60 },
            border: {
              bottom: { color: accentColor, space: 3, value: BorderStyle.SINGLE, size: 8 }
            },
            children: [
              new TextRun({ text: "EDUCATION & CREDENTIALS", bold: true, size: 20, color: primaryColor })
            ]
          }),
          ...contactInfo.education.map(edu => 
            new Paragraph({
              spacing: { after: 50 },
              children: [
                new TextRun({ text: `${edu.school} — `, bold: true, size: 18, color: primaryColor }),
                new TextRun({ text: edu.credential, size: 18, color: "1F2937" }),
                new TextRun({ text: ` (${edu.year})`, italics: true, size: 17, color: mutedColor })
              ]
            })
          ),

          // LANGUAGES & METHODOLOGIES
          new Paragraph({
            spacing: { before: 80, after: 60 },
            children: [
              new TextRun({ text: "Languages: ", bold: true, size: 18, color: primaryColor }),
              new TextRun({ text: contactInfo.languages, size: 18, color: "334155" })
            ]
          }),
          new Paragraph({
            spacing: { after: 40 },
            children: [
              new TextRun({ text: "Engineering Practices: ", bold: true, size: 18, color: primaryColor }),
              new TextRun({ text: "Test-Driven Development (TDD), Agile / Scrum, Code Reviews, CI/CD, Documentation Integrity", size: 18, color: "334155" })
            ]
          })
        ]
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  const outPath = path.join(resumesDir, `${data.fileNamePrefix}.docx`);
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Generated DOCX: ${data.fileNamePrefix}.docx`);
}

// ==========================================
// 2. HTML TO PDF GENERATION FUNCTION
// ==========================================

function generateHtml(data) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>${data.roleTitle} - ${contactInfo.name}</title>
  <style>
    @page {
      size: A4;
      margin: 12mm 15mm 12mm 15mm;
    }
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      color: #1f2937;
      background: #ffffff;
      line-height: 1.42;
      font-size: 9.3pt;
    }
    .page-break {
      page-break-before: always;
      padding-top: 4px;
    }
    .header {
      text-align: center;
      border-bottom: 2px solid #0284c7;
      padding-bottom: 7px;
      margin-bottom: 11px;
    }
    .name {
      font-size: 18pt;
      font-weight: 800;
      color: #0f172a;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      margin-bottom: 2px;
    }
    .role-title {
      font-size: 10.5pt;
      font-weight: 700;
      color: #0284c7;
      letter-spacing: 0.03em;
      margin-bottom: 5px;
    }
    .contact-bar {
      font-size: 8.5pt;
      color: #4b5563;
      display: flex;
      justify-content: center;
      gap: 10px;
      flex-wrap: wrap;
    }
    .contact-bar a {
      color: #0284c7;
      text-decoration: none;
    }
    .section {
      margin-bottom: 11px;
    }
    .section-title {
      font-size: 10pt;
      font-weight: 700;
      color: #0f172a;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      border-bottom: 1px solid #cbd5e1;
      padding-bottom: 2px;
      margin-bottom: 6px;
    }
    .summary-text {
      font-size: 9pt;
      color: #334155;
      text-align: justify;
      line-height: 1.42;
    }
    .skills-list {
      list-style: none;
    }
    .skills-item {
      font-size: 8.8pt;
      margin-bottom: 3px;
      line-height: 1.38;
    }
    .skills-label {
      font-weight: 700;
      color: #0f172a;
    }
    .skills-val {
      color: #334155;
    }
    .exp-item {
      margin-bottom: 9px;
    }
    .exp-header {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      margin-bottom: 1px;
    }
    .exp-title {
      font-size: 9.5pt;
      font-weight: 700;
      color: #0f172a;
    }
    .exp-role {
      font-size: 8.8pt;
      font-weight: 600;
      color: #0284c7;
    }
    .exp-tech {
      font-size: 8pt;
      color: #64748b;
      font-style: italic;
      margin-bottom: 3px;
    }
    .exp-bullets {
      padding-left: 15px;
      margin: 0;
    }
    .exp-bullets li {
      font-size: 8.6pt;
      color: #334155;
      margin-bottom: 2.2px;
      line-height: 1.34;
    }
    .edu-item {
      display: flex;
      justify-content: space-between;
      font-size: 8.8pt;
      margin-bottom: 3px;
    }
    .edu-school {
      font-weight: 700;
      color: #0f172a;
    }
    .edu-year {
      color: #64748b;
      font-style: italic;
    }
  </style>
</head>
<body>
  <!-- ==================== PAGE 1 ==================== -->
  <div class="header">
    <div class="name">${contactInfo.name}</div>
    <div class="role-title">${data.roleTitle}</div>
    <div class="contact-bar">
      <span>${contactInfo.phone}</span>
      <span>•</span>
      <span>${contactInfo.email}</span>
      <span>•</span>
      <span>${contactInfo.location}</span>
      <span>•</span>
      <span><a href="${contactInfo.github}">github.com/Psalmstrings</a></span>
      <span>•</span>
      <span><a href="${contactInfo.portfolio}">psalmstrings.vercel.app</a></span>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Professional Summary</div>
    <div class="summary-text">${data.summary}</div>
  </div>

  <div class="section">
    <div class="section-title">Technical Skills &amp; Engineering Proficiencies</div>
    <ul class="skills-list">
      ${data.skillsGroups.map(g => `
        <li class="skills-item">
          <span class="skills-label">${g.title}:</span>
          <span class="skills-val">${g.skills}</span>
        </li>
      `).join('')}
    </ul>
  </div>

  <div class="section">
    <div class="section-title">Featured Software Engineering Projects</div>
    ${data.page1Experience.map(e => `
      <div class="exp-item">
        <div class="exp-header">
          <span class="exp-title">${e.title}</span>
          <span class="exp-role">${e.role}</span>
        </div>
        <div class="exp-tech">Technologies: ${e.tech}</div>
        <ul class="exp-bullets">
          ${e.bullets.map(b => `<li>${b}</li>`).join('')}
        </ul>
      </div>
    `).join('')}
  </div>

  <!-- ==================== PAGE 2 ==================== -->
  <div class="page-break">
    <div class="section">
      <div class="section-title">Projects &amp; System Implementations (Continued)</div>
      ${data.page2Experience.map(e => `
        <div class="exp-item">
          <div class="exp-header">
            <span class="exp-title">${e.title}</span>
            <span class="exp-role">${e.role}</span>
          </div>
          <div class="exp-tech">Technologies: ${e.tech}</div>
          <ul class="exp-bullets">
            ${e.bullets.map(b => `<li>${b}</li>`).join('')}
          </ul>
        </div>
      `).join('')}
    </div>

    <div class="section">
      <div class="section-title">Education &amp; Credentials</div>
      ${contactInfo.education.map(ed => `
        <div class="edu-item">
          <span><strong class="edu-school">${ed.school}</strong> — ${ed.credential}</span>
          <span class="edu-year">${ed.year}</span>
        </div>
      `).join('')}
      <div style="font-size: 8.6pt; color: #475569; margin-top: 4px;">
        <strong>Languages:</strong> ${contactInfo.languages}
      </div>
      <div style="font-size: 8.6pt; color: #475569; margin-top: 2px;">
        <strong>Engineering Practices:</strong> Test-Driven Development (TDD), Agile / Scrum, Code Reviews, Continuous Integration, Strict Documentation Integrity.
      </div>
    </div>
  </div>
</body>
</html>`;
}

async function run() {
  const edgePath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";

  for (const data of resumesData) {
    // 1. Generate DOCX
    await generateDocx(data);

    // 2. Generate HTML & PDF
    const htmlContent = generateHtml(data);
    const htmlFile = path.join(resumesDir, `${data.fileNamePrefix}.html`);
    const pdfFile = path.join(resumesDir, `${data.fileNamePrefix}.pdf`);

    fs.writeFileSync(htmlFile, htmlContent, 'utf-8');

    try {
      execSync(`"${edgePath}" --headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf="${pdfFile}" "${htmlFile}"`, { stdio: 'pipe' });
      console.log(`✓ Generated PDF:  ${data.fileNamePrefix}.pdf`);
    } catch (err) {
      console.error(`Error generating PDF for ${data.id}:`, err.message);
    }
  }

  console.log("\nAll 3 Resumes (DOCX and PDF) regenerated successfully with 2 balanced pages!");
}

run();
