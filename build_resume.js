const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, TabStopType, TabStopPosition, LevelFormat, ExternalHyperlink
} = require('docx');
const fs = require('fs');

const NAVY  = "0F1F4B";
const GOLD  = "C8993A";
const GRAY  = "555555";
const LGRAY = "888888";

// ── Helpers ──────────────────────────────────────────────────────────

function gap(pts = 6) {
  return new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "", size: pts * 2 })] });
}

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 180, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 3 } },
    children: [new TextRun({
      text,
      bold: true,
      font: "Arial",
      size: 18,
      color: NAVY,
      characterSpacing: 60,
      allCaps: true,
    })]
  });
}

function jobTitle(title, company, location, dates) {
  return new Paragraph({
    spacing: { before: 140, after: 0 },
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    children: [
      new TextRun({ text: title, bold: true, font: "Arial", size: 22, color: NAVY }),
      new TextRun({ text: "  |  ", font: "Arial", size: 20, color: LGRAY }),
      new TextRun({ text: company, bold: true, font: "Arial", size: 20, color: NAVY }),
      new TextRun({ text: ", " + location, font: "Arial", size: 20, color: GRAY }),
      new TextRun({ text: "\t", font: "Arial", size: 20 }),
      new TextRun({ text: dates, font: "Arial", size: 20, color: LGRAY, italics: true }),
    ]
  });
}

function jobSubtitle(text) {
  return new Paragraph({
    spacing: { before: 30, after: 60 },
    children: [new TextRun({ text, italics: true, font: "Arial", size: 20, color: GRAY })]
  });
}

function bullet(runs) {
  const children = typeof runs === 'string'
    ? [new TextRun({ text: runs, font: "Arial", size: 20, color: GRAY })]
    : runs;
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 20, after: 20 },
    children,
  });
}

function boldRun(text) {
  return new TextRun({ text, bold: true, font: "Arial", size: 20, color: NAVY });
}

function bodyRun(text) {
  return new TextRun({ text, font: "Arial", size: 20, color: GRAY });
}

function inlineRow(items) {
  // items: array of strings, joined with " | "
  const children = [];
  items.forEach((item, i) => {
    children.push(new TextRun({ text: item, font: "Arial", size: 20, color: GRAY }));
    if (i < items.length - 1) {
      children.push(new TextRun({ text: "  |  ", font: "Arial", size: 20, color: LGRAY }));
    }
  });
  return new Paragraph({ spacing: { before: 30, after: 30 }, children });
}

function miniJobLine(text) {
  return new Paragraph({
    spacing: { before: 30, after: 0 },
    children: [new TextRun({ text, font: "Arial", size: 18, color: LGRAY, italics: true })]
  });
}

// ── Document ─────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "•",
        alignment: AlignmentType.LEFT,
        style: {
          paragraph: {
            indent: { left: 360, hanging: 200 },
            spacing: { before: 20, after: 20 },
          },
          run: { font: "Arial", size: 20, color: GOLD }
        }
      }]
    }]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 20 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 900, right: 1080, bottom: 900, left: 1080 }
      }
    },
    children: [

      // ── NAME ──────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({
          text: "JEREL SMITH",
          bold: true,
          font: "Arial",
          size: 52,
          color: NAVY,
          characterSpacing: 80,
        })]
      }),

      // ── SUBTITLE ──────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({
          text: "Senior Product Manager  |  AI-Powered Products  |  Go-to-Market & Product Marketing",
          italics: true,
          font: "Arial",
          size: 20,
          color: GRAY,
        })]
      }),

      // ── CONTACT ───────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "E0E4EE", space: 6 } },
        children: [
          new TextRun({ text: "ContactJerel@pm.me", font: "Arial", size: 18, color: LGRAY }),
          new TextRun({ text: "  •  ", font: "Arial", size: 18, color: GOLD }),
          new TextRun({ text: "917-837-1562", font: "Arial", size: 18, color: LGRAY }),
          new TextRun({ text: "  •  ", font: "Arial", size: 18, color: GOLD }),
          new TextRun({ text: "Brooklyn, NY", font: "Arial", size: 18, color: LGRAY }),
        ]
      }),

      gap(6),

      // ── PROFESSIONAL SUMMARY ───────────────────────────────────
      sectionHeader("Professional Summary"),

      new Paragraph({
        spacing: { before: 80, after: 0 },
        children: [new TextRun({
          text: "I'm a Senior PM with 10+ years experience across e-commerce, B2B SaaS, gaming, and healthcare. Rare blend of product strategy and product marketing fluency — hands-on builder with working knowledge of LLM integration, hybrid search, and generative AI. Known for translating complex AI capabilities into customer narratives that drive adoption and revenue.",
          font: "Arial",
          size: 20,
          color: GRAY,
        })]
      }),

      gap(4),

      // ── CORE COMPETENCIES ─────────────────────────────────────
      sectionHeader("Core Competencies"),

      inlineRow([
        "AI Product Strategy",
        "Go-to-Market Planning & Execution",
        "Product Positioning & Messaging",
        "Roadmap & Backlog Management",
        "A/B Testing & Experimentation",
        "Customer Adoption & Retention",
        "Competitive Intelligence",
        "Search & Personalization",
        "Cross-functional Leadership",
        "B2B & B2C Commerce",
      ]),

      gap(4),

      // ── AI PROJECTS ───────────────────────────────────────────
      sectionHeader("AI Projects & Technical Work"),

      new Paragraph({
        spacing: { before: 100, after: 20 },
        children: [
          new TextRun({ text: "Hybrid Search Demo App", bold: true, font: "Arial", size: 21, color: NAVY }),
          new TextRun({ text: "  (demo.AiPMBuilder.com)  —  Self-Hosted Homelab", font: "Arial", size: 20, color: GRAY }),
        ]
      }),

      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [new TextRun({
          text: "Built a production-grade, multi-vertical AI-powered search application from scratch, demonstrating hands-on AI product capability well beyond typical PM scope.",
          font: "Arial", size: 20, color: GRAY,
        })]
      }),

      bullet([boldRun("SOLR 9 + Ollama"), bodyRun(" (nomic-embed-text) for hybrid keyword + vector search across six industry verticals: cameras, legal, healthcare, recruiting, SaaS, and weddings.")]),
      bullet([boldRun("LLM-generated summaries"), bodyRun(" via qwen2.5 with real-time SSE streaming, hybrid relevance scoring, and a timing breakdown UI — a working RAG-style architecture demo.")]),
      bullet([boldRun("Full Flask web interface"), bodyRun(" deployed on a self-hosted ThinkPad server, with AiPM Builder branding and mobile-responsive design, used for stakeholder demonstrations.")]),

      gap(4),

      // ── PROFESSIONAL EXPERIENCE ───────────────────────────────
      sectionHeader("Professional Experience"),

      // B&H — B2B & Used (most recent)
      jobTitle("Senior Product Manager — B2B & Used", "B&H Photo Video", "New York, NY", "Feb 2026 – Present"),
      jobSubtitle("Promoted to lead B2B and Used department web experience — owning end-to-end strategy for condition integrity, B2B portal, and education/government verticals."),

      bullet([boldRun("Led complete overhaul of condition rating system"), bodyRun(" presented to COO — projected 8.5% CSAT improvement, closing a longstanding gap between listed condition and customer expectation.")]),
      bullet([boldRun("Audited the full Sell Used funnel"), bodyRun(" and built an image upload prototype targeting expectation mismatch between listed condition and received product.")]),
      bullet([boldRun("Led EDU Advantage landing page redesign"), bodyRun(" after identifying dead-click patterns via FullStory and GA; drove Figma handoff improving benefit clarity, sign-up flow, and editorial control over curated content.")]),
      bullet([boldRun("Defined content architecture for a dedicated K-12 School Safety landing page"), bodyRun(" — five product pillars, compliance resource layer, and territory-routed lead capture funnel — producing three wireframe directions for VP-level stakeholder review.")]),
      bullet([boldRun("Built a functional CMS proof of concept"), bodyRun(" demonstrating editor-controlled content management for B2B pages — a decision artifact designed to move the build-vs-buy conversation from abstract to tangible.")]),

      // B&H — Search & Navigation
      jobTitle("Senior Product Manager — Search & Navigation", "B&H Photo Video", "New York, NY", "Sep 2022 – Feb 2026"),
      jobSubtitle("Lead PM for Search and Navigation on bhphotovideo.com — owning roadmap strategy, cross-functional execution, and go-to-market alignment across a high-traffic e-commerce platform doing hundreds of millions in transactions."),

      bullet([boldRun("Led full product lifecycle for AI-enhanced search features"), bodyRun(" including My Camera accessory finder (+5.2% attachment rate, +$38 AOV) — from discovery and positioning through launch and post-launch measurement.")]),
      bullet([boldRun("Owned search and navigation strategy:"), bodyRun(" sort logic optimization and zero-result reduction yielded a 3.7% conversion lift and 2.3% bounce rate reduction; Kit PDP interactive redesign increased kit sales 4.3%.")]),
      bullet([boldRun("Established a 30-day product health framework"), bodyRun(" tracking funnel completion, engagement, and conversion; grew A/B test velocity 18% YoY; embedded live chat and consolidated promotional pages into a revenue-generating hub.")]),

      // Net32
      jobTitle("Senior Product Manager", "Net32", "Cary, NC", "2021 – 2022"),
      jobSubtitle("Led product and growth initiatives for a B2B dental supply marketplace, with emphasis on lifecycle messaging, adoption campaigns, and customer segmentation."),

      bullet([boldRun("Intelligent recommendation engine"), bodyRun(" for cart and post-add-to-cart experiences, growing average cart size 5.2% and customer lifetime value 4.7%.")]),
      bullet([boldRun("CRO initiatives"), bodyRun(" targeting 25% AOV increase — acquiring 600+ new dental practice buyers and reactivating dormant accounts through re-engagement strategies.")]),
      bullet([boldRun("Checkout flow redesign"), bodyRun(" based on user pain point research, boosting conversion 7% in Q3 2021; drove a 3.5% repeat purchase rate increase through behavioral insights from Heap Analytics.")]),

      // HCA
      jobTitle("Product Manager", "HCA HealthTrust", "Nashville, TN", "2017 – 2021"),
      jobSubtitle("Managed product catalog, contract lifecycle, and roadmap for a Group Purchasing Operations suite processing $4.4B in annual transactions."),

      bullet([bodyRun("Drove feature adoption through ROI-focused stakeholder communication and internal launch campaigns, contributing to "), boldRun("12–14% annual portfolio growth"), bodyRun(".")]),
      bullet([bodyRun("Evaluated and launched an internal "), boldRun("PPE Marketplace on Shopify Plus"), bodyRun(" — managed end-to-end positioning, stakeholder messaging, and rollout communications.")]),
      bullet([bodyRun("Led cross-functional workshops to surface customer pain points and align teams around product narratives grounded in qualitative and quantitative research.")]),

      // VGT
      jobTitle("Product Owner", "Video Game Technologies", "Cool Springs, TN", "2015 – 2017"),

      bullet([bodyRun("Managed Quote-to-Cash solution delivery pipeline; built executive-level communication on scope, impact, and risk for VP-level stakeholders.")]),
      bullet([bodyRun("Led UAT for "), boldRun("Dynamics 365 CRM"), bodyRun(" across sales, manufacturing, field service, and inventory management teams.")]),

      // TEGNA
      jobTitle("Web Development Dept. Manager", "TEGNA", "Nashville, TN", "2013 – 2015"),

      bullet([bodyRun("Led team of 15 developers across multiple client projects; designed IA and user flows for mobile-responsive websites.")]),
      bullet([bodyRun("Evaluated vendors and prepared RFP responses for website platforms; facilitated client requirements reviews.")]),

      gap(4),

      // ── EARLIER ROLES ─────────────────────────────────────────
      new Paragraph({
        spacing: { before: 60, after: 0 },
        children: [new TextRun({ text: "Earlier roles:", bold: true, font: "Arial", size: 18, color: NAVY })]
      }),
      miniJobLine("Dickson Resources — FE Developer (2011–2013)  ·  R&R Studios — FE Developer (2008–2011)  ·  Bloomberg LP — API Specialist (2005–2008)  ·  Continuum Health Partners — Managed Care Liaison (2003–2005)"),

      gap(4),

      // ── EDUCATION ─────────────────────────────────────────────
      sectionHeader("Education"),

      new Paragraph({
        spacing: { before: 80, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: "Bachelor’s Degree: Human & Organizational Development / Computer Science", bold: true, font: "Arial", size: 20, color: NAVY }),
          new TextRun({ text: "\t", font: "Arial", size: 20 }),
          new TextRun({ text: "May 2003", font: "Arial", size: 20, color: LGRAY, italics: true }),
        ]
      }),
      new Paragraph({
        spacing: { before: 20, after: 0 },
        children: [new TextRun({ text: "Vanderbilt University", italics: true, font: "Arial", size: 20, color: GRAY })]
      }),

      gap(4),

      // ── CERTIFICATIONS ────────────────────────────────────────
      sectionHeader("Certifications"),

      inlineRow(["SAFe Certified", "Certified Product Owner", "Certified Scrum Master"]),

      gap(4),

      // ── TECHNICAL FLUENCY ─────────────────────────────────────
      sectionHeader("AI & Technical Fluency"),

      inlineRow([
        "Generative AI (Ollama, LLM integration, RAG pipelines)",
        "SOLR 9 & Hybrid Search Architecture",
        "Python & Flask",
      ]),
      inlineRow([
        "Linux",
        "A/B Testing (Google Optimize, CXL)",
        "Heap Analytics & Google Analytics",
        "Jira, Asana, Monday",
        "Figma",
      ]),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/sessions/intelligent-gallant-goldberg/mnt/outputs/Jerel_Smith_Resume.docx', buf);
  console.log('Done.');
});
