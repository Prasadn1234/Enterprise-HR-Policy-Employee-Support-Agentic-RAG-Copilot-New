"""Build FDE roles and responsibilities document from the resume."""
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:\Users\LENOVO\OneDrive\Desktop\FDE Roles & Responsibilities - Nimmalapudi Prasad.docx"

doc = Document()
section = doc.sections[0]
section.top_margin = Inches(0.75)
section.bottom_margin = Inches(0.75)
section.left_margin = Inches(0.85)
section.right_margin = Inches(0.85)

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.space_before = Pt(0)


def shade_cell(cell, hex_color):
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def add_heading_styled(text, size=16):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    run.font.name = "Calibri"
    return p


def add_body(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.name = "Calibri"
    return p


def add_bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.clear()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(1)
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.name = "Calibri"
    return p


title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.LEFT
title.paragraph_format.space_after = Pt(0)
r = title.add_run("Forward Deployed Engineer (FDE)")
r.bold = True
r.font.size = Pt(22)
r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
r.font.name = "Calibri"

sub = doc.add_paragraph()
sub.paragraph_format.space_before = Pt(0)
sub.paragraph_format.space_after = Pt(2)
r = sub.add_run("Roles and Responsibilities")
r.bold = True
r.font.size = Pt(14)
r.font.color.rgb = RGBColor(0x2E, 0x5A, 0x88)
r.font.name = "Calibri"

who = doc.add_paragraph()
who.paragraph_format.space_after = Pt(2)
r = who.add_run("Nimmalapudi Prasad")
r.bold = True
r.font.size = Pt(12)
r.font.name = "Calibri"

role = doc.add_paragraph()
role.paragraph_format.space_after = Pt(8)
r = role.add_run("AI Solution Architect  |  Forward Deployed Engineer")
r.italic = True
r.font.size = Pt(11)
r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
r.font.name = "Calibri"

add_body(
    "Prepared from the professional record in Prasad Nimmalapudi’s FDE / AI Solution Architect resume. "
    "The statements below describe the Forward Deployed Engineer role as practiced in enterprise customer "
    "engagements, principally at Ford Technology Services India (August 2012 – June 2026), and the capabilities "
    "that role draws on."
)

add_heading_styled("Role purpose", 14)
add_body(
    "A Forward Deployed Engineer works inside the customer’s problem, not only on a backlog. The role designs, "
    "builds, customizes, and operates production Generative AI and Agentic AI systems with the customer’s "
    "stakeholders, then leaves behind an architecture the customer’s teams can run. The work spans manufacturing "
    "and banking, and joins solution architecture with hands-on platform engineering, evaluation, and delivery leadership."
)

add_heading_styled("1. Embed with the customer and own the outcome", 14)
add_body("Sit with enterprise customers and turn business objectives into systems that run in production.")
for item in [
    "Partner directly with customer stakeholders to understand objectives, constraints, and what “done” means for the business.",
    "Deploy and customize AI/ML solutions in the customer environment, including large language models, retrieval-augmented generation (RAG) pipelines, and vector databases.",
    "Fine-tune models, optimize prompts, and set deployment approaches that the customer can scale after the engagement.",
    "Translate ambiguous business requirements into production-ready designs for agentic pipelines, RAG, and LLM integrations.",
    "Stay accountable for adoption: the system has to fit the customer’s workflow, data, and operating model, not only a demo.",
]:
    add_bullet(item)

add_heading_styled("2. Architect agentic and generative AI solutions", 14)
add_body("Define the end-to-end technical shape of enterprise AI applications and build the agent systems that carry the work.")
for item in [
    "Design architectures for Agentic AI, multi-agent systems, LLM and RAG solutions, and AI-powered enterprise applications.",
    "Build multi-agent workloads with CrewAI and LangGraph. One such system reduced content production time by 40%.",
    "Standardize how agents call external tools and APIs through the Model Context Protocol (MCP), so new tools can be added without rewriting the agent core.",
    "Engineer autonomous agents that connect LLMs to enterprise systems and complete multi-step workflows.",
    "Integrate foundation models, including Anthropic Claude 3.5 Sonnet, into product architecture. That integration automated 40% of routine code-refactoring tasks.",
    "Build autonomous coding agents. One delivered a 15% improvement in developer productivity.",
    "Use LangChain, LangGraph, CrewAI, AutoGen, Semantic Kernel, LlamaIndex, and related agent patterns (including ReAct) where they fit the customer problem.",
]:
    add_bullet(item)

add_heading_styled("3. Put AI on a production platform", 14)
add_body("Move prototypes onto infrastructure the customer can operate, observe, and scale.")
for item in [
    "Architect production-grade Generative AI platforms using MCP, LangChain, LangGraph, CrewAI, Pinecone, Tavily, Redis, Docker, and Kubernetes.",
    "Deploy MCP servers and LLM inference services on Docker and Kubernetes with autoscaling.",
    "Apply platform engineering, MLOps, and LLMOps so models, prompts, and agents can be released, monitored, and rolled back.",
    "Improve API performance with distributed caching. Redis and Memcached reduced database load by 30%.",
    "Run the platform as a service: observability with Prometheus and Grafana, SLO and SLA definition, and incident response.",
    "Use cloud platforms (Google Cloud, AWS, and Azure) and container tooling (Docker, Kubernetes, Helm) to match the customer’s estate.",
]:
    add_bullet(item)

add_heading_styled("4. Evaluate, govern, and secure the system", 14)
add_body("Treat quality, safety, and compliance as part of delivery, not a later review.")
for item in [
    "Evaluate and tune models with Precision, Recall, and F1, together with LLM fine-tuning and prompt engineering.",
    "Set responsible-AI controls, guardrails, and enterprise-readiness checks before a workload is treated as production.",
    "Use Llama Guard-based guardrails to reduce hallucination and block unauthorized data exposure.",
    "Define reliability targets (SLO/SLI) and operate to them, including error-budget and incident practices.",
]:
    add_bullet(item)

add_heading_styled("5. Lead delivery with the customer’s teams", 14)
add_body("Keep architecture, product, engineering, and business moving on the same plan.")
for item in [
    "Influence senior stakeholders and set architectural direction, governance, and the roadmap.",
    "Align Product, Engineering, Data Science, and Business, and clear technical and delivery blockers.",
    "Run Agile/Scrum planning and execution. This reduced AI product development cycle time by 20%.",
    "Use GenAI-assisted code generation and automated root-cause analysis on high-priority defects. Sprint carryover fell by 60%.",
    "Improve sprint completion and program execution by 20% through cross-functional collaboration and resource allocation.",
    "Apply AI-driven predictive modeling to program spend. Expenditure forecasts landed within 5% of actuals and supported real-time budget adjustment.",
]:
    add_bullet(item)

add_heading_styled("6. Use the data and enterprise foundation", 14)
add_body(
    "Earlier delivery as a data scientist and consultant (IBM, Michigan State University / General Motors, "
    "Fannie Mae, and Intermountain Healthcare) is what the FDE role stands on when the customer problem is data, not only a model."
)
for item in [
    "Design and implement data pipelines on Google Cloud (BigQuery, Dataflow, Dataproc, Pub/Sub) when the AI system depends on reliable enterprise data.",
    "Apply ETL and data-integration judgment (DataStage, Informatica, PL/SQL, performance tuning) so retrieval and analytics inputs are fit for production use.",
    "Coordinate on-site and offshore teams on analytics and data programs, including Bank of America Information Analytics Foundation and NetApp Managed Services work at IBM.",
]:
    add_bullet(item)

add_heading_styled("Outcomes tied to the role", 14)
add_body("Results stated on the resume and used above:")

outcomes = [
    ("Content production", "40% less time, via a multi-agent system (CrewAI and LangGraph)"),
    ("Routine code refactoring", "40% automated by integrating Claude 3.5 Sonnet into the product architecture"),
    ("Developer productivity", "15% improvement from an autonomous coding agent"),
    ("Database load", "30% reduction through Redis and Memcached distributed caching"),
    ("AI product cycle time", "20% reduction through Agile delivery, planning, and execution"),
    ("Sprint carryover", "60% reduction using GenAI code generation and automated root-cause analysis"),
    ("Sprint and program execution", "20% improvement through cross-functional collaboration"),
    ("Program spend forecasts", "Within 5% of actuals, using AI-driven predictive modeling"),
]
table = doc.add_table(rows=1, cols=2)
table.style = "Table Grid"
hdr = table.rows[0].cells
hdr[0].text = "Area"
hdr[1].text = "Result"
for cell in hdr:
    for p in cell.paragraphs:
        for run in p.runs:
            run.bold = True
            run.font.size = Pt(11)
            run.font.name = "Calibri"
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    shade_cell(cell, "1F3A5F")

for area, result in outcomes:
    row = table.add_row().cells
    row[0].text = area
    row[1].text = result
    for cell in row:
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.size = Pt(10.5)
                run.font.name = "Calibri"

add_heading_styled("What the role is accountable for", 14)
add_body("In plain terms, the Forward Deployed Engineer is accountable for:")
for item in [
    "A working production system in the customer’s environment, not a prototype left in a notebook.",
    "An architecture the customer’s engineers can extend: agents, tools (MCP), retrieval, and inference, deployed on containers with a scaling plan.",
    "Quality and safety evidence: evaluation metrics, guardrails, and reliability targets.",
    "A delivery rhythm that senior stakeholders can see: scope, sprint health, forecast accuracy, and a clear architectural roadmap.",
    "Measurable business effect, of the kind listed above, agreed with the customer before the build and checked after it ships.",
]:
    add_bullet(item)

note = doc.add_paragraph()
note.paragraph_format.space_before = Pt(16)
r = note.add_run(
    "Source: Prasad_Nimmalapudi_FDE_AI Solution Architect_Resume.pdf. "
    "Figures and tools are taken from that resume and are not expanded beyond it."
)
r.italic = True
r.font.size = Pt(9)
r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
r.font.name = "Calibri"

doc.save(OUT)
print(OUT)
