"""Ford MSS incentives workflow document with an embedded flowchart."""
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:\Users\LENOVO\OneDrive\Desktop\Ford MSS Customer Incentives and Rewards - GenAI Workflow with Flowchart - Nimmalapudi Prasad.docx"
IMG = r"c:\Users\LENOVO\OneDrive\Desktop\Enterprise-HR-Policy-Employee-Support-Agentic-RAG-Copilot-New\Enterprise-HR-Policy-Employee-Support-Agentic-RAG-Copilot-New\tmp\ford_incentives_flow.png"

NAVY = (31, 58, 95)
TEAL = (30, 107, 122)
GOLD = (140, 98, 32)
GREEN = (36, 102, 72)
SLATE = (55, 65, 78)
WHITE = (255, 255, 255)
LINE = (31, 58, 95)
BG = (248, 249, 251)


def font(size, bold=False):
    path = r"C:\Windows\Fonts\calibrib.ttf" if bold else r"C:\Windows\Fonts\calibri.ttf"
    return ImageFont.truetype(path, size)


def rounded(draw, box, radius, fill, outline):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=2)


def wrap(draw, text, fnt, max_w):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = w if not cur else cur + " " + w
        if draw.textlength(trial, font=fnt) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def text_center(draw, xy, text, fnt, fill):
    box = draw.textbbox((0, 0), text, font=fnt)
    w, h = box[2] - box[0], box[3] - box[1]
    draw.text((xy[0] - w / 2, xy[1] - h / 2), text, font=fnt, fill=fill)


def arrow_down(draw, x, y1, y2, label=None, fnt=None):
    draw.line((x, y1, x, y2), fill=LINE, width=4)
    draw.polygon([(x - 9, y2 - 14), (x + 9, y2 - 14), (x, y2)], fill=LINE)
    if label and fnt:
        draw.text((x + 12, (y1 + y2) / 2 - 14), label, font=fnt, fill=GREEN)


def arrow_right(draw, x1, x2, y, label=None, fnt=None):
    draw.line((x1, y, x2, y), fill=LINE, width=4)
    draw.polygon([(x2 - 14, y - 8), (x2 - 14, y + 8), (x2, y)], fill=LINE)
    if label and fnt:
        draw.text((x1 + 10, y - 32), label, font=fnt, fill=(140, 50, 40))


def process_box(draw, x, y, w, h, title, sub, color, f_title, f_sub):
    rounded(draw, (x, y, x + w, y + h), 16, WHITE, color)
    draw.rectangle((x, y + 8, x + 10, y + h - 8), fill=color)
    draw.text((x + 28, y + 16), title, font=f_title, fill=color)
    draw.text((x + 28, y + 56), sub, font=f_sub, fill=SLATE)
    return x + w / 2, y, y + h, x + w


def diamond(draw, cx, top, w, h, text, fnt):
    mid_y = top + h / 2
    pts = [(cx, top), (cx + w / 2, mid_y), (cx, top + h), (cx - w / 2, mid_y)]
    draw.polygon(pts, fill=(255, 248, 236), outline=GOLD)
    draw.line(pts + [pts[0]], fill=GOLD, width=3)
    text_center(draw, (cx, mid_y - 2), text, fnt, GOLD)
    return cx, top, top + h, cx + w / 2


def terminal(draw, x, y, w, h, text, fnt, fill, outline):
    rounded(draw, (x, y, x + w, y + h), 18, fill, outline)
    text_center(draw, (x + w / 2, y + h / 2), text, fnt, WHITE if fill != WHITE else outline)
    return x + w / 2, y, y + h


def draw_flow():
    W, H = 1680, 2480
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)
    f_title = font(30, True)
    f_sub = font(22)
    f_dec = font(26, True)
    f_lab = font(22, True)
    f_end = font(24, True)
    f_head = font(34, True)

    draw.text((48, 28), "Ford customer incentives and rewards", font=f_head, fill=NAVY)
    draw.text((48, 74), "Generative AI decision flow", font=f_sub, fill=SLATE)

    mx, mw, mh = 70, 760, 108
    y = 130
    gap = 46

    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "1  Marketing, Sales and Service request",
        "Reward, eligible customer, and success metric",
        NAVY, f_title, f_sub,
    )
    y = bot + gap
    arrow_down(draw, cx, bot, y)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "2  Customer, sales, and service data",
        "Cloud Storage, BigQuery, Pub/Sub, Dataflow",
        TEAL, f_title, f_sub,
    )
    y = bot + gap
    arrow_down(draw, cx, bot, y)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "3  Incentive rules and reward catalog",
        "Published policy is the only source of a reward",
        TEAL, f_title, f_sub,
    )
    y = bot + gap
    arrow_down(draw, cx, bot, y)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "4  Embed rules and retrieve",
        "Vertex AI embeddings, Vector Search or Pinecone",
        NAVY, f_title, f_sub,
    )

    y = bot + gap
    arrow_down(draw, cx, bot, y)
    dw, dh = 420, 150
    dx, dtop, dbot, dright = diamond(draw, cx, y, dw, dh, "Rule found?", f_dec)
    side_x = 1080
    side_w = 500
    side_h = 84
    arrow_right(draw, dright, side_x, dtop + dh / 2, "No", f_lab)
    terminal(draw, side_x, dtop + dh / 2 - side_h / 2, side_w, side_h, "Human review", f_end, (120, 72, 48), (120, 72, 48))

    y = dbot + gap
    arrow_down(draw, cx, dbot, y, "Yes", f_lab)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "5  RAG eligibility on Vertex AI",
        "Decide only from retrieved rules and customer facts",
        NAVY, f_title, f_sub,
    )
    y = bot + gap
    arrow_down(draw, cx, bot, y)
    dx, dtop, dbot, dright = diamond(draw, cx, y, dw, dh, "Eligible?", f_dec)
    arrow_right(draw, dright, side_x, dtop + dh / 2, "No", f_lab)
    terminal(draw, side_x, dtop + dh / 2 - side_h / 2, side_w, side_h, "Do not issue reward", f_end, (140, 50, 40), (140, 50, 40))

    y = dbot + gap
    arrow_down(draw, cx, dbot, y, "Yes", f_lab)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "6  Incentive spend forecast",
        "Release the segment only within 5% of forecast",
        GOLD, f_title, f_sub,
    )
    y = bot + gap
    arrow_down(draw, cx, bot, y)
    dx, dtop, dbot, dright = diamond(draw, cx, y, dw, dh, "Within 5%?", f_dec)
    arrow_right(draw, dright, side_x, dtop + dh / 2, "No", f_lab)
    terminal(draw, side_x, dtop + dh / 2 - side_h / 2, side_w, side_h, "Hold and notify Marketing", f_end, (140, 98, 32), (140, 98, 32))

    y = dbot + gap
    arrow_down(draw, cx, dbot, y, "Yes", f_lab)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "7  Agents draft the offer",
        "LangGraph path, CrewAI writer and reviewer",
        TEAL, f_title, f_sub,
    )
    y = bot + gap
    arrow_down(draw, cx, bot, y)
    dx, dtop, dbot, dright = diamond(draw, cx, y, 520, dh, "Llama Guard pass?", f_dec)
    arrow_right(draw, dright, side_x, dtop + dh / 2, "No", f_lab)
    terminal(draw, side_x, dtop + dh / 2 - side_h / 2, side_w, side_h, "Refuse or send to reviewer", f_end, (36, 102, 72), (36, 102, 72))

    y = dbot + gap
    arrow_down(draw, cx, dbot, y, "Yes", f_lab)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "8  Deliver one approved reward",
        "MCP tools: email, sales, and service",
        NAVY, f_title, f_sub,
    )
    y = bot + gap
    arrow_down(draw, cx, bot, y)
    cx, top, bot, right = process_box(
        draw, mx, y, mw, mh,
        "9  Run and measure",
        "GKE, Redis, Precision / Recall / F1, Prometheus",
        SLATE, f_title, f_sub,
    )
    y = bot + 40
    arrow_down(draw, cx, bot, y)
    terminal(draw, mx + 160, y, 440, 78, "Reward issued", f_end, GREEN, GREEN)

    img = img.crop((0, 0, W, min(H, y + 78 + 40)))
    img.save(IMG, "PNG")
    return IMG


def shade(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def h(doc, text, size=16):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    r.font.name = "Calibri"


def body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    r.font.size = Pt(11)
    r.font.name = "Calibri"
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.clear()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.size = Pt(11)
    r.font.name = "Calibri"


def main():
    draw_flow()
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.75)
    sec.bottom_margin = Inches(0.75)
    sec.left_margin = Inches(0.85)
    sec.right_margin = Inches(0.85)
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    t = doc.add_paragraph()
    r = t.add_run("Ford Marketing, Sales and Service")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    r.font.name = "Calibri"

    s = doc.add_paragraph()
    r = s.add_run("Customer Incentives and Rewards Program")
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x6B, 0x7A)
    r.font.name = "Calibri"

    s2 = doc.add_paragraph()
    r = s2.add_run("Generative AI workflow on Google Cloud  |  Nimmalapudi Prasad")
    r.italic = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    r.font.name = "Calibri"

    body(
        doc,
        "This document explains each step of the Generative AI workflow used for Ford's Marketing, Sales and Service customer incentives and rewards program. "
        "A customer becomes eligible for a reward only when sales and service facts match published incentive rules, the offer stays inside the program budget, and the message that goes out is grounded in those rules. "
        "The platform choices and the measured results below are the ones on the Ford Forward Deployed Engineer resume: RAG, Vertex AI, LangGraph, CrewAI, MCP, GKE, Memorystore, Llama Guard, and the delivery and forecast outcomes.",
    )

    h(doc, "Flowchart", 16)
    body(doc, "Read the chart from top to bottom. Diamonds are decisions. A No branch stops the reward. A Yes branch continues. The sections after the chart explain each step.")
    doc.add_picture(IMG, width=Inches(6.7))

    h(doc, "Step 1. Take the request from Marketing, Sales and Service", 14)
    body(
        doc,
        "The program starts with the business request, not with a model. Marketing defines the reward, such as a service credit, a loyalty offer, or a sales incentive. Sales and Service define who can receive it: the vehicle, the ownership event, the service visit, or the campaign segment. Finance defines the budget the program must not exceed.",
    )
    body(doc, "Before any build, three things are written down:")
    bullet(doc, "The decision: is this customer eligible, and for which reward.")
    bullet(doc, "The data the decision may use: sales, service, and published incentive rules. Nothing else.")
    bullet(doc, "The success metric: a correct reward, a message the customer can act on, and program spend that stays close to the forecast.")
    body(
        doc,
        "As the Forward Deployed Engineer you sit with those stakeholders, turn the request into an architecture, and stay until the offer runs in production.",
    )

    h(doc, "Step 2. Land customer, sales, and service data on GCP", 14)
    body(
        doc,
        "Eligibility is only as good as the customer record. Purchase and ownership facts, service history, campaign membership, and prior rewards are brought into Google Cloud. Raw files land in Cloud Storage. Structured facts used for eligibility and for the spend forecast land in BigQuery. New sales and service events arrive on Pub/Sub so a reward can be considered when the event happens, not only in a nightly batch. Dataflow prepares and joins those inputs. Cloud Composer schedules the refresh.",
    )
    body(
        doc,
        "This is the same data-pipeline discipline as the Ford email datamart work. The 620 Email DataMart process was redesigned from 7 hours 30 minutes to 2 hours 10 minutes, so campaign data is ready in the window Marketing needs. IAM and Secret Manager keep customer data inside the project, and agents reach it only through approved tools.",
    )

    h(doc, "Step 3. Publish the incentive rules and the reward catalog", 14)
    body(
        doc,
        "The rules of the program are documents: who qualifies, which vehicles, which service events, what the reward is, what the exclusions are, and when the offer expires. Those documents are the source of truth. They are split into chunks with a stable id, the program name, the effective date, and the source file.",
    )
    body(
        doc,
        "A reward the model cannot point back to a chunk is not a reward the program may issue. This is what keeps a generative system from inventing a discount.",
    )

    h(doc, "Step 4. Embed the rules and index them", 14)
    body(
        doc,
        "Each chunk is embedded with a Vertex AI embedding model and written to a vector index. On a GCP-native stack that index is Vertex AI Vector Search. The resume also names Pinecone, so the same retrieval step can sit on Pinecone when that is the index already in the customer environment.",
    )
    body(
        doc,
        "At decision time the customer event is embedded as a query. The index returns the rule chunks that match the vehicle, the region, and the incentive, and it returns the catalog entry for the reward itself.",
    )

    h(doc, "Step 5. Decide eligibility with RAG", 14)
    body(
        doc,
        "Retrieval-augmented generation combines the retrieved rules with the customer facts from BigQuery. Vertex AI answers a narrow question: eligible or not, which reward, and which clause supports it. The prompt, the model version, and the retrieval settings are versioned together so a bad change can be rolled back.",
    )
    body(
        doc,
        "Fine-tuning and prompt engineering are used when the first prompt is not stable enough. The release gate is Precision, Recall, and F1 on a golden set of known customers, not a demo conversation. If retrieval returns no rule, the workflow stops and a person reviews the case.",
    )

    h(doc, "Step 6. Forecast incentive spend before the offer goes out", 14)
    body(
        doc,
        "An eligible customer is not automatically paid. Program expenditure is forecast from the sales and service pipeline in BigQuery, and the offer is released only when the forecast still fits the budget. On the Ford resume, AI-driven predictive modeling kept program expenditure forecasts within 5% of actuals and allowed the budget to be adjusted while the program was running.",
    )
    body(
        doc,
        "The forecast is an input to the agent, not a slide after the campaign. If accepting the next segment would break the 5% band, the workflow holds the segment and notifies Marketing.",
    )

    h(doc, "Step 7. Generate the customer reward message with agents", 14)
    body(
        doc,
        "LangGraph holds the path: retrieve the rule, confirm eligibility, check the forecast, draft the message, check the draft, then send or stop. CrewAI splits the content work into roles, for example a writer and a reviewer, which is the multi-agent pattern that cut content production time by 40%.",
    )
    body(
        doc,
        "The message states the reward, the condition the customer met, the expiry, and the action to take at sales or service. It does not add a benefit that was not in the retrieved catalog. External actions go through MCP servers: the agent calls a tool, and the tool calls the CRM, the email system, or the dealer channel. A new channel is a new MCP server, not a rewrite of the agent.",
    )

    h(doc, "Step 8. Check the offer with Llama Guard", 14)
    body(
        doc,
        "Llama Guard runs before the model sees a free-text request and again before a message is released. It blocks attempts to pull customers outside the allowed data, and it blocks answers that are off policy or that are not grounded in the retrieved incentive text.",
    )
    body(
        doc,
        "A failed check does not produce a softer offer. The case is refused or sent to a Marketing or Service reviewer. That is the control that keeps a rewards program from hallucinating a benefit.",
    )

    h(doc, "Step 9. Deliver through Marketing, Sales, and Service", 14)
    body(
        doc,
        "The approved reward is delivered on the channel the program named. Email uses the customer communication path that the email datamart feeds. Hard-bounce handling, recognized at Ford for a think-out-of-the-box approach, keeps invalid addresses out of the send so a reward is not reported as delivered when it never arrived. Sales and service channels receive the same approved payload through their MCP tool, so the dealer and the inbox do not show two different offers.",
    )
    body(
        doc,
        "Repeated lookups of the same eligibility and catalog reads sit behind Memorystore for Redis and Memcached. That distributed cache is the change that reduced database load by 30% and keeps the decision path fast during a campaign peak.",
    )

    h(doc, "Step 10. Deploy, evaluate, and operate", 14)
    body(
        doc,
        "The API, the MCP servers, and the inference services are container images in Artifact Registry, released with Cloud Build and Helm onto GKE. Pods autoscale with demand. Prometheus and Grafana show latency, retrieval hit rate, guardrail blocks, token use, and error rate. SLOs define the availability and latency the program owes Marketing during a launch. An incident rolls back the prompt or the model version and keeps the last good index.",
    )
    body(doc, "A change ships only if evaluation still holds:")
    bullet(doc, "Precision, Recall, and F1 on eligibility stay at the agreed bar.")
    bullet(doc, "Every issued reward cites a rule chunk.")
    bullet(doc, "Spend for the released segment remains inside the 5% forecast band.")
    bullet(doc, "Content for the offer is produced by the agent workflow, which is the path tied to the 40% reduction in content time.")
    body(
        doc,
        "Delivery of the surrounding program work follows the same cadence as the Ford GenAI delivery results: about 20% shorter cycle time, about 60% less sprint carryover through GenAI code generation and automated root-cause analysis, and about 20% better sprint completion through cross-team execution.",
    )

    h(doc, "What each Ford result means in this program", 14)
    rows = [
        ("Email datamart runtime", "7 hours 30 minutes reduced to 2 hours 10 minutes, so reward audiences are ready in the campaign window"),
        ("Hard-bounce handling", "Reward email is not counted as delivered when the address is dead"),
        ("Content production", "40% less time, via CrewAI and LangGraph offer messages"),
        ("Database load", "30% lower, via Redis and Memcached on repeated eligibility reads"),
        ("Program spend forecast", "Within 5% of actuals before a segment is released"),
        ("Eligibility quality", "Precision, Recall, and F1 on a golden customer set"),
        ("Delivery health", "Prometheus, Grafana, and SLO/SLA on the offer API"),
    ]
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Result"
    hdr[1].text = "Role in incentives and rewards"
    for c in hdr:
        shade(c, "1F3A5F")
        for p in c.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(11)
                run.font.name = "Calibri"
    for a, b in rows:
        cells = table.add_row().cells
        cells[0].text = a
        cells[1].text = b
        for c in cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(10.5)
                    run.font.name = "Calibri"

    h(doc, "One-minute version", 14)
    body(
        doc,
        "For Ford Marketing, Sales and Service, the incentives and rewards program takes a customer event, matches it to published rules with RAG, checks the reward against a spend forecast held within 5%, and lets LangGraph and CrewAI write the offer. Llama Guard stops anything the rules do not support. MCP tools deliver one approved reward to email, sales, and service. The service runs on GKE with Redis in front, and Prometheus, Grafana, and Precision, Recall, and F1 decide whether it stays in production.",
    )

    note = doc.add_paragraph()
    note.paragraph_format.space_before = Pt(12)
    r = note.add_run(
        "Source: Prasad_Nimmalapudi_FDE_AI Solution Architect_Resume.pdf, Ford Technology Services India. "
        "Measured figures are the resume figures. They are applied here to the customer incentives and rewards workflow and are not new results."
    )
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    r.font.name = "Calibri"

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
