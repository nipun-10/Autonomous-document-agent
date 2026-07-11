"""
Autonomous AI Agent - FastAPI Application
Uses Groq (free tier) for LLM + python-docx for Word document generation
"""

import os
import json
import uuid
import re
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any
import asyncio

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
import httpx

# ── Config ────────────────────────────────────────────────────────────────────
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")          # set env var before running
GROQ_MODEL   = "llama-3.3-70b-versatile"               # free Groq tier
OUTPUT_DIR   = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Autonomous AI Agent", version="1.0")


# ── Schemas ───────────────────────────────────────────────────────────────────
class AgentRequest(BaseModel):
    request: str


class TaskItem(BaseModel):
    id: int
    title: str
    description: str
    status: str = "pending"
    result: str = ""


class AgentResponse(BaseModel):
    request_id: str
    original_request: str
    interpreted_goal: str
    assumptions: list[str]
    task_list: list[dict]
    execution_summary: str
    document_path: str
    document_name: str
    completed_at: str


# ── LLM Helper ────────────────────────────────────────────────────────────────
async def call_groq(system: str, user: str, temperature: float = 0.4) -> str:
    """Call Groq API (free tier). Falls back to Gemini if key missing."""
    if not GROQ_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="GROQ_API_KEY not set. Export it: export GROQ_API_KEY=your_key"
        )

    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        "temperature": temperature,
        "max_tokens": 3000,
    }

    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers=headers,
            json=payload,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]


# ── Phase 1: Planning ─────────────────────────────────────────────────────────
async def plan_tasks(user_request: str) -> dict:
    """Ask LLM to interpret the request and produce a task plan."""
    system = """You are an autonomous AI planning agent. Given a user request,
you must:
1. Interpret the true goal (even if vague or ambiguous).
2. List any assumptions you are making.
3. Break the work into 4-7 concrete tasks.
4. Decide what kind of Word document to produce.

Respond ONLY with valid JSON matching this schema:
{
  "interpreted_goal": "...",
  "document_type": "...",   // e.g. "Project Plan", "Business Proposal", "SOP", "Meeting Minutes"
  "assumptions": ["..."],
  "tasks": [
    {"id": 1, "title": "...", "description": "..."},
    ...
  ]
}"""

    raw = await call_groq(system, user_request)
    # Strip markdown fences if present
    clean = re.sub(r"```json|```", "", raw).strip()
    try:
        return json.loads(clean)
    except json.JSONDecodeError:
        # Fallback: extract JSON block
        match = re.search(r"\{.*\}", clean, re.DOTALL)
        if match:
            return json.loads(match.group())
        raise HTTPException(status_code=500, detail=f"LLM returned invalid JSON: {raw[:300]}")


# ── Phase 2: Execute each task ────────────────────────────────────────────────
async def execute_task(task: dict, context: dict) -> str:
    """Ask LLM to produce content for a single task."""
    system = f"""You are an expert business analyst and writer executing step {task['id']} of a multi-step plan.
Goal: {context['interpreted_goal']}
Document type: {context['document_type']}
Assumptions: {json.dumps(context['assumptions'])}

Generate rich, professional content for this task. Use realistic mock data where needed.
Be specific, detailed, and production-quality. Do NOT add JSON wrappers — just write the content."""

    user = f"Task: {task['title']}\nDetails: {task['description']}"
    return await call_groq(system, user, temperature=0.6)


# ── Phase 3: Build .docx via Node script ─────────────────────────────────────
def build_docx(plan: dict, task_results: list[dict], output_path: Path) -> None:
    """Generate the Word document using docx (npm) via a Node.js script."""

    # Sanitize text for safe JSON embedding
    def safe(text: str) -> str:
        return text.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "")

    doc_type   = safe(plan.get("document_type", "Business Document"))
    goal       = safe(plan.get("interpreted_goal", ""))
    assumptions = plan.get("assumptions", [])
    date_str   = datetime.now().strftime("%B %d, %Y")

    # Build task sections JSON
    sections_js = "[\n"
    for t in task_results:
        title   = safe(t["title"])
        content = safe(t["result"])
        sections_js += f'  {{ title: "{title}", content: "{content}" }},\n'
    sections_js += "]"

    # Build assumptions list JS
    assumptions_js = json.dumps([str(a) for a in assumptions])

    node_script = f"""
const {{ Document, Packer, Paragraph, TextRun, HeadingLevel,
        AlignmentType, BorderStyle, ShadingType, Table, TableRow,
        TableCell, WidthType, PageNumber, NumberFormat }} = require("docx");
const fs = require("fs");

const docType      = "{doc_type}";
const goal         = "{goal}";
const dateStr      = "{date_str}";
const assumptions  = {assumptions_js};
const sections     = {sections_js};
const outputPath   = "{output_path.as_posix()}";

// ── Helpers ────────────────────────────────────────
function hr() {{
  return new Paragraph({{
    border: {{ bottom: {{ style: BorderStyle.SINGLE, size: 6, color: "2E74B5" }} }},
    spacing: {{ after: 200 }},
  }});
}}

function heading1(text) {{
  return new Paragraph({{
    text,
    heading: HeadingLevel.HEADING_1,
    spacing: {{ before: 400, after: 160 }},
  }});
}}

function heading2(text) {{
  return new Paragraph({{
    text,
    heading: HeadingLevel.HEADING_2,
    spacing: {{ before: 280, after: 120 }},
  }});
}}

function body(text) {{
  return text.split("\\\\n").map(line =>
    new Paragraph({{
      children: [new TextRun({{ text: line, size: 22 }})],
      spacing: {{ after: 120 }},
    }})
  );
}}

// ── Cover Page ─────────────────────────────────────
const coverChildren = [
  new Paragraph({{ spacing: {{ before: 1440 }} }}),
  new Paragraph({{
    children: [new TextRun({{ text: docType.toUpperCase(), bold: true, size: 56, color: "1F3864" }})],
    alignment: AlignmentType.CENTER,
    spacing: {{ after: 240 }},
  }}),
  hr(),
  new Paragraph({{
    children: [new TextRun({{ text: goal, size: 28, color: "2E74B5", italics: true }})],
    alignment: AlignmentType.CENTER,
    spacing: {{ after: 480 }},
  }}),
  new Paragraph({{
    children: [new TextRun({{ text: "Prepared by Autonomous AI Agent", size: 22, color: "595959" }})],
    alignment: AlignmentType.CENTER,
  }}),
  new Paragraph({{
    children: [new TextRun({{ text: dateStr, size: 22, color: "595959" }})],
    alignment: AlignmentType.CENTER,
    spacing: {{ after: 2880 }},
  }}),
  new Paragraph({{ pageBreakBefore: true }}),
];

// ── Assumptions ────────────────────────────────────
const assumptionChildren = [
  heading1("Agent Assumptions & Scope"),
  hr(),
  new Paragraph({{
    children: [new TextRun({{ text: "The following assumptions were made during autonomous planning:", size: 22, italics: true, color: "595959" }})],
    spacing: {{ after: 160 }},
  }}),
  ...assumptions.map((a, i) =>
    new Paragraph({{
      children: [
        new TextRun({{ text: `${{i + 1}}. `, bold: true, size: 22 }}),
        new TextRun({{ text: a, size: 22 }}),
      ],
      spacing: {{ after: 100 }},
    }})
  ),
  new Paragraph({{ pageBreakBefore: true }}),
];

// ── Main Sections ──────────────────────────────────
const mainChildren = [];
sections.forEach((sec, idx) => {{
  mainChildren.push(heading1(`${{idx + 1}}. ${{sec.title}}`));
  mainChildren.push(hr());
  mainChildren.push(...body(sec.content));
  if (idx < sections.length - 1) {{
    mainChildren.push(new Paragraph({{ pageBreakBefore: true }}));
  }}
}});

// ── Document ───────────────────────────────────────
const doc = new Document({{
  styles: {{
    default: {{
      heading1: {{
        run: {{ bold: true, size: 32, color: "1F3864", font: "Calibri" }},
      }},
      heading2: {{
        run: {{ bold: true, size: 26, color: "2E74B5", font: "Calibri" }},
      }},
    }},
  }},
  sections: [
    {{
      properties: {{ page: {{ size: {{ width: 12240, height: 15840 }} }} }},
      children: [
        ...coverChildren,
        ...assumptionChildren,
        ...mainChildren,
      ],
    }},
  ],
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync(outputPath, buffer);
  console.log("DOCX written:", outputPath);
}}).catch(err => {{
  console.error(err);
  process.exit(1);
}});
"""

    script_path = OUTPUT_DIR / "_build_doc.js"
    script_path.write_text(node_script, encoding="utf-8")

    result = subprocess.run(
        ["node", str(script_path)],
        capture_output=True, text=True,
        cwd=str(Path(__file__).parent)
    )
    if result.returncode != 0:
        raise RuntimeError(f"Node docx builder failed:\n{result.stderr}")


# ── Phase 4: Summarise execution ──────────────────────────────────────────────
async def summarise(plan: dict, task_results: list[dict]) -> str:
    titles = [t["title"] for t in task_results]
    system = "You are an executive summariser. Write a crisp 3-5 sentence executive summary."
    user   = (
        f"Goal: {plan['interpreted_goal']}\n"
        f"Document type: {plan['document_type']}\n"
        f"Completed tasks: {titles}\n"
        "Summarise what was accomplished and the value of the document produced."
    )
    return await call_groq(system, user)


# ── Main Endpoint ─────────────────────────────────────────────────────────────
@app.post("/agent", response_model=AgentResponse)
async def run_agent(body: AgentRequest):
    request_id = str(uuid.uuid4())[:8]
    print(f"\n[{request_id}] REQUEST: {body.request}")

    # ── Phase 1: Plan ──────────────────────────────
    print(f"[{request_id}] Phase 1: Planning...")
    plan = await plan_tasks(body.request)
    print(f"[{request_id}]   Goal: {plan['interpreted_goal']}")
    print(f"[{request_id}]   Tasks: {len(plan['tasks'])}")

    # ── Phase 2: Execute tasks ─────────────────────
    task_results = []
    for task in plan["tasks"]:
        print(f"[{request_id}]   Executing task {task['id']}: {task['title']}")
        result_text = await execute_task(task, plan)
        task_results.append({**task, "status": "done", "result": result_text})
        await asyncio.sleep(3)

    # ── Phase 3: Build document ────────────────────
    print(f"[{request_id}] Phase 3: Building .docx...")
    doc_name   = f"agent_{request_id}_{plan['document_type'].replace(' ', '_')}.docx"
    doc_path   = OUTPUT_DIR / doc_name
    build_docx(plan, task_results, doc_path)
    print(f"[{request_id}]   Saved: {doc_path}")

    # ── Phase 4: Summarise ─────────────────────────
    print(f"[{request_id}] Phase 4: Summarising...")
    summary = await summarise(plan, task_results)

    return AgentResponse(
        request_id       = request_id,
        original_request = body.request,
        interpreted_goal = plan["interpreted_goal"],
        assumptions      = plan.get("assumptions", []),
        task_list        = task_results,
        execution_summary= summary,
        document_path    = str(doc_path),
        document_name    = doc_name,
        completed_at     = datetime.now().isoformat(),
    )


@app.get("/download/{filename}")
async def download_doc(filename: str):
    path = OUTPUT_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )


@app.get("/")
async def root():
    return {"status": "running", "endpoints": ["POST /agent", "GET /download/{filename}"]}
