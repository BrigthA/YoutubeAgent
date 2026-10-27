"""Live Streamlit dashboard for AutoGen YouTube content agents with nested SCALE JSON formatting, XLSX export, and Gmail SMTP Attachment."""
import os
import threading
import time
import smtplib
import json
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from uuid import uuid4

import streamlit as st
from dotenv import load_dotenv
from autogen import AssistantAgent, UserProxyAgent, GroupChat, GroupChatManager
import openpyxl

load_dotenv()

NAMES = ["YouTube_Researcher", "Scriptwriter", "SEO_Expert", "Video_Prompt_Engineer", "Email_Notifier_Agent"]

PROMPTS = {
    "YouTube_Researcher": "You research YouTube content. Provide 3 compelling angles and clearly distinguish verifiable facts. Finish with concise research notes.",
    "Scriptwriter": "You write engaging 60-second YouTube Shorts scripts using the research in the conversation. Include a 10-second hook, paced body, and a clear call-to-action.",
    "SEO_Expert": "You optimize YouTube Shorts. Provide 3 high-CTR titles, an optimized description with realistic timestamps, and tags.",
    "Video_Prompt_Engineer": (
        "You convert the YouTube script scenes into the strict SCALE framework JSON format. "
        "You MUST output a valid JSON array of scene objects. Each object MUST match this exact schema:\n"
        "{\n"
        '  "framework": "SCALE",\n'
        '  "scene_number": 1,\n'
        '  "title": "Scene description title",\n'
        '  "subject": { "primary": "...", "people": [...], "cultural_details": [...] },\n'
        '  "composition": { "framing": "...", "camera_position": "...", "foreground": "...", "midground": "...", "background": "...", "depth": "...", "visual_balance": "...", "composition_style": "..." },\n'
        '  "action": { "primary_action": "...", "details": [...], "movement": "...", "expression": "..." },\n'
        '  "location": { "real_place": "...", "architecture": [...], "environment": "...", "avoid": [...] },\n'
        '  "time_of_day": { "specific_time": "...", "season": "...", "historical_context": "..." },\n'
        '  "lighting": { "source": "...", "direction": "...", "quality": "...", "atmosphere": "...", "avoid": [...] }\n'
        "}\n"
        "Return ONLY the valid JSON array representing the scenes for the script."
    ),
    "Email_Notifier_Agent": "You handle summarizing the final YouTube package so it is ready to be emailed to the user."
}

def parse_scale_json(content_str):
    """Safely extracts and parses JSON data from the agent's output text."""
    try:
        cleaned = re.sub(r"```json|```", "", content_str).strip()
        start_idx = cleaned.find('[')
        end_idx = cleaned.rfind(']')
        if start_idx != -1 and end_idx != -1:
            data = json.loads(cleaned[start_idx:end_idx+1])
            if isinstance(data, list):
                return data
    except Exception:
        pass
    return None

def generate_excel_package(topic, outputs, filepath):
    """Creates a polished Excel spreadsheet containing all agent outputs, placing the complete nested SCALE JSON string directly into the 5th column."""
    wb = openpyxl.Workbook()
    
    # Setup Overview Tab
    ws_summary = wb.active
    ws_summary.title = "Overview"
    ws_summary.append(["YouTube Content Production Package"])
    ws_summary.append(["Topic:", topic])
    ws_summary.append(["Generated At:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws_summary.append([])
    ws_summary.append(["Section", "Agent Responsible"])
    ws_summary.append(["1. Research & Angles", NAMES[0]])
    ws_summary.append(["2. YouTube Script", NAMES[1]])
    ws_summary.append(["3. SEO Package", NAMES[2]])
    ws_summary.append(["4. SCALE Video Prompts (Nested JSON)", NAMES[3]])

    # Standard text sheets
    standard_sections = [
        ("Research & Angles", NAMES[0]),
        ("YouTube Script", NAMES[1]),
        ("SEO Metadata", NAMES[2]),
        ("Email Summary", NAMES[4])
    ]

    for title, name in standard_sections:
        ws = wb.create_sheet(title=title[:31])
        ws.append(["Section", title])
        ws.append(["Content", outputs.get(name, "No response generated.")])
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 80

    # Structured SCALE Video Prompts Sheet (Column 5 contains the raw scene JSON object)
    ws_scale = wb.create_sheet(title="SCALE Video Prompts")
    ws_scale.append(["Scene #", "Title", "Primary Subject", "Location", "SCALE JSON Object (Full Schema)"])
    
    scale_raw = outputs.get(NAMES[3], "")
    parsed_scenes = parse_scale_json(scale_raw)

    if parsed_scenes:
        for scene in parsed_scenes:
            # Extract basic overview fields for columns 1-4, and place the full JSON string in column 5
            scene_num = str(scene.get("scene_number", "1"))
            scene_title = str(scene.get("title", ""))
            primary_subj = str(scene.get("subject", {}).get("primary", ""))
            location_place = str(scene.get("location", {}).get("real_place", ""))
            full_json_str = json.dumps(scene, indent=2, ensure_ascii=False)

            ws_scale.append([
                scene_num,
                scene_title,
                primary_subj,
                location_place,
                full_json_str
            ])
    else:
        # Fallback if JSON parsing fails
        ws_scale.append(["Raw Output", "", "", "", scale_raw])

    # Adjust column widths (making Column 5 wide to comfortably display the JSON schema)
    col_widths = {'A': 12, 'B': 30, 'C': 35, 'D': 30, 'E': 75}
    for col, width in col_widths.items():
        ws_scale.column_dimensions[col].width = width

    wb.save(filepath)

def send_smtp_email_with_attachment(recipient_email, subject, body, attachment_path):
    """Sends the email with the .xlsx file attached via Gmail SMTP App Password."""
    sender_email = os.getenv("SMTP_SENDER_EMAIL")
    app_password = os.getenv("SMTP_APP_PASSWORD")
    
    if not sender_email or not app_password or not recipient_email:
        print("SMTP credentials or recipient email missing. Skipping email dispatch.")
        return False

    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))

        if attachment_path and Path(attachment_path).exists():
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {Path(attachment_path).name}")
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
            
        print(f"Email with XLSX attachment successfully sent to {recipient_email}")
        return True
    except Exception as exc:
        print(f"Failed to send email via SMTP: {exc}")
        return False

def new_run(topic, model, recipient_email):
    return {
        "id": uuid4().hex, "topic": topic, "model": model, "recipient_email": recipient_email, "lock": threading.Lock(),
        "status": "Running", "error": "", "started": time.monotonic(), "finished": None,
        "agents": {name: {"status": "Pending", "response": "", "prompt_tokens": 0,
                          "completion_tokens": 0, "total_tokens": 0, "cost": 0.0,
                          "duration": 0.0, "calls": 0, "usage_available": False}
                   for name in NAMES},
        "output": "", "path": "", "events": [],
    }

def update(run, name, **values):
    with run["lock"]:
        run["agents"][name].update(values)
        run["events"].append((datetime.now().strftime("%H:%M:%S"), name, values.get("status", "Updated")))

def usage_for(agent):
    try:
        usage = agent.get_actual_usage() or {}
        models = [v for v in usage.values() if isinstance(v, dict)]
        if not models:
            return None
        return {
            "prompt_tokens": sum(int(v.get("prompt_tokens", 0)) for v in models),
            "completion_tokens": sum(int(v.get("completion_tokens", 0)) for v in models),
            "total_tokens": sum(int(v.get("total_tokens", 0)) for v in models),
            "cost": float(usage.get("total_cost", 0) or 0),
        }
    except Exception:
        return None

def worker(run, api_key):
    try:
        config = {"config_list": [{"model": run["model"], "api_key": api_key}], "temperature": 0.7}

        class MonitoredAgent(AssistantAgent):
            def generate_reply(self, messages=None, sender=None, **kwargs):
                update(run, self.name, status="Running")
                start = time.monotonic()
                try:
                    reply = super().generate_reply(messages=messages, sender=sender, **kwargs)
                    content = reply.get("content", "") if isinstance(reply, dict) else str(reply or "")
                    usage = usage_for(self)
                    fields = {"status": "Completed", "response": content,
                              "duration": round(time.monotonic() - start, 2), "calls": 1,
                              "usage_available": usage is not None}
                    if usage:
                        fields.update(usage)
                    update(run, self.name, **fields)
                    return reply
                except Exception as exc:
                    update(run, self.name, status="Failed", response=str(exc),
                           duration=round(time.monotonic() - start, 2))
                    raise

        admin = UserProxyAgent(name="Admin", 
                               human_input_mode="NEVER", 
                               code_execution_config=False,
                               default_auto_reply="Continue.")
        agents = [MonitoredAgent(
                    name=name, 
                    system_message=PROMPTS[name], 
                    llm_config=config)
                  for name in NAMES]
        
        group = GroupChat(agents=[admin, *agents], 
                          messages=[], 
                          max_round=6,
                          speaker_selection_method="round_robin")
                         
        manager = GroupChatManager(groupchat=group, llm_config=config)
        admin.initiate_chat(manager, message=f"Create a complete 60-second YouTube package (Research, Script, SEO, and nested JSON SCALE video generation prompts) about: {run['topic']}",
                            clear_history=True, silent=True)

        with run["lock"]:
            outputs = {name: run["agents"][name]["response"] for name in NAMES}
            
        # Generate XLSX in root directory
        output_dir = Path(__file__).resolve().parent
        xlsx_filename = f"youtube_content_{datetime.now():%Y-%m-%d_%H%M%S}_{run['id'][:6]}.xlsx"
        xlsx_path = output_dir / xlsx_filename
        
        generate_excel_package(run["topic"], outputs, str(xlsx_path))

        email_body = f"Hello,\n\nYour automated YouTube AI production package for the topic '{run['topic']}' is ready.\n\nPlease find the structured Excel workbook (.xlsx) attached containing your Research, Script, SEO Metadata, and the full nested SCALE framework JSON schemas in column 5.\n\nBest regards,\nYour YouTube AI Studio"

        if run["recipient_email"]:
            send_smtp_email_with_attachment(
                run["recipient_email"], 
                f"YouTube AI Package (XLSX): {run['topic']}", 
                email_body, 
                str(xlsx_path)
            )

        with run["lock"]:
            run.update(status="Completed", finished=time.monotonic(), output=f"Excel Workbook generated and saved at: {xlsx_path}", path=str(xlsx_path))
    except Exception as exc:
        with run["lock"]:
            run.update(status="Failed", error=f"{type(exc).__name__}: {exc}", finished=time.monotonic())

st.set_page_config(page_title="YouTube AI Agent Monitor", page_icon="🎬", layout="wide")
st.title("🎬 YouTube AI Agent Dashboard + Nested SCALE JSON & Email Attachment")
st.caption("AutoGen GroupChat • Structured Excel Table with JSON Column 5 • Gmail App Password Attachment")

with st.sidebar:
    st.header("New generation")
    topic = st.text_area("YouTube topic", value="Authentic Tamil Nadu Chettinad Heritage", height=110)
    model = st.selectbox("OpenAI model", ["gpt-4o-mini", "gpt-4o"])
    recipient_email = st.text_input("Recipient Email Address", value=os.getenv("USER_EMAIL", ""))
    st.caption("API keys and SMTP credentials are read from your .env file.")
    
    active = st.session_state.get("active_run")
    busy = bool(active and active["status"] == "Running")
    if st.button("▶ Generate XLSX & Email", type="primary", use_container_width=True, disabled=busy):
        key = os.getenv("OPENAI_API_KEY")
        if not key:
            st.error("Missing OPENAI_API_KEY. Add it to your .env file.")
        elif not topic.strip():
            st.error("Enter a topic.")
        else:
            run = new_run(topic.strip(), model, recipient_email.strip())
            st.session_state.active_run = run
            threading.Thread(target=worker, args=(run, key), daemon=True).start()
            st.rerun()
    st.divider()
    st.caption("Generated XLSX files are saved directly in the root directory.")

@st.fragment(run_every="1s")
def dashboard():
    run = st.session_state.get("active_run")
    if not run:
        st.info("Choose a topic, confirm your email, and click Generate XLSX & Email to start.")
        return
    with run["lock"]:
        status = run["status"]
        error = run["error"]
        agents = {k: v.copy() for k, v in run["agents"].items()}
        events = list(run["events"])
        output = run["output"]
        path = run["path"]
        elapsed = (run["finished"] or time.monotonic()) - run["started"]

    if status == "Running":
        st.info("🔴 Live — agents are working (Research, Script, SEO, Nested SCALE JSON Video, Excel Table Export, Email Attachment). Refreshing every second.")
    elif status == "Completed":
        st.success("All agent turns completed. Structured Excel table with raw JSON column saved to root and emailed as an attachment.")
    else:
        st.error(f"Generation failed: {error}")

    completed = sum(a["status"] == "Completed" for a in agents.values())
    total_calls = sum(a["calls"] for a in agents.values())
    prompt_tokens = sum(a["prompt_tokens"] for a in agents.values())
    completion_tokens = sum(a["completion_tokens"] for a in agents.values())
    total_cost = sum(a["cost"] for a in agents.values())
    c = st.columns(6)
    for col, label, value in zip(c,
        ["Agents complete", "Agent LLM calls*", "Prompt tokens", "Completion tokens", "Total tokens", "Elapsed"],
        [f"{completed}/5", str(total_calls), f"{prompt_tokens:,}", f"{completion_tokens:,}",
         f"{prompt_tokens + completion_tokens:,}", f"{elapsed:.1f}s"]):
        col.metric(label, value)
    st.progress(completed / 5)
    st.metric("Reported API cost (USD)", f"${total_cost:.6f}")

    st.subheader("Agent watch live")
    tabs = st.tabs([f"{name} · {agents[name]['status']}" for name in NAMES])
    for tab, name in zip(tabs, NAMES):
        a = agents[name]
        with tab:
            st.write(f"**Agent:** {name}   |   **Status:** {a['status']}   |   **Duration:** {a['duration']:.2f}s")
            x, y, z = st.columns(3)
            x.metric("Input tokens", f"{a['prompt_tokens']:,}")
            y.metric("Output tokens", f"{a['completion_tokens']:,}")
            z.metric("Total tokens", f"{a['total_tokens']:,}")
            st.markdown("**Response**")
            st.text_area("Agent output", value=a["response"] or "Waiting for agent response...",
                         height=280, disabled=True, key=f"response_{run['id']}_{name}", label_visibility="collapsed")

    st.subheader("Execution timeline")
    if events:
        st.dataframe([{"Time": t, "Agent": n, "Event": e} for t, n, e in events],
                    use_container_width=True, hide_index=True)
    if path and Path(path).exists():
        st.subheader("Final deliverable")
        with open(path, "rb") as f:
            st.download_button("⬇ Download root XLSX", data=f, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               file_name=Path(path).name, key=f"download_{run['id']}")

dashboard()