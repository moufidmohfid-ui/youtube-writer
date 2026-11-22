import hashlib
import hmac
import io
import os
import re
import secrets
import shutil
import sqlite3
import subprocess
import time
from datetime import datetime
from urllib.parse import quote

import httpx
from dotenv import load_dotenv
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from fastapi import Depends, FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from google import genai
from google.genai import types
from pydantic import BaseModel

load_dotenv()

os.makedirs("static/media", exist_ok=True)
DB_PATH = "database.db"
MEDIA_DIR = os.path.join("static", "media")


def get_secret_key() -> bytes:
    env_secret = os.environ.get("SECRET_KEY", "").strip()
    if env_secret:
        return env_secret.encode("utf-8")
    secret_file = ".secret_key"
    if os.path.exists(secret_file):
        with open(secret_file, "rb") as f:
            return f.read()
    generated = secrets.token_bytes(32)
    with open(secret_file, "wb") as f:
        f.write(generated)
    return generated


SECRET_KEY = get_secret_key()
app = FastAPI(title="منصة تفريغ وكتابة محتوى يوتيوب الاحترافية")
app.mount("/static", StaticFiles(directory="static"), name="static")


def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                youtube_url TEXT NOT NULL,
                video_id TEXT NOT NULL,
                title TEXT NOT NULL,
                language TEXT NOT NULL,
                mode TEXT NOT NULL,
                cover_image TEXT,
                markdown_content TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(id)
            )"""
        )
        conn.commit()


init_db()


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120_000)
    return f"{salt.hex()}:{dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        salt_hex, dk_hex = stored.split(":")
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(dk_hex)
        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120_000)
        return hmac.compare_digest(actual, expected)
    except Exception:
        return False


def sign_token(user_id: int, username: str) -> str:
    payload = f"{user_id}:{username}"
    sig = hmac.new(SECRET_KEY, payload.encode("utf-8"), hashlib.sha256).hexdigest()
    return f"{payload}:{sig}"


def verify_token(token: str) -> dict | None:
    try:
        parts = token.split(":")
        if len(parts) != 3:
            return None
        user_id_str, username, sig = parts
        payload = f"{user_id_str}:{username}"
        expected_sig = hmac.new(SECRET_KEY, payload.encode("utf-8"), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected_sig):
            return None
        return {"id": int(user_id_str), "username": username}
    except Exception:
        return None


def get_current_user(request: Request) -> dict:
    token = request.cookies.get("session_token")
    if not token:
        raise HTTPException(status_code=401, detail="يرجى تسجيل الدخول أولاً.")
    user = verify_token(token)
    if not user:
        raise HTTPException(status_code=401, detail="الجلسة غير صالحة أو منتهية.")
    return user


class AuthRequest(BaseModel):
    username: str
    password: str


class VideoRequest(BaseModel):
    url: str
    language: str = "العربية الفصحى الاحترافية"
    mode: str = "article"
    include_frames: bool = True


def extract_youtube_id(url: str) -> str | None:
    pattern = (
        r"(?:https?:\/\/)?(?:www\.|m\.)?"
        r"(?:youtube\.com\/(?:watch\?v=|shorts\/|embed\/|live\/)|youtu\.be\/)"
        r"([a-zA-Z0-9_-]{11})"
    )
    match = re.search(pattern, url.strip())
    return match.group(1) if match else None


def fetch_video_title(video_id: str) -> str:
    oembed_url = f"https://www.youtube.com/oembed?url=https://www.youtube.com/watch?v={video_id}&format=json"
    try:
        with httpx.Client(timeout=8.0) as client:
            r = client.get(oembed_url)
            if r.status_code == 200:
                return r.json().get("title") or f"فيديو يوتيوب ({video_id})"
    except Exception:
        pass
    return f"فيديو يوتيوب ({video_id})"


def download_cover_image(video_id: str) -> str | None:
    filename = f"{video_id}_cover.jpg"
    filepath = os.path.join(MEDIA_DIR, filename)
    if os.path.exists(filepath):
        return f"/static/media/{filename}"
    candidates = [
        f"https://img.youtube.com/vi/{video_id}/maxresdefault.jpg",
        f"https://img.youtube.com/vi/{video_id}/hqdefault.jpg",
    ]
    try:
        with httpx.Client(timeout=10.0) as client:
            for img_url in candidates:
                r = client.get(img_url)
                if r.status_code == 200 and len(r.content) > 2000:
                    with open(filepath, "wb") as f:
                        f.write(r.content)
                    return f"/static/media/{filename}"
    except Exception:
        pass
    return None


def timestamp_to_seconds(ts: str) -> int:
    parts = [int(p) for p in ts.strip().split(":")]
    if len(parts) == 2:
        return parts[0] * 60 + parts[1]
    if len(parts) == 3:
        return parts[0] * 3600 + parts[1] * 60 + parts[2]
    return 0


def get_stream_url(youtube_url: str) -> str | None:
    if not shutil.which("yt-dlp") or not shutil.which("ffmpeg"):
        return None
    try:
        res = subprocess.run(
            ["yt-dlp", "-f", "best[height<=720][ext=mp4]/best[ext=mp4]/best", "-g", youtube_url],
            capture_output=True, text=True, timeout=20,
        )
        if res.returncode == 0 and res.stdout.strip():
            return res.stdout.strip().splitlines()[0]
    except Exception:
        pass
    return None


def capture_frame(stream_url: str, timestamp: str, out_path: str) -> bool:
    if not shutil.which("ffmpeg"):
        return False
    sec = timestamp_to_seconds(timestamp)
    try:
        res = subprocess.run(
            ["ffmpeg", "-ss", str(sec), "-i", stream_url, "-frames:v", "1", "-q:v", "2", out_path, "-y"],
            capture_output=True, timeout=20,
        )
        return res.returncode == 0 and os.path.exists(out_path) and os.path.getsize(out_path) > 1000
    except Exception:
        return False


def enrich_markdown_with_frames(markdown_text: str, video_id: str, youtube_url: str, include_frames: bool) -> str:
    frame_pattern = re.compile(r"\[\[FRAME:\s*(\d{1,2}:\d{2}(?::\d{2})?)\s*\|\s*(.*?)\s*\]\]")
    matches = list(frame_pattern.finditer(markdown_text))
    if not matches:
        return markdown_text
    stream_url = get_stream_url(youtube_url) if (include_frames and shutil.which("ffmpeg")) else None
    captured_count = 0

    def replacer(match: re.Match) -> str:
        nonlocal captured_count
        ts = match.group(1).strip()
        desc = match.group(2).strip()
        if stream_url and captured_count < 6:
            safe_ts = ts.replace(":", "-")
            filename = f"{video_id}_frame_{safe_ts}_{captured_count}.jpg"
            out_path = os.path.join(MEDIA_DIR, filename)
            if capture_frame(stream_url, ts, out_path):
                captured_count += 1
                return f"\n\n![{desc} ({ts})](/static/media/{filename})\n*لقطة من الفيديو عند الدقيقة {ts}: {desc}*\n\n"
        return f"\n\n> 🎬 **لقطة مرجعية من الفيديو [{ts}]:** {desc}\n\n"

    return frame_pattern.sub(replacer, markdown_text)


def build_prompt(language: str, mode: str, include_frames: bool) -> str:
    arabic_note = ""
    if "عرب" in language or "arabic" in language.lower():
        arabic_note = (
            "- التزم باللغة العربية الفصحى السليمة والرصينة، وتجنب الترجمة الحرفية الركيكة.\n"
            "- استخدم علامات الترقيم بدقة وصِغ العبارات بسلاسة واحترافية عالية.\n"
        )
    frames_instruction = ""
    if include_frames:
        frames_instruction = (
            "- أدرج من 3 إلى 5 علامات للقطات البصرية المهمة أو الشرائح المعروضة في الفيديو داخل النص، "
            "في سطر مستقل وبالتنسيق التالي حرفياً:\n"
            "  [[FRAME: MM:SS | وصف موجز ومفيد للقطة الظاهرة]]\n"
            "  (احرص على أن يكون التوقيت MM:SS دقيقاً ومطابقاً للمشهد).\n"
        )
    if mode == "transcript":
        task = (
            "قم بتفريغ وكتابة كل ما يُقال ويُعرض في هذا الفيديو بشكل كامل ومفصل ودقيق.\n"
            "- قسّم النص إلى فقرات واضحة ومنظمة مع ذكر الطوابع الزمنية [MM:SS] عند الانتقال بين الأفكار."
        )
    elif mode == "article":
        task = (
            "حوّل محتوى هذا الفيديو كاملاً إلى مقال احترافي متكامل وجاهز للنشر.\n"
            "- ابدأ بعنوان رئيسي جذاب (# عنوان)، ثم مقدمة مميزة.\n"
            "- قسّم المقال إلى عناوين فرعية واضحة (## عنوان فرعي) مع شرح كافة التفاصيل والأمثلة.\n"
            "- اختم بخلاصة مركزة لأهم ما ورد في الفيديو."
        )
    else:
        task = (
            "قدّم تلخيصاً شاملاً واحترافياً لمحتوى الفيديو.\n"
            "- ابدأ بعنوان مناسب ثم الفكرة العامة.\n"
            "- استخرج جميع النقاط الرئيسية والأفكار العملية في شكل نقاط مرتبة مع الطوابع الزمنية [MM:SS].\n"
            "- أضف أبرز الاستنتاجات والتوصيات."
        )
    return (
        f"أنت محرر ومترجم محترف متخصص في تحويل محتوى الفيديو إلى مقالات ونصوص مكتوبة عالية الجودة.\n"
        f"اللغة المطلوبة للمخرجات بالكامل: {language}.\n\n"
        f"المطلوب:\n{task}\n\n"
        f"ضوابط الجودة:\n{arabic_note}{frames_instruction}"
        f"- لا تختلق أي معلومات غير موجودة في الفيديو.\n"
        f"- نسّق المخرجات باستخدام Markdown احترافي (عناوين، نقاط، خط عريض للمصطلحات المهمة)."
    )


def set_paragraph_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    pPr.append(bidi)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def add_formatted_runs(paragraph, text: str):
    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            run = paragraph.add_run(part)
        run.font.name = "Arial"
        run.font.size = Pt(12)


def markdown_to_docx_bytes(title: str, markdown_text: str) -> io.BytesIO:
    doc = Document()
    h = doc.add_heading(title, level=0)
    set_paragraph_rtl(h)
    img_pattern = re.compile(r"^!\[(.*?)\]\((.*?)\)$")

    for raw_line in markdown_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        img_match = img_pattern.match(line)
        if img_match:
            alt_text, img_url = img_match.group(1), img_match.group(2)
            local_path = os.path.join(MEDIA_DIR, os.path.basename(img_url))
            if os.path.exists(local_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(local_path, width=Inches(5.6))
                if alt_text:
                    cap = doc.add_paragraph(alt_text)
                    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_paragraph_rtl(cap)
            continue

        if line.startswith("### "):
            p = doc.add_heading("", level=3)
            add_formatted_runs(p, line[4:])
            set_paragraph_rtl(p)
        elif line.startswith("## "):
            p = doc.add_heading("", level=2)
            add_formatted_runs(p, line[3:])
            set_paragraph_rtl(p)
        elif line.startswith("# "):
            p = doc.add_heading("", level=1)
            add_formatted_runs(p, line[2:])
            set_paragraph_rtl(p)
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            add_formatted_runs(p, line[2:])
            set_paragraph_rtl(p)
        elif re.match(r"^\d+\.\s+", line):
            p = doc.add_paragraph(style="List Number")
            add_formatted_runs(p, re.sub(r"^\d+\.\s+", "", line))
            set_paragraph_rtl(p)
        elif line.startswith("> "):
            p = doc.add_paragraph()
            add_formatted_runs(p, line[2:])
            set_paragraph_rtl(p)
        else:
            clean_line = line.strip("*") if (line.startswith("*") and line.endswith("*") and not line.startswith("**")) else line
            p = doc.add_paragraph()
            add_formatted_runs(p, clean_line)
            set_paragraph_rtl(p)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


@app.get("/")
def read_root():
    return FileResponse("static/index.html")


@app.get("/api/auth/me")
def auth_me(request: Request):
    token = request.cookies.get("session_token")
    user = verify_token(token) if token else None
    allow_reg = os.environ.get("ALLOW_REGISTRATION", "true").lower() == "true"
    return {"user": user, "allow_registration": allow_reg}


@app.post("/api/auth/register")
def register(req: AuthRequest, response: Response):
    allow_reg = os.environ.get("ALLOW_REGISTRATION", "true").lower() == "true"
    if not allow_reg:
        raise HTTPException(status_code=403, detail="التسجيل مغلق حالياً من قِبل مدير الموقع.")
    username = req.username.strip()
    if len(username) < 3 or len(req.password) < 6:
        raise HTTPException(status_code=400, detail="اسم المستخدم يجب أن يكون 3 أحرف على الأقل وكلمة المرور 6 أحرف.")
    pw_hash = hash_password(req.password)
    now = datetime.utcnow().isoformat()
    try:
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "INSERT INTO users (username, password_hash, created_at) VALUES (?, ?, ?)",
                (username, pw_hash, now),
            )
            user_id = cur.lastrowid
            conn.commit()
    except sqlite3.IntegrityError:
        raise HTTPException(status_code=400, detail="اسم المستخدم مسجل مسبقاً.")
    token = sign_token(user_id, username)
    response.set_cookie("session_token", token, httponly=True, samesite="lax", max_age=60 * 60 * 24 * 7)
    return {"user": {"id": user_id, "username": username}}


@app.post("/api/auth/login")
def login(req: AuthRequest, response: Response):
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT id, username, password_hash FROM users WHERE username = ?",
            (req.username.strip(),),
        ).fetchone()
    if not row or not verify_password(req.password, row[2]):
        raise HTTPException(status_code=401, detail="اسم المستخدم أو كلمة المرور غير صحيحة.")
    token = sign_token(row[0], row[1])
    response.set_cookie("session_token", token, httponly=True, samesite="lax", max_age=60 * 60 * 24 * 7)
    return {"user": {"id": row[0], "username": row[1]}}


@app.post("/api/auth/logout")
def logout(response: Response):
    response.delete_cookie("session_token")
    return {"ok": True}


@app.post("/api/process")
def process_video(req: VideoRequest, user: dict = Depends(get_current_user)):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="مفتاح GEMINI_API_KEY غير مضبوط في متغيرات البيئة.")
    video_id = extract_youtube_id(req.url)
    if not video_id:
        raise HTTPException(status_code=400, detail="رابط يوتيوب غير صالح.")

    clean_url = f"https://www.youtube.com/watch?v={video_id}"
    title = fetch_video_title(video_id)
    cover_path = download_cover_image(video_id)
    prompt = build_prompt(req.language, req.mode, req.include_frames)

    try:
        client = genai.Client(
            api_key=api_key,
            http_options=types.HttpOptions(api_version="v1beta"),
        )
        models_to_try = [
            "gemini-4-hydrogen",
            "gemini-3.6-flash",
            "gemini-3.5-flash-lite",
            "gemini-3.8-flash",
        ]
        response = None
        last_error = None
        for model_name in models_to_try:
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=types.Content(
                        parts=[
                            types.Part(file_data=types.FileData(file_uri=clean_url)),
                            types.Part(text=prompt),
                        ]
                    ),
                )
                if response and response.text:
                    break
            except Exception as model_err:
                last_error = model_err
                err_str = str(model_err).upper()
                if any(k in err_str for k in ("404", "NOT_FOUND", "429", "RESOURCE_EXHAUSTED", "503", "UNAVAILABLE", "OVERLOADED")):
                    time.sleep(1)
                    continue
                raise

        if not response or not response.text:
            raise HTTPException(status_code=500, detail=f"تعذر الحصول على استجابة من النموذج: {last_error}")

        enriched_md = enrich_markdown_with_frames(response.text, video_id, clean_url, req.include_frames)
        if cover_path:
            enriched_md = f"![صورة غلاف الفيديو: {title}]({cover_path})\n\n" + enriched_md

        now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                """INSERT INTO history (user_id, youtube_url, video_id, title, language, mode, cover_image, markdown_content, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (user["id"], clean_url, video_id, title, req.language, req.mode, cover_path, enriched_md, now),
            )
            item_id = cur.lastrowid
            conn.commit()

        return {
            "id": item_id,
            "title": title,
            "video_id": video_id,
            "cover_image": cover_path,
            "result": enriched_md,
            "created_at": now,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"تعذرت معالجة الفيديو (تأكد أنه عام وليس خاصاً): {str(e)}",
        )


@app.get("/api/history")
def get_history(user: dict = Depends(get_current_user)):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """SELECT id, youtube_url, video_id, title, language, mode, cover_image, created_at
               FROM history WHERE user_id = ? ORDER BY id DESC LIMIT 50""",
            (user["id"],),
        ).fetchall()
    return {"items": [dict(r) for r in rows]}


@app.get("/api/history/{item_id}")
def get_history_item(item_id: int, user: dict = Depends(get_current_user)):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT * FROM history WHERE id = ? AND user_id = ?",
            (item_id, user["id"]),
        ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="العنصر غير موجود.")
    return dict(row)


@app.delete("/api/history/{item_id}")
def delete_history_item(item_id: int, user: dict = Depends(get_current_user)):
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            "DELETE FROM history WHERE id = ? AND user_id = ?",
            (item_id, user["id"]),
        )
        conn.commit()
    if cur.rowcount == 0:
        raise HTTPException(status_code=404, detail="العنصر غير موجود.")
    return {"ok": True}


@app.get("/api/export/docx/{item_id}")
def export_docx(item_id: int, user: dict = Depends(get_current_user)):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT title, markdown_content FROM history WHERE id = ? AND user_id = ?",
            (item_id, user["id"]),
        ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="العنصر غير موجود.")

    docx_stream = markdown_to_docx_bytes(row["title"], row["markdown_content"])
    encoded_name = quote(f"{row['title'][:40]}.docx")
    return StreamingResponse(
        docx_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
    )
