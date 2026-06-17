import streamlit as st
import os
import imaplib, email, json, io, smtplib, urllib.request
from email.header    import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from typing   import List, Dict, Tuple, Optional
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Agent Mail IA", page_icon="🦾",
                   layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
.main { background:#f8f9fa; }
.stButton>button { border-radius:8px; font-weight:600; }
.metric-card {
    background:white; border-radius:10px; padding:16px;
    box-shadow:0 2px 8px rgba(0,0,0,.08); margin:6px 0;
}
.urgence-critique { border-left:4px solid #e94560; }
.urgence-haute    { border-left:4px solid #f39c12; }
.urgence-normale  { border-left:4px solid #3498db; }
.urgence-faible   { border-left:4px solid #95a5a6; }
.chat-user {
    background:#e8f4fd; border-radius:12px 12px 2px 12px;
    padding:10px 14px; margin:6px 0 6px 40px; font-size:.9em;
}
.chat-ai {
    background:#f0f9f0; border-radius:12px 12px 12px 2px;
    padding:10px 14px; margin:6px 40px 6px 0; font-size:.9em;
    border-left:3px solid #2ecc71;
}
.log-container {
    background:#1e1e1e; color:#a8ff78; border-radius:8px;
    padding:12px; font-family:monospace; font-size:.82em;
    height:280px; overflow-y:auto;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════
for key, default in [
    ("mails",        []),
    ("logs",         []),
    ("historique",   []),
    ("connected",    False),
    ("ollama_ok",    False),
    ("ollama_models",[]),
    ("triage_done",  False),
    # conversations[mail_uid] = list of {role, content}
    ("conversations",{}),
    ("ssh_tunnel_proc", None),
    ("ssh_tunnel_active", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

URGENCY_EMOJI = {"critique":"🔴","haute":"🟠","normale":"🟡","faible":"⚪"}
URGENCY_COLOR = {"critique":"#e94560","haute":"#f39c12","normale":"#3498db","faible":"#95a5a6"}
MODELE_CONSEILLE = "qwen2.5:7b-instruct-q4_K_M"

# ══════════════════════════════════════════════════════════════
# UTILS
# ══════════════════════════════════════════════════════════════
def log(msg:str, level:str="INFO"):
    ts  = datetime.now().strftime("%H:%M:%S")
    tag = {"INFO":"ℹ️","OK":"✅","ERR":"❌","LLM":"🧠","MAIL":"📧","SMTP":"📤"}
    st.session_state.logs.append(f"[{ts}] {tag.get(level,'•')} {msg}")

def check_ollama(url:str) -> Tuple[bool, List[str]]:
    try:
        r = urllib.request.urlopen(f"{url}/api/tags", timeout=3)
        d = json.loads(r.read())
        return True, [m["name"] for m in d.get("models",[])]
    except:
        return False, []

def ollama_chat(messages:List[Dict], model:str, url:str,
                temperature:float=0.3) -> str:
    """
    Envoie TOUT l'historique messages à Ollama.
    Retourne le texte de la réponse.
    """
    payload = json.dumps({
        "model": model,
        "messages": messages,
        "stream": False,
        "options": {"temperature": temperature, "num_ctx": 8192},
    }).encode()
    req = urllib.request.Request(
        f"{url}/api/chat", data=payload,
        headers={"Content-Type":"application/json"})
    with urllib.request.urlopen(req, timeout=300) as resp:
        data = json.loads(resp.read())
    return data["message"]["content"].strip()

# ══════════════════════════════════════════════════════════════
# GMAIL READER
# ══════════════════════════════════════════════════════════════
class GmailReader:
    def __init__(self, user, pw):
        self.user, self.pw, self.conn = user, pw, None

    def connect(self):
        try:
            self.conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
            self.conn.login(self.user, self.pw)
            return True, "OK"
        except Exception as e:
            return False, str(e)

    def disconnect(self):
        try: self.conn and self.conn.logout()
        except: pass

    @staticmethod
    def _dec(v):
        if not v: return ""
        parts = decode_header(v)
        out = []
        for b, enc in parts:
            if isinstance(b, bytes):
                out.append(b.decode(enc or "utf-8", errors="replace"))
            else:
                out.append(str(b))
        return " ".join(out).strip()

    def _body(self, msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ct = part.get_content_type()
                cd = str(part.get("Content-Disposition",""))
                if ct=="text/plain" and "attachment" not in cd:
                    try:
                        body += part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="replace")
                    except: pass
        else:
            try:
                body = msg.get_payload(decode=True).decode(
                    msg.get_content_charset() or "utf-8", errors="replace")
            except: body=""
        return body.replace("\r\n","\n").replace("\r","\n")[:6000]

    def _attachments(self, msg):
        atts = []
        for part in msg.walk():
            fn = self._dec(part.get_filename())
            if not fn: continue
            data = part.get_payload(decode=True) or b""
            atts.append({"filename":fn,
                          "content_type":part.get_content_type(),
                          "data":data,
                          "size_kb":round(len(data)/1024,1)})
        return atts

    def fetch(self, n=10, unread_only=False):
        if not self.conn: return []
        self.conn.select("INBOX")
        _, ids = self.conn.search(None, "UNSEEN" if unread_only else "ALL")
        ids = ids[0].split()
        if not ids: return []
        mails = []
        for uid in ids[-n:][::-1]:
            try:
                _, data = self.conn.fetch(uid,"(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                mails.append({
                    "uid":         uid.decode(),
                    "from":        self._dec(msg["From"]),
                    "to":          self._dec(msg["To"]),
                    "subject":     self._dec(msg["Subject"]) or "(Sans objet)",
                    "date":        msg["Date"] or "",
                    "body":        self._body(msg),
                    "attachments": self._attachments(msg),
                    "analyse":     None,
                    "pj_content":  None,
                    "brouillon_envoye": False,
                })
            except: pass
        return mails

# ══════════════════════════════════════════════════════════════
# ATTACHMENT PARSER
# ══════════════════════════════════════════════════════════════
class AttachmentParser:
    MAX = 4000
    def parse(self, att) -> str:
        ct = att["content_type"].lower()
        fn = att["filename"].lower()
        try:
            if "pdf"   in ct or fn.endswith(".pdf"):             return self._pdf(att)
            if "csv"   in ct or fn.endswith(".csv"):             return self._csv(att)
            if "excel" in ct or fn.endswith((".xlsx",".xls")):   return self._excel(att)
            if "text"  in ct or fn.endswith(".txt"):             return self._txt(att)
            if ct.startswith("image/"):  return f"[IMAGE: {att['filename']} — non lisible en texte]"
            return f"[{att['filename']} — type non pris en charge]"
        except Exception as e:
            return f"[Erreur lecture {att['filename']}: {e}]"

    def _pdf(self, att):
        try:
            import PyPDF2
            r    = PyPDF2.PdfReader(io.BytesIO(att["data"]))
            text = "\n".join(p.extract_text() or "" for p in r.pages)
            return f"[PDF — {att['filename']} ({len(r.pages)} pages)]\n{text[:self.MAX]}"
        except ImportError:
            return f"[PDF {att['filename']} — installer PyPDF2]"

    def _csv(self, att):
        text = att["data"].decode("utf-8", errors="replace")
        try:
            df = pd.read_csv(io.StringIO(text))
            return (f"[CSV — {att['filename']} : {len(df)} lignes]\n"
                    f"Colonnes : {list(df.columns)}\n"
                    f"Aperçu :\n{df.head(8).to_string()}\n"
                    f"Statistiques :\n{df.describe().round(2).to_string()}")[:self.MAX]
        except:
            return f"[CSV {att['filename']}]\n{text[:self.MAX]}"

    def _excel(self, att):
        try:
            df = pd.read_excel(io.BytesIO(att["data"]))
            return (f"[Excel — {att['filename']}]\n"
                    f"Colonnes : {list(df.columns)}\n"
                    f"Aperçu :\n{df.head(8).to_string()}")[:self.MAX]
        except Exception as e:
            return f"[Excel {att['filename']} — pip install openpyxl : {e}]"

    def _txt(self, att):
        return f"[Texte — {att['filename']}]\n{att['data'].decode('utf-8',errors='replace')[:self.MAX]}"

PARSER = AttachmentParser()

# ══════════════════════════════════════════════════════════════
# TRIAGE JSON
# ══════════════════════════════════════════════════════════════
def trier_mail(mail:Dict, model:str, url:str) -> Dict:
    pj_names = [a["filename"] for a in mail["attachments"]]
    pj_info  = f"\nPièces jointes : {', '.join(pj_names)}" if pj_names else ""
    schema = ('{"urgence":"critique|haute|normale|faible",'
              '"categorie":"action_requise|information|question|spam|commercial",'
              '"resume":"1 phrase max 120 car",'
              '"action_suggeree":"max 100 car",'
              '"delai_reponse":"immediat|aujourd_hui|cette_semaine|aucun",'
              '"confiance":0.95}')
    prompt = (f"Analyse cet email. Réponds UNIQUEMENT en JSON valide :\n{schema}\n\n"
              f"De: {mail['from']}\nObjet: {mail['subject']}\n"
              f"Date: {mail['date']}{pj_info}\n\n{mail['body'][:2000]}\nJSON:")
    log(f"Triage : {mail['subject'][:45]}", "LLM")
    raw = ollama_chat(
        [{"role":"system","content":"Tu réponds UNIQUEMENT en JSON valide sans texte autour. Français."},
         {"role":"user","content":prompt}],
        model=model, url=url, temperature=0.05)
    raw = raw.replace("```json","").replace("```","").strip()
    s,e = raw.find("{"), raw.rfind("}")+1
    if s>=0 and e>s: raw=raw[s:e]
    try:
        res = json.loads(raw)
        log(f"Triage OK → {res.get('urgence','?')} | {res.get('categorie','?')}", "OK")
        return res
    except:
        log("Triage JSON parse error — fallback", "ERR")
        return {"urgence":"normale","categorie":"information",
                "resume":mail["subject"][:120],"action_suggeree":"Lire",
                "delai_reponse":"cette_semaine","confiance":0.5}

# ══════════════════════════════════════════════════════════════
# CONVERSATION AVEC HISTORIQUE
# ══════════════════════════════════════════════════════════════
def build_system_prompt(mail:Dict) -> str:
    """Construit le message système enrichi avec le contexte complet du mail."""
    a        = mail.get("analyse") or {}
    pj_text  = mail.get("pj_content") or ""
    pj_sect  = f"\n\n--- CONTENU DES PIÈCES JOINTES ---\n{pj_text[:4000]}" if pj_text else ""
    att_names= ", ".join(x["filename"] for x in mail["attachments"]) if mail["attachments"] else "aucune"

    return f"""Tu es un assistant de messagerie professionnelle. Tu aides à rédiger des réponses d'emails en français.

=== CONTEXTE DU MAIL EN COURS ===
De      : {mail['from']}
À       : {mail['to']}
Objet   : {mail['subject']}
Date    : {mail['date']}
PJ      : {att_names}
Urgence : {a.get('urgence','non analysé')}
Catégorie: {a.get('categorie','non analysée')}
Résumé  : {a.get('resume','')}
Action  : {a.get('action_suggeree','')}

=== CORPS DU MAIL ORIGINAL ===
{mail['body'][:3000]}{pj_sect}
=================================

RÈGLES :
- Tu rédiges uniquement en français professionnel
- Tu tiens compte du contenu des PJ dans tes réponses
- Tu te souviens de tout ce qui a été dit dans cette conversation
- Quand on te demande un brouillon : commence par "Bonjour [prénom]," et conclus par "Cordialement,\\n[Votre nom]"
- Tu peux reformuler, raccourcir, changer le ton sur demande"""

def chat_avec_historique(mail:Dict, user_message:str,
                         model:str, url:str, temperature:float=0.35) -> str:
    """
    Envoie un message en maintenant tout l'historique de la conversation.
    La conversation est stockée dans st.session_state.conversations[mail_uid].
    """
    uid = mail["uid"]

    # Initialiser la conversation si elle n'existe pas
    if uid not in st.session_state.conversations:
        st.session_state.conversations[uid] = []

    conv = st.session_state.conversations[uid]

    # Message système (toujours en tête, mis à jour si les PJ ont changé)
    system_msg = {"role":"system", "content": build_system_prompt(mail)}

    # Ajouter le message utilisateur à l'historique
    conv.append({"role":"user","content":user_message})

    # Construire la liste complète : system + tout l'historique
    messages = [system_msg] + conv

    log(f"Chat ({len(conv)} msgs historique) : {user_message[:50]}", "LLM")

    # Appel Ollama avec tout l'historique
    response = ollama_chat(messages, model=model, url=url, temperature=temperature)

    # Ajouter la réponse à l'historique
    conv.append({"role":"assistant","content":response})
    st.session_state.conversations[uid] = conv

    log(f"Réponse reçue ({len(response)} car)", "OK")
    return response

# ══════════════════════════════════════════════════════════════
# SMTP
# ══════════════════════════════════════════════════════════════
def envoyer_email(user:str, pw:str, to:str,
                  subject:str, body:str) -> Tuple[bool,str]:
    try:
        msg = MIMEMultipart("alternative")
        msg["From"]    = user
        msg["To"]      = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(user, pw)
            s.send_message(msg)
        log(f"Email envoyé à {to}", "SMTP")
        return True, "OK"
    except Exception as e:
        log(f"Erreur SMTP : {e}", "ERR")
        return False, str(e)

# ══════════════════════════════════════════════════════════════
# ─────────────────────  SIDEBAR  ─────────────────────────────
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Paramètres")

    st.markdown("### 📧 Gmail")
    gmail_user = st.text_input("Email", value="anne.centrale.mediterranee@gmail.com")
    gmail_pass = st.text_input("App Password", value="", type="password",
                               placeholder="16 caractères sans espaces")

    st.markdown("### 🦙 Ollama")

    # ── Mode local ou SSH ─────────────────────────────────────
    mode_ollama = st.radio("Mode inférence", ["💻 Local", "🖥️ Serveur SSH"],
                           horizontal=True, key="mode_ollama")

    if mode_ollama == "💻 Local":
        # ── MODE LOCAL ─────────────────────────────────────────
        ollama_url = "http://localhost:11434"
        st.caption("Ollama tourne sur cette machine.")

        # Fermer le tunnel SSH s'il était ouvert
        if st.session_state.get("ssh_tunnel_active"):
            proc = st.session_state.get("ssh_tunnel_proc")
            if proc:
                try: proc.terminate()
                except: pass
            st.session_state.ssh_tunnel_active = False
            st.session_state.ssh_tunnel_proc   = None
            log("Tunnel SSH fermé", "OK")

    else:
        # ── MODE SSH ───────────────────────────────────────────
        st.markdown("**Connexion SSH**")
        ssh_host     = st.text_input("Hôte",         value="195.221.220.18", key="ssh_host")
        ssh_port     = st.number_input("Port SSH",   value=22003, step=1, key="ssh_port")
        ssh_user     = st.text_input("Utilisateur",  value="sysadmin", key="ssh_user")
        ssh_key      = st.text_input("Clé privée",
                                     value=os.path.expanduser("~/.ssh/id_rsa"),
                                     key="ssh_key",
                                     help="Laisser vide pour utiliser le mot de passe")
        ssh_pass     = st.text_input("Mot de passe SSH (si pas de clé)", value="",
                                     type="password", key="ssh_pass")
        remote_port  = st.number_input("Port Ollama distant", value=11434, step=1,
                                       key="remote_port")
        local_port   = st.number_input("Port local tunnel",   value=11435, step=1,
                                       key="local_port")

        ollama_url = f"http://localhost:{int(local_port)}"

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            btn_open_tunnel = st.button("🔗 Ouvrir tunnel",
                                        type="primary",
                                        width='stretch',
                                        disabled=st.session_state.ssh_tunnel_active,
                                        key="btn_open_tunnel")
        with col_t2:
            btn_close_tunnel = st.button("✂️ Fermer tunnel",
                                         width='stretch',
                                         disabled=not st.session_state.ssh_tunnel_active,
                                         key="btn_close_tunnel")

        if btn_open_tunnel:
            import subprocess, shutil
            ssh_bin = shutil.which("ssh")
            if not ssh_bin:
                st.error("Commande ssh introuvable sur cette machine.")
            else:
                cmd = [
                    ssh_bin,
                    "-N",                          # pas de commande distante
                    "-L", f"{int(local_port)}:localhost:{int(remote_port)}",
                    "-p", str(int(ssh_port)),
                    "-o", "StrictHostKeyChecking=no",
                    "-o", "ServerAliveInterval=30",
                    "-o", "ServerAliveCountMax=3",
                    "-o", "ExitOnForwardFailure=yes",
                    "-o", "ConnectTimeout=10",
                ]
                # Clé privée ou mot de passe
                if ssh_key and os.path.isfile(ssh_key):
                    cmd += ["-i", ssh_key]
                elif ssh_pass:
                    # sshpass requis pour le mot de passe non-interactif
                    sshpass = shutil.which("sshpass")
                    if sshpass:
                        cmd = [sshpass, "-p", ssh_pass] + cmd
                    else:
                        st.warning("sshpass non installé — utiliser une clé SSH ou "
                                   "installer sshpass : `sudo apt install sshpass`")
                cmd.append(f"{ssh_user}@{ssh_host}")

                try:
                    proc = subprocess.Popen(
                        cmd,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.PIPE,
                    )
                    import time; time.sleep(2)  # laisser le tunnel s'établir

                    if proc.poll() is not None:
                        err = proc.stderr.read().decode(errors="replace")
                        st.error(f"Tunnel SSH échoué : {err[:300]}")
                        log(f"Tunnel SSH échoué : {err[:200]}", "ERR")
                    else:
                        st.session_state.ssh_tunnel_proc   = proc
                        st.session_state.ssh_tunnel_active = True
                        log(f"Tunnel SSH ouvert : localhost:{int(local_port)} → "
                            f"{ssh_host}:{int(remote_port)}", "OK")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erreur ouverture tunnel : {e}")
                    log(f"Erreur tunnel : {e}", "ERR")

        if btn_close_tunnel:
            proc = st.session_state.get("ssh_tunnel_proc")
            if proc:
                try: proc.terminate()
                except: pass
            st.session_state.ssh_tunnel_active = False
            st.session_state.ssh_tunnel_proc   = None
            log("Tunnel SSH fermé", "OK")
            st.rerun()

        # Statut tunnel
        if st.session_state.ssh_tunnel_active:
            proc = st.session_state.get("ssh_tunnel_proc")
            alive = proc and proc.poll() is None
            if alive:
                st.success(f"Tunnel actif : localhost:{int(local_port)} → "
                           f"{ssh_host}:{int(remote_port)}")
            else:
                st.warning("Tunnel SSH coupé — rouvrir")
                st.session_state.ssh_tunnel_active = False
        else:
            st.info("Tunnel non actif — cliquer 'Ouvrir tunnel'")

        with st.expander("📋 Commande SSH équivalente"):
            key_part = f"-i {ssh_key} " if ssh_key and os.path.isfile(ssh_key) else ""
            st.code(
                f"ssh -N "
                f"-L {int(local_port)}:localhost:{int(remote_port)} "
                f"-p {int(ssh_port)} {key_part}"
                f"{ssh_user}@{ssh_host}",
                language="bash"
            )
            st.caption("Vous pouvez aussi ouvrir ce tunnel manuellement dans un terminal.")

    # ── Vérification Ollama (local ou via tunnel) ─────────────
    ok_ol, models_ol = check_ollama(ollama_url)
    st.session_state.ollama_ok     = ok_ol
    st.session_state.ollama_models = models_ol

    if ok_ol:
        src = "SSH" if mode_ollama != "💻 Local" else "local"
        st.success(f"Ollama {src} actif — {len(models_ol)} modèle(s)")
        default_idx = 0
        if MODELE_CONSEILLE in models_ol:
            default_idx = models_ol.index(MODELE_CONSEILLE)
        ollama_model = st.selectbox("Modèle", models_ol, index=default_idx)
    else:
        st.error("Ollama non disponible")
        if mode_ollama == "💻 Local":
            st.code(f"ollama serve\nollama pull {MODELE_CONSEILLE}")
        else:
            st.code(f"# Sur le serveur distant :\nOLLAMA_HOST=0.0.0.0 ollama serve")
        ollama_model = st.text_input("Modèle (manuel)", value=MODELE_CONSEILLE)

    with st.expander("💡 Modèle recommandé"):
        st.markdown(f"""
**CPU local 16 GB → `{MODELE_CONSEILLE}`** (~5-10 tok/s)

**Serveur GPU distant → bien mieux !**
- `qwen2.5:14b-instruct-q4_K_M` si GPU 16GB
- `qwen2.5:32b-instruct-q4_K_M` si GPU 24GB+
- `llama3.3:70b-instruct-q4_K_M` si GPU 48GB+

```bash
# Sur le serveur distant
ollama pull qwen2.5:14b-instruct-q4_K_M
OLLAMA_HOST=0.0.0.0:11434 ollama serve
```
        """)

    st.markdown("### 📬 Lecture")
    nb_mails    = st.slider("Nb de mails", 5, 50, 10)
    unread_only = st.checkbox("Non-lus seulement", False)

    st.markdown("### 🎚️ Génération")
    temp_reply = st.slider("Température réponse", 0.1, 0.9, 0.35, 0.05)
    auto_pj    = st.checkbox("Parser PJ automatiquement", True)
    urg_brouillons = st.multiselect("Triage auto — urgences",
        ["critique","haute","normale","faible"], default=["critique","haute"])

    st.divider()
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        if st.button("🗑️ Logs", width='stretch'):
            st.session_state.logs = []
    with col_r2:
        if st.button("🔄 Reset", width='stretch'):
            for k in ["mails","logs","historique","conversations","triage_done"]:
                st.session_state[k] = [] if k != "triage_done" and k != "conversations" else (False if k=="triage_done" else {})
            st.rerun()

# ══════════════════════════════════════════════════════════════
# ─────────────────────  MAIN  ────────────────────────────────
# ══════════════════════════════════════════════════════════════
st.title("🦾 Agent Mail IA — Gmail + Ollama")
st.caption(f"Lecture · Triage · Analyse PJ · Chat avec historique · Envoi SMTP | Modèle conseillé : {MODELE_CONSEILLE}")

tab_mails, tab_triage, tab_redaction, tab_stats, tab_logs = st.tabs([
    "📬 Mails", "🔴 Triage", "✉️ Rédaction & Envoi", "📊 Stats", "🖥️ Logs"
])

# ══════════════════════════════════════════════════════════════
# TAB 1 — MAILS
# ══════════════════════════════════════════════════════════════
with tab_mails:
    col_h, col_btn = st.columns([3,1])
    with col_h:
        st.subheader("Boîte de réception Gmail")
    with col_btn:
        if st.button("📨 Lire mes mails", type="primary", width='stretch'):
            if not gmail_pass:
                st.error("App Password manquant")
            else:
                with st.spinner("Connexion IMAP..."):
                    reader = GmailReader(gmail_user, gmail_pass)
                    ok, msg = reader.connect()
                    if not ok:
                        st.error(f"Erreur IMAP : {msg}")
                        log(f"Erreur connexion : {msg}", "ERR")
                    else:
                        log(f"Connecté ({gmail_user})", "OK")
                        mails = reader.fetch(n=nb_mails, unread_only=unread_only)
                        reader.disconnect()
                        st.session_state.mails = mails
                        st.session_state.triage_done = False
                        st.session_state.conversations = {}
                        st.session_state.sel_mail_idx = 0
                        # Effacer tous les états liés aux mails précédents
                        for k in list(st.session_state.keys()):
                            if any(k.startswith(p) for p in
                                   ["draft_","pending_","msg_input_","to_","subj_"]):
                                del st.session_state[k]
                        log(f"{len(mails)} mails chargés", "OK")
                        st.rerun()

    if st.session_state.mails:
        st.write(f"**{len(st.session_state.mails)} mails chargés**")
        for i, m in enumerate(st.session_state.mails):
            a   = m.get("analyse") or {}
            urg = a.get("urgence","")
            emo = URGENCY_EMOJI.get(urg,"") if urg else ""
            pj  = len(m["attachments"])
            env = "✅ " if m.get("brouillon_envoye") else ""
            uid = m["uid"]
            has_conv = uid in st.session_state.conversations and len(st.session_state.conversations[uid]) > 0
            chat_badge = f" 💬×{len(st.session_state.conversations[uid])//2}" if has_conv else ""

            with st.expander(f"{env}{emo} [{i+1}] {m['subject'][:60]}"
                             + (f" 📎×{pj}" if pj else "") + chat_badge):
                c1,c2 = st.columns([2,1])
                with c1:
                    st.write(f"**De :** {m['from']}")
                    st.write(f"**Date :** {m['date']}")
                    if pj: st.write(f"**PJ :** {', '.join(x['filename'] for x in m['attachments'])}")
                with c2:
                    if a:
                        st.write(f"**Urgence :** {emo} {urg.upper()}")
                        st.write(f"**Action :** {a.get('action_suggeree','')}")
                st.text_area("Corps", m["body"][:1500], height=160,
                             key=f"body_{i}", disabled=True)
                if pj and not m.get("pj_content"):
                    if st.button("📎 Parser les PJ", key=f"pj_btn_{i}"):
                        with st.spinner("Lecture des PJ..."):
                            parts = [PARSER.parse(a_) for a_ in m["attachments"]]
                            m["pj_content"] = "\n\n".join(parts)
                            log(f"PJ parsées : {', '.join(x['filename'] for x in m['attachments'])}", "OK")
                        st.rerun()
                elif m.get("pj_content"):
                    st.text_area("Contenu PJ (extrait)", m["pj_content"][:800],
                                 height=120, key=f"pj_view_{i}", disabled=True)
    else:
        st.info("Cliquer sur 'Lire mes mails' pour charger la boîte.")

# ══════════════════════════════════════════════════════════════
# TAB 2 — TRIAGE
# ══════════════════════════════════════════════════════════════
with tab_triage:
    st.subheader("Triage automatique par Ollama")

    _, col_triage_btn = st.columns([3,1])
    with col_triage_btn:
        btn_triage = st.button("🔴 Lancer le triage", type="primary",
                               width='stretch',
                               disabled=not (st.session_state.mails and st.session_state.ollama_ok))

    if btn_triage:
        total = len(st.session_state.mails)
        prog  = st.progress(0, text="Triage en cours...")
        for i, mail in enumerate(st.session_state.mails):
            prog.progress((i+1)/total, text=f"{i+1}/{total} : {mail['subject'][:40]}")
            if auto_pj and mail["attachments"] and not mail.get("pj_content"):
                mail["pj_content"] = "\n\n".join(PARSER.parse(a) for a in mail["attachments"])
            try:
                mail["analyse"] = trier_mail(mail, ollama_model, ollama_url)
            except Exception as e:
                log(f"Erreur triage : {e}", "ERR")
                mail["analyse"] = {"urgence":"normale","categorie":"information",
                    "resume":mail["subject"][:120],"action_suggeree":"Lire",
                    "delai_reponse":"cette_semaine","confiance":0.5}
        st.session_state.triage_done = True
        prog.progress(1.0, text="Triage terminé !")
        st.success(f"{total} mails triés")
        st.rerun()

    if st.session_state.triage_done and st.session_state.mails:
        rank = {"critique":0,"haute":1,"normale":2,"faible":3}
        triees = sorted(st.session_state.mails,
                        key=lambda m: rank.get((m.get("analyse") or {}).get("urgence","normale"),2))
        for m in triees:
            a   = m.get("analyse") or {}
            urg = a.get("urgence","normale")
            col = URGENCY_COLOR.get(urg,"#95a5a6")
            emo = URGENCY_EMOJI.get(urg,"🟡")
            pj  = len(m["attachments"])
            env = "✅ Envoyé | " if m.get("brouillon_envoye") else ""
            st.markdown(
                f'<div class="metric-card urgence-{urg}">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                f'<b>{emo} {m["subject"][:65]}</b>'
                f'<span style="background:{col};color:white;padding:2px 10px;'
                f'border-radius:12px;font-size:.8em;">{urg.upper()}</span></div>'
                f'<div style="color:#666;font-size:.88em;margin-top:4px;">'
                f'{env}De : {m["from"][:45]} | {m["date"][:25]}'
                f'{"| 📎 "+str(pj)+" PJ" if pj else ""}</div>'
                f'<div style="margin-top:6px;"><b>Résumé :</b> {a.get("resume","")}<br>'
                f'<b>Action :</b> {a.get("action_suggeree","")} | '
                f'<b>Délai :</b> {a.get("delai_reponse","")} | '
                f'<b>Confiance :</b> {a.get("confiance",0):.0%}</div></div>',
                unsafe_allow_html=True)
    elif not st.session_state.mails:
        st.info("Charger les mails d'abord.")

# ══════════════════════════════════════════════════════════════
# TAB 3 — REDACTION & ENVOI (avec historique de conversation)
# ══════════════════════════════════════════════════════════════
with tab_redaction:
    st.subheader("✉️ Rédaction, conversation IA & envoi")

    if not st.session_state.mails:
        st.info("Charger les mails d'abord (onglet 📬 Mails)")
    else:
        rank_urg = {"critique":0,"haute":1,"normale":2,"faible":3}
        mails_tries = sorted(
            st.session_state.mails,
            key=lambda m: rank_urg.get((m.get("analyse") or {}).get("urgence","normale"), 2)
        )

        # ── Sélection du mail — clé ancrée en session_state ──────────────
        # On stocke l'index sélectionné pour le préserver entre les reruns
        if "sel_mail_idx" not in st.session_state:
            st.session_state.sel_mail_idx = 0
        # Garder l'index dans les bornes si les mails ont changé
        st.session_state.sel_mail_idx = min(
            int(st.session_state.sel_mail_idx), len(mails_tries)-1)

        def mail_label(m):
            a   = m.get("analyse") or {}
            urg = a.get("urgence","")
            emo = URGENCY_EMOJI.get(urg,"📧") if urg else "📧"
            env = "✅ " if m.get("brouillon_envoye") else ""
            nb_msg = len(st.session_state.conversations.get(m["uid"],[]))//2
            chat = f" 💬×{nb_msg}" if nb_msg else ""
            return f"{env}{emo}  {m['subject'][:55]}  —  {m['from'][:28]}{chat}"

        # on_change met à jour sel_mail_idx ET efface le pending_prompt
        def on_mail_change():
            st.session_state.sel_mail_idx = int(st.session_state["_mail_select_widget"])
            st.session_state.pending_prompt = ""

        st.selectbox(
            "1️⃣  Mail auquel répondre",
            options=range(len(mails_tries)),
            format_func=lambda i: mail_label(mails_tries[i]),
            index=st.session_state.sel_mail_idx,
            key="_mail_select_widget",
            on_change=on_mail_change,
        )

        mail = mails_tries[st.session_state.sel_mail_idx]
        a    = mail.get("analyse") or {}
        urg  = a.get("urgence","normale")
        uid  = mail["uid"]

        # Parser PJ automatiquement
        if auto_pj and mail["attachments"] and not mail.get("pj_content"):
            with st.spinner("Lecture des pièces jointes..."):
                mail["pj_content"] = "\n\n".join(
                    PARSER.parse(att) for att in mail["attachments"])
                log(f"PJ auto-parsées pour {mail['subject'][:40]}", "OK")

        st.divider()

        # ── Aperçu mail original ──────────────────────────────────────────
        with st.expander("📨 Mail original", expanded=False):
            st.markdown(f"**De :** {mail['from']}")
            st.markdown(f"**Objet :** {mail['subject']}")
            st.markdown(f"**Date :** {mail['date']}")
            if a:
                emo_a = URGENCY_EMOJI.get(urg,"")
                st.markdown(f"**Analyse :** {emo_a} {urg.upper()} | "
                            f"{a.get('categorie','')} | {a.get('resume','')}")
            if mail["attachments"]:
                st.markdown(f"**PJ :** {', '.join(x['filename'] for x in mail['attachments'])}")
            st.text_area("Corps", mail["body"][:2000], height=180,
                         disabled=True, key=f"orig_body_{uid}")
            if mail.get("pj_content"):
                st.text_area("Contenu PJ", mail["pj_content"][:600],
                             height=100, disabled=True, key=f"pj_prev_{uid}")

        st.divider()

        # ── Conversation IA ───────────────────────────────────────────────
        st.markdown("### 2️⃣  Conversation avec l'IA")

        conv = st.session_state.conversations.get(uid, [])

        # Affichage de l'historique
        if conv:
            st.caption(f"{len(conv)//2} échange(s) — l'IA se souvient de tout ce contexte")
            for msg in conv:
                if msg["role"] == "user":
                    st.markdown(
                        f'<div class="chat-user">👤 <b>Vous</b><br>'
                        f'{msg["content"].replace(chr(10),"<br>")}</div>',
                        unsafe_allow_html=True)
                else:
                    st.markdown(
                        f'<div class="chat-ai">🧠 <b>IA</b><br>'
                        f'{msg["content"].replace(chr(10),"<br>")}</div>',
                        unsafe_allow_html=True)
        else:
            st.info("Pas encore d'échange. Utilisez les raccourcis ou tapez un message.")

        # ── Zone de saisie ────────────────────────────────────────────────
        # Règle Streamlit : on NE PEUT PAS écrire dans session_state[key]
        # après que le widget avec cette key a été rendu.
        # Solution : stocker le texte dans "pending_{uid}" (clé sans widget),
        # le lire comme value= AVANT de créer le widget, puis l'effacer.

        pending_key = f"pending_{uid}"
        if pending_key not in st.session_state:
            st.session_state[pending_key] = ""

        # Raccourcis — s'exécutent AVANT le text_area, donc OK
        col_q1, col_q2, col_q3, col_q4 = st.columns(4)
        with col_q1:
            if st.button("📝 Brouillon", width='stretch', key="q1"):
                st.session_state[pending_key] = \
                    "Génère un brouillon de réponse professionnelle à cet email."
                st.rerun()
        with col_q2:
            if st.button("✂️ Plus court", width='stretch', key="q2"):
                st.session_state[pending_key] = \
                    "Reformule le dernier brouillon en version plus courte."
                st.rerun()
        with col_q3:
            if st.button("📎 Inclure PJ", width='stretch', key="q3"):
                st.session_state[pending_key] = \
                    "Reformule en intégrant des références précises au contenu des pièces jointes."
                st.rerun()
        with col_q4:
            if st.button("🔄 Autre version", width='stretch', key="q4"):
                st.session_state[pending_key] = \
                    "Propose une alternative avec un ton différent."
                st.rerun()

        # Le text_area lit pending_key comme valeur initiale, puis on l'efface
        prefill = st.session_state.get(pending_key, "")
        user_input = st.text_area(
            "Votre message",
            value=prefill,
            placeholder=(
                "Ex : Génère un brouillon de réponse\n"
                "Ex : Reformule plus court et plus formel\n"
                "Ex : Ajoute une référence au fichier joint\n"
                "Ex : Décale le rendez-vous à vendredi"
            ),
            height=90,
            key=f"msg_input_{uid}",
        )
        # Effacer le pending après affichage (le widget a sa propre valeur maintenant)
        if prefill:
            st.session_state[pending_key] = ""

        col_send_ia, col_clear = st.columns([3,1])
        with col_send_ia:
            btn_ia = st.button("🧠 Envoyer à l'IA", type="primary",
                               width='stretch',
                               disabled=not (st.session_state.ollama_ok and user_input.strip()),
                               key="btn_send_ia")
        with col_clear:
            if st.button("🗑️ Effacer conv.", width='stretch', key="clear_conv"):
                st.session_state.conversations.pop(uid, None)
                st.session_state.pop(f"draft_saved_{uid}", None)
                st.session_state.pop(f"draft_last_injected_{uid}", None)
                st.rerun()

        if btn_ia and user_input.strip():
            with st.spinner("L'IA réfléchit..."):
                try:
                    response = chat_avec_historique(
                        mail=mail,
                        user_message=user_input.strip(),
                        model=ollama_model,
                        url=ollama_url,
                        temperature=temp_reply,
                    )
                    # Sauvegarder la dernière réponse IA comme brouillon
                    st.session_state[f"draft_saved_{uid}"] = response
                    # Vider la zone de saisie : supprimer la clé widget
                    # (sera recréée vide au prochain rerun)
                    st.session_state.pop(f"msg_input_{uid}", None)
                    st.session_state[pending_key] = ""
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur Ollama : {e}")
                    log(f"Erreur chat : {e}", "ERR")

        st.divider()

        # ── Éditeur final ─────────────────────────────────────────────────
        st.markdown("### 3️⃣  Éditer et envoyer")

        # Brouillon : dernière réponse IA sauvegardée, sinon template vide
        # On n'écrase PAS si l'user a déjà modifié (clé draft_edit_{uid} dans session)
        draft_ia = st.session_state.get(f"draft_saved_{uid}", "")
        draft_edit_key = f"draft_edit_{uid}"

        # Si une nouvelle réponse IA vient d'arriver, la mettre dans l'éditeur
        if draft_ia and draft_ia != st.session_state.get(f"draft_last_injected_{uid}", ""):
            st.session_state[draft_edit_key] = draft_ia
            st.session_state[f"draft_last_injected_{uid}"] = draft_ia

        # Valeur initiale si rien encore
        if draft_edit_key not in st.session_state:
            sender = mail["from"].split("<")[0].strip().split()
            prenom = sender[0] if sender else "Madame/Monsieur"
            st.session_state[draft_edit_key] = (
                f"Bonjour {prenom},\n\n"
                f"Suite à votre message du {mail['date'][:16]}, "
                f"je vous contacte concernant : {mail['subject']}.\n\n"
                f"[Votre réponse ici]\n\n"
                f"N'hésitez pas à me recontacter pour tout renseignement complémentaire.\n\n"
                f"Cordialement,\n[Votre prénom Nom]\n"
                f"[Votre titre] — Centrale Méditerranée"
            )

        # Destinataire et objet — clés uid-spécifiques
        col_to, col_subj = st.columns(2)
        with col_to:
            to_key = f"to_{uid}"
            if to_key not in st.session_state:
                st.session_state[to_key] = mail["from"]
            st.text_input("À", key=to_key)
            to_addr = st.session_state[to_key]
        with col_subj:
            subj_key = f"subj_{uid}"
            if subj_key not in st.session_state:
                subj_val = mail["subject"]
                if not subj_val.lower().startswith("re:"):
                    subj_val = f"Re: {subj_val}"
                st.session_state[subj_key] = subj_val
            st.text_input("Objet", key=subj_key)
            subj_addr = st.session_state[subj_key]

        st.text_area(
            "✏️ Votre réponse — éditez librement avant d'envoyer",
            height=360,
            key=draft_edit_key,
            help="Pré-rempli avec la dernière réponse de l'IA. Modifiez librement.",
        )
        reponse_finale = st.session_state[draft_edit_key]

        with st.expander("👁️ Aperçu de l'email final", expanded=False):
            st.markdown(f"**De :** {gmail_user}")
            st.markdown(f"**À :** {to_addr}")
            st.markdown(f"**Objet :** {subj_addr}")
            st.divider()
            st.text(reponse_finale)

        if mail.get("brouillon_envoye"):
            st.success("✅ Une réponse a déjà été envoyée pour ce mail.")

        col_s, col_dl, col_clr = st.columns([2,1,1])
        with col_s:
            send_ok = st.button("📤 Envoyer la réponse", type="primary",
                                width='stretch', disabled=not gmail_pass,
                                key="btn_send_email")
        with col_dl:
            st.download_button("💾 Sauvegarder",
                               data=reponse_finale,
                               file_name=f"rep_{mail['subject'][:20].replace(' ','_')}.txt",
                               width='stretch', key="dl_final")
        with col_clr:
            if st.button("🗑️ Réinitialiser", width='stretch', key="clr_draft"):
                st.session_state.pop(draft_edit_key, None)
                st.session_state.pop(f"draft_saved_{uid}", None)
                st.session_state.pop(f"draft_last_injected_{uid}", None)
                st.rerun()

        if not gmail_pass:
            st.warning("⚠️ App Password manquant — renseignez-le dans la sidebar.")

        if send_ok:
            if not reponse_finale.strip():
                st.error("Le message est vide.")
            elif not gmail_pass:
                st.error("App Password manquant.")
            else:
                with st.spinner("Envoi SMTP..."):
                    ok_s, msg_s = envoyer_email(
                        gmail_user, gmail_pass,
                        to=to_addr,
                        subject=subj_addr.replace("Re: Re:", "Re:"),
                        body=reponse_finale)
                if ok_s:
                    mail["brouillon_envoye"] = True
                    conv_now = st.session_state.conversations.get(uid, [])
                    st.session_state.historique.append({
                        "ts":          datetime.now().isoformat(),
                        "to":          to_addr,
                        "subject":     subj_addr,
                        "urgence":     urg,
                        "corps":       reponse_finale[:300],
                        "nb_echanges": len(conv_now)//2,
                    })
                    st.success(f"✅ Email envoyé à {to_addr}")
                    st.balloons()
                    log(f"Email envoyé → {to_addr} | {subj_addr[:40]}", "SMTP")
                    st.rerun()
                else:
                    st.error(f"❌ Erreur SMTP : {msg_s}")

        # Historique des envois récents
        if st.session_state.historique:
            st.divider()
            st.markdown(f"### 📋 Envois récents ({len(st.session_state.historique)})")
            for h in reversed(st.session_state.historique[-5:]):
                n_ia = h.get("nb_echanges", 0)
                st.markdown(
                    f'<div class="metric-card" style="border-left:3px solid #2ecc71">'
                    f'✅ <b>{h["subject"][:50]}</b> → {h["to"][:35]}'
                    f'{"  |  🧠 "+str(n_ia)+" échange(s) IA" if n_ia else ""}'
                    f'<br><small style="color:#888">{h["ts"][:19]}</small>'
                    f'</div>',
                    unsafe_allow_html=True)


# TAB 4 — STATS
# ══════════════════════════════════════════════════════════════
with tab_stats:
    st.subheader("Tableau de bord")

    if st.session_state.triage_done and st.session_state.mails:
        analysed = [m for m in st.session_state.mails if m.get("analyse")]
        df_t = pd.DataFrame([{
            "from":     m["from"][:35],
            "subject":  m["subject"][:50],
            "urgence":  (m.get("analyse") or {}).get("urgence","normale"),
            "categorie":(m.get("analyse") or {}).get("categorie","information"),
            "delai":    (m.get("analyse") or {}).get("delai_reponse","cette_semaine"),
            "nb_pj":    len(m["attachments"]),
            "confiance":(m.get("analyse") or {}).get("confiance",0.5),
            "envoye":   m.get("brouillon_envoye",False),
            "nb_echanges": len(st.session_state.conversations.get(m["uid"],[]))//2,
        } for m in analysed])

        c1,c2,c3,c4 = st.columns(4)
        urg_c = df_t["urgence"].value_counts()
        c1.metric("Total mails", len(df_t))
        c2.metric("🔴 Critiques", urg_c.get("critique",0))
        c3.metric("Réponses envoyées", df_t["envoye"].sum())
        c4.metric("Échanges IA total", df_t["nb_echanges"].sum())

        st.divider()

        urg_order  = ["critique","haute","normale","faible"]
        urg_colors = ["#e94560","#f39c12","#3498db","#95a5a6"]
        del_order  = ["immediat","aujourd_hui","cette_semaine","aucun"]
        del_colors = ["#e94560","#f39c12","#3498db","#95a5a6"]

        urg_vals = df_t["urgence"].value_counts().reindex(urg_order, fill_value=0)
        cat_vals = df_t["categorie"].value_counts()
        del_vals = df_t["delai"].value_counts().reindex(del_order, fill_value=0)

        fig = make_subplots(rows=1, cols=3,
            specs=[[{"type":"domain"},{"type":"xy"},{"type":"xy"}]],
            subplot_titles=("Répartition par urgence",
                            "Répartition par catégorie",
                            "Délai de réponse"))
        fig.add_trace(go.Pie(labels=urg_order, values=urg_vals.values,
            marker=dict(colors=urg_colors), hole=0.45,
            textinfo="label+percent"), row=1, col=1)
        fig.add_trace(go.Bar(x=cat_vals.values, y=cat_vals.index,
            orientation="h", marker_color="#0f3460"), row=1, col=2)
        fig.add_trace(go.Bar(x=del_order, y=del_vals.values,
            marker_color=del_colors, showlegend=False), row=1, col=3)
        fig.update_layout(height=380, template="plotly_white", showlegend=False,
            title=f"Dashboard — {len(df_t)} mails analysés | {ollama_model}")
        st.plotly_chart(fig, width='stretch')

        st.markdown(f"**Confiance moyenne :** {df_t['confiance'].mean():.0%}")
        st.progress(float(df_t["confiance"].mean()))

        pj_df = df_t[df_t["nb_pj"]>0].sort_values("nb_pj", ascending=False)
        if not pj_df.empty:
            st.markdown(f"**{len(pj_df)} mails avec PJ :**")
            st.dataframe(pj_df[["from","subject","urgence","nb_pj","envoye"]],
                         width='stretch')

        if st.session_state.historique:
            st.divider()
            st.subheader(f"Historique des envois ({len(st.session_state.historique)})")
            st.dataframe(pd.DataFrame(st.session_state.historique),
                         width='stretch')
    else:
        st.info("Effectuer le triage pour voir les statistiques.")

# ══════════════════════════════════════════════════════════════
# TAB 5 — LOGS
# ══════════════════════════════════════════════════════════════
with tab_logs:
    st.subheader("Logs en temps réel")
    c1,c2,c3 = st.columns(3)
    c1.metric("Logs", len(st.session_state.logs))
    c2.metric("Ollama", "Actif ✅" if st.session_state.ollama_ok else "Inactif ❌")
    c3.metric("Conversations actives",
              len([k for k,v in st.session_state.conversations.items() if v]))

    if st.button("🔄 Actualiser", key="refresh_logs"):
        st.rerun()

    logs_txt = "\n".join(st.session_state.logs[-80:])
    st.markdown(
        f'<div class="log-container"><pre>{logs_txt}</pre></div>',
        unsafe_allow_html=True)

    if st.session_state.ollama_ok:
        st.success(f"Ollama — {len(st.session_state.ollama_models)} modèle(s) disponible(s)")
        for m in st.session_state.ollama_models:
            star = " ⭐" if m == MODELE_CONSEILLE else ""
            st.write(f"  🦙 {m}{star}")
    else:
        st.error("Ollama non disponible")
        st.code(f"ollama serve\nollama pull {MODELE_CONSEILLE}")

    st.divider()
    st.download_button("💾 Exporter logs",
        data="\n".join(st.session_state.logs),
        file_name=f"logs_{datetime.now().strftime('%Y%m%d_%H%M')}.txt")
