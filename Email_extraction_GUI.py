import streamlit as st
import win32com.client
import pandas as pd
import re
import requests
import pdfplumber
import io
from datetime import datetime

# API URL for Ollama (Make sure Ollama is running locally)
OLLAMA_API_URL = "http://localhost:11434/api/generate"

# Define patterns for extracting policy numbers
policy_patterns = [
    r"(?i)polizza\s*(?:RC PROF\.\s*)?N\.?\s*([A-Z0-9-]+)",  # Matches "POLIZZA RC PROF. N. XYZ12345" & "POLIZZA N. XYZ12345"
    r"(?i)POL\.\s*([A-Z0-9-]+)",  # Matches "POL. XYZ12345"
    r"(?i)NR\. POLIZZA\s*([A-Z0-9-]+)",  # Matches "NR. POLIZZA XYZ12345"
    r"(?i)POLIZZA GLOBAL ASSISTANCE N\s*([A-Z0-9-]+)",  # Matches "POLIZZA GLOBAL ASSISTANCE N XYZ12345"
    r"(?i)Polizza:\s*([A-Z0-9-]+)"  # Matches "Polizza: XYZ12345"
]

# Define claim-related keywords
claim_keywords = ["sinistro", "risarcimento", "denuncia", "documenti per un sinistro", "apertura sinistro"]

# Function to extract policy number
def extract_policy_number(text):
    for pattern in policy_patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return "Not Found"

# Function to remove image links
def clean_email_body(body):
    return re.sub(r"<https?://\S+>", "", body)

# Function to extract text from PDF attachments
def extract_attachments_text(message):
    attachment_text = ""
    for attachment in message.Attachments:
        try:
            if attachment.FileName.lower().endswith(".pdf"):
                pdf_data = io.BytesIO(attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102"))
                with pdfplumber.open(pdf_data) as pdf:
                    for page in pdf.pages:
                        attachment_text += page.extract_text() + "\n"
        except Exception:
            pass
    return attachment_text

# Function to extract the latest reply in an email thread
def extract_latest_message(body):
    split_markers = [
        r"On .*? wrote:", r"From: .*?", r"Sent: .*?", r"To: .*?", 
        r"Subject: .*?", r"-----Original Message-----", r"Da: .*?", 
        r"Inviato: .*?", r"A: .*?"
    ]
    pattern = "|".join(split_markers)
    split_body = re.split(pattern, body, maxsplit=1, flags=re.IGNORECASE)
    return split_body[0].strip() if split_body else body.strip()

# Function to check claim-related emails
def is_claim_related(subject, body):
    text = (subject + " " + body).lower()
    return any(keyword in text for keyword in claim_keywords)

# Function to analyze email using Ollama
def analyze_email_with_ollama(subject, body, max_body_chars=2000):
    if len(body) > max_body_chars:
        body = body[:max_body_chars] + " ...[Email truncated]"
    
    prompt = f"""
    Agisci come assistente all'analisi delle e-mail per le richieste di risarcimento assicurativo.
    Determina se il testo seguente riguarda una richiesta di risarcimento e fornisci un riepilogo.
    - Rispondi con "Yes" o "No" all'inizio della risposta.
    - Se rispondi "Yes", aggiungi un breve riassunto.
    - Se non è chiaro, rispondi "No", con un breve riassunto.
    
    Soggetto: {subject}
    Corpo dell'email: {body}
    """
    
    response = requests.post(OLLAMA_API_URL, json={"model": "mistral", "prompt": prompt, "stream": False})
    if response.status_code == 200:
        return response.json()["response"].strip()
    return "Error"

# Streamlit UI
st.title("📩 Email Claim Classifier using Ollama & Outlook")
if st.button("📥 Extract Emails & Analyze Claims"):
    st.write("Fetching emails from Outlook...")
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.Folders["sinistri@bsa-assicurazioni.it"].Folders["Inbox"]
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    
    today_date = datetime.today().strftime("%Y-%m-%d")
    email_data = []
    
    for message in messages:
        try:
            sender_name = message.SenderName
            sender_email = message.Sender.Address if hasattr(message.Sender, "Address") else "Unknown"
            subject = message.Subject
            received_date = message.ReceivedTime.strftime("%Y-%m-%d")
            if not received_date.startswith(today_date):
                continue
            
            full_body = message.Body.strip() if hasattr(message, "Body") else "No Body"
            recent_body = extract_latest_message(full_body)
            cleaned_body = clean_email_body(recent_body)
            attachment_text = extract_attachments_text(message)
            combined_text = cleaned_body + "\n" + attachment_text
            policy_number = extract_policy_number(subject + " " + combined_text)
            
            llm_response = analyze_email_with_ollama(subject, combined_text)
            
            if "yes" in llm_response.lower() or is_claim_related(subject, combined_text):
                email_data.append({
                    "Sender Name": sender_name,
                    "Sender Email": sender_email,
                    "Subject": subject,
                    "Received Date": received_date,
                    "AI Response": llm_response,
                    "Policy Number": policy_number,
                    "Body": recent_body
                })
        except Exception as e:
            st.write(f"❌ Error processing email: {e}")
    
    if email_data:
        df = pd.DataFrame(email_data)
        st.dataframe(df)
        st.download_button("📥 Download CSV", df.to_csv(index=False).encode("utf-8"), "claims_emails.csv", "text/csv")
    else:
        st.write("📭 No claim-related emails found today.")
