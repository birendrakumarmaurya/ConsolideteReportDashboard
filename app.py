import streamlit as st
import os
import json
import re
import random
from docx import Document
# No longer need BytesIO for the template since we read it from disk

# --- CONFIGURATION ---
TEMPLATE_FILENAME = 'PreData.html'  # This file must exist in your GitHub repo

st.set_page_config(page_title="Dashboard Generator", layout="wide")

st.title("📊 Observation Dashboard Compiler")
st.markdown("""
Upload your **Consolidated Word Reports (.docx)**.
The system will automatically use the standard template to generate your dashboard.
""")

# --- LOGIC ---

def get_full_text_from_docx(file_stream):
    doc = Document(file_stream)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = []
            for cell in row.cells:
                if cell.text.strip():
                    row_text.append(cell.text)
            if row_text:
                full_text.append(" | ".join(row_text))
    return '\n'.join(full_text)

def parse_docx(file_stream, filename):
    full_text = get_full_text_from_docx(file_stream)
    
    def extract_section(start, end):
        try:
            pattern = re.compile(f"{start}(.*?){end}", re.DOTALL | re.IGNORECASE)
            match = pattern.search(full_text)
            if match: return match.group(1).strip()
            return ""
        except: return ""

    # Metadata
    name_match = re.search(r"Teacher Observed:[\s:|]*(.*)", full_text, re.IGNORECASE)
    raw_name = name_match.group(1).split('|')[0].strip() if name_match else "Unknown"
    clean_name = raw_name.replace("Ms.", "").replace("Mr.", "").replace("Mrs.", "").strip()
    first_name = clean_name.split()[0] if clean_name else "Unk"

    subject_match = re.search(r"Subject:[\s:|]*(.*)", full_text, re.IGNORECASE)
    subject = subject_match.group(1).split('|')[0].strip() if subject_match else "General"
    
    class_match = re.search(r"Class(?:es)? Observed:[\s:|]*(.*)", full_text, re.IGNORECASE)
    classes = class_match.group(1).split('|')[0].strip() if class_match else ""

    # Sections
    glows_text = extract_section(r"2\.\s*Pedagogical Strengths \(Glows\)", r"3\.\s*Areas for Development")
    glows = [l.strip().lstrip('•-').strip() for l in glows_text.split('\n') if len(l) > 5]
    
    grows_text = extract_section(r"3\.\s*Areas for Development \(Grows\)", r"4\.\s*Action Plan")
    grows = [l.strip().lstrip('•-').strip() for l in grows_text.split('\n') if len(l) > 5]
    
    overview = extract_section(r"1\.\s*Instructional Overview", r"2\.\s*Pedagogical Strengths")
    plan = extract_section(r"4\.\s*Action Plan for Next Cycle", r"5\.\s*Teacher–Student Talk Ratio")

    # Talk Ratio
    talk_ratio = {"teacher": 50, "student": 50} 
    ratios = re.findall(r"(\d{2})%\s*[,|]\s*(\d{2})%", full_text)
    avg_student = 50
    if ratios:
        t_sum, s_sum, count = 0, 0, 0
        for t, s in ratios:
            if int(t) + int(s) == 100:
                t_sum += int(t); s_sum += int(s); count += 1
        if count:
            talk_ratio = {"teacher": round(t_sum/count), "student": round(s_sum/count)}
            avg_student = talk_ratio['student']

    # Rating
    score = 3.5 + (len(glows) * 0.25) - (len(grows) * 0.15)
    if avg_student >= 55: score += 0.2
    if avg_student <= 30: score -= 0.2
    rating = round(max(2.5, min(5.0, score)), 1)

    # Competency
    base_comp = int(rating)
    competency = []
    for _ in range(5):
        val = rating + random.choice([0.5, 0, -0.5])
        competency.append(round(max(2.0, min(5.0, val)), 1))

    colors = ["purple", "cyan", "teal", "rose", "amber", "indigo"]
    
    return {
        "id": re.sub(r'[^a-zA-Z0-9]', '_', raw_name.lower()),
        "name": raw_name, "initials": first_name,
        "subject": subject, "classes": classes,
        "color": colors[len(raw_name) % len(colors)],
        "rating": rating, "overview": overview[:150]+"...",
        "glows": glows[:4], "grows": grows[:3],
        "plan": plan[:100]+"...", "competency": competency, "talk": talk_ratio
    }

# --- UI SECTION ---

st.subheader("1. Upload Files")
uploaded_files = st.file_uploader("Upload Word Reports (.docx)", type=['docx'], accept_multiple_files=True)

if uploaded_files:
    st.subheader("2. Processing")
    if st.button("Generate Dashboard"):
        
        # 1. Check if Template Exists in Repo
        if not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"⚠️ Template file '{TEMPLATE_FILENAME}' not found in the repository. Please add it to GitHub.")
            st.stop()
            
        with st.spinner("Analyzing reports..."):
            teachers_data = []
            for uploaded_file in uploaded_files:
                try:
                    data = parse_docx(uploaded_file, uploaded_file.name)
                    teachers_data.append(data)
                    st.success(f"Parsed: {data['name']} ({data['rating']} ★)")
                except Exception as e:
                    st.error(f"Error parsing {uploaded_file.name}: {e}")
            
            # 2. Read Template from Disk (Repo)
            try:
                with open(TEMPLATE_FILENAME, 'r', encoding='utf-8') as f:
                    template_content = f.read()
            except Exception as e:
                st.error(f"Error reading template file: {e}")
                st.stop()

            # 3. Inject Data
            json_data = json.dumps(teachers_data, indent=4, ensure_ascii=False)
            
            try:
                # Use lambda to safely inject JSON
                new_html = re.sub(
                    r'const teachersData = \[.*?\];', 
                    lambda m: f'const teachersData = {json_data};', 
                    template_content, 
                    flags=re.DOTALL
                )
                
                # Inject Summary Script (Total Count Update)
                summary_script = f"""
                <script>
                    document.addEventListener('DOMContentLoaded', () => {{
                        const countEl = document.getElementById('totalCount') || document.querySelector('.shadow-sm h3');
                        if(countEl) countEl.innerText = "{len(teachers_data)}";
                    }});
                </script>
                </body>
                """
                new_html = new_html.replace('</body>', summary_script)
                
                # DOWNLOAD BUTTON
                st.subheader("3. Download")
                st.download_button(
                    label="Download Dashboard HTML",
                    data=new_html,
                    file_name="Dashboard_Generated.html",
                    mime="text/html"
                )
                
            except Exception as e:
                st.error(f"Error injecting data into HTML: {e}")

elif not uploaded_files:
    st.info("Please upload .docx files to start.")
