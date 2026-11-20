import io
import re
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- Formatting Constants ---
BOLD_TERMS = [
    r"diperakui dengan terma atau syarat",
    r"diperakui",
    r"tidak dapat dipertimbangkan",
    r"kepada"
]

CUSTOM_ENGLISH_WORDS = {
    "accelerator", "accident", "activities", "acuity", "adl", "admin", "administration", "administrative", "administrator", "adolescent", "affiliated",
    "all", "ambulatory", "amp", "anaesthesia", "anaesthesiology", "analog", "and", "angio", "angiografi", "area", "assessment", "association", "asthma", "at",
    "audiologist", "audiometric", "audiometry", "autoclave", "available", "avenue", "bank", "barrier", "based", "basement", "bassinet", "bassinets", "bay",
    "bays", "beam", "bed", "bedded", "bereavement", "billing", "bin", "block", "blood", "bobath", "body", "bone", "booth", "brachytherapy", "breastfeeding",
    "bridge", "bronchoscopy", "bulk", "bunker", "business", "c-arm", "cabinet", "cafeteria", "cardiac", "care", "cashier", "catheterization", "centre", "chair",
    "change", "changing", "chemo", "chemotherapy", "child", "children", "city", "clean", "cleaning", "clerk", "clinic", "clinical", "clinics", "closet", "cold",
    "collection", "colonoscopy", "community", "complex", "cone", "consultation", "consumable", "control", "corridor", "couch", "counselling", "counter",
    "counterpulsation", "crrt", "cssd", "cssu", "cycling", "cytotoxic", "day", "daycare", "dayward", "delivery", "densitometer", "densitometry", "dental",
    "department", "dermatology", "desk", "development", "developmental", "diagnostic", "dialysis", "dietitian", "digital", "dining", "dirty", "discharge",
    "discussion", "diseases", "dispensary", "disposal", "district", "doctor", "doctor's", "domestic", "donor", "double", "double-bedded", "drug", "east", "ecg", "echo",
    "eco", "education", "eeg", "eighth", "electrical", "electrotherapy", "elevator", "emergency", "endoscopy", "enhanced", "entrance", "equipment", "exam", "examination",
    "exec", "executive", "existing", "extended", "extension", "external", "eye", "facilities", "facillities", "family", "female", "fertility", "fetal", "fifth", "file",
    "filing", "first", "floor", "focused", "food", "for", "four", "four-bedded", "fourth", "function", "gamma", "gas", "gastroenterology", "gateway", "general", "grievance",
    "grill", "grilled", "ground", "gym", "gymnasium", "gynaecology", "haemodialysis", "hall", "hand", "hdu", "head", "health", "high", "hold", "holding", "home", "homemaker",
    "hospice", "hospital", "hot", "hybrid", "icu", "imaging", "in", "infectious", "inhouse", "inpatient", "intensity", "intensive", "intermediate", "international", "intra",
    "intraoperative", "isolation", "it", "janitor", "janitor's", "joint", "kitchen", "knife", "lab", "laboratory", "labour", "laser", "ldr", "level", "licensed", "lift", "linear", "linen", "link", "lobby", "local", "locker", "lounge", "lower", "lung", "machine", "machines", "major", "male", "mall", "mammogram", "mammography", "management", "manager", "maternal", "maternity", "matron", "maxillofacial", "medical", "medication", "medicine", "meeting", "mezzanine", "mineral", "minor", "mobile", "multidisiplin", "multipurpose", "negative", "nephrologist", "nephrology", "nerve", "neurology", "new", "ninth", "non-hepatitis", "nurse", "nursery", "nursing", "observation", "obstetrics", "occupational", "of", "off", "office", "officers", "oncology", "open", "operating", "operation", "ophthalmology", "optometrist", "optometry", "oral", "orthopaedic", "outpatient", "outsourced", "over", "packing", "paediatric", "paediatrician", "paediatrics", "paeds", "palliative", "pantry", "paraffin", "park", "pass", "patient", "pediatric", "pendant", "pharmacy", "phlebotomy", "physical", "physician's", "physio", "physiotherapy", "plastic", "play", "portable", "positive", "pre-function", "pre-operating", "prep", "preparation", "pressure", "private", "procedure", "processing", "psychiatry", "public", "quarantine", "radiation", "radiography", "radiologist", "radiology", "radiotherapy", "readily", "receiving", "reception", "reconstitution", "record", "recovery", "refraction", "refuge", "registration", "rehab", "rehabilitation", "relokasi", "reporting", "reprocessor", "reproductive", "respiratory", "respite", "resusitasi", "retail", "robotic", "room", "rooming-in", "rooms", "sample", "satellite", "scanner", "scientific", "scn", "scope", "screen", "screening", "second", "sensory", "serum", "server", "services", "settlement", "seventh", "silent", "simulator", "single", "single-bedded", "sister", "sixth", "sleep", "sluice", "society", "soiled", "special", "specialist", "specialists", "speech", "splinting", "sports", "staff", "stair", "station", "sterilization", "storage", "store", "stress", "stretcher", "study", "sub", "suite", "supervisor", "supply", "support", "surgery", "surgical", "table", "taking", "tea", "teaching", "technical", "tenth", "test", "theatre", "therapy", "third", "three", "three-bedded", "times", "to", "toilet", "topical", "tower", "traction", "training", "treadmill", "treadmills", "treatment", "triage", "trolley", "ultrasound", "uro-gynaecology", "uroflowmetry", "utility", "vacant", "vascular", "vertigo", "vestibular", "vision", "visiting", "visitor", "visual", "volunteer", "wait", "waiting", "ward", "waste", "wax", "wheelchair", "wing", "with", "women's", "work", "x-ray", "airlock", "auditory", "autologous", "baby", "basis", "brainstem", "breast", "cancers", "cell", "chamber", "clinically", "committee", "disabled", "electrophysiology", "endocrine", "enhancement", "ethical", "execute", "formula", "green", "hearing", "hyperbaric", "immune", "institutional", "main", "marketing", "multiplace", "nourishment", "only", "oxygen", "policies", "procedures", "progress", "protocol", "proven", "psychiatric", "ramp", "red", "response", "standard", "stem", "therapeutic", "trial", "value", "wardrobe", "yellow", "zone", "gynae-oncology", "pet/ct", "spect/ct/pet"
}

# --- Extract Text and Lists Helper ---
def extract_cell_text(cell):
    """Reusable helper to extract text while tracking dynamic XML numbering."""
    cell_text_parts = []
    list_counters = {}
    last_ilvl = -1
    
    for p in cell.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        is_bullet = False
        is_number = False
        ilvl = 0
        
        style_name = p.style.name if p.style else ""
        if 'Bullet' in style_name:
            is_bullet = True
            
        num_pr = p._element.xpath('.//w:numPr')
        if num_pr:
            ilvl_nodes = p._element.xpath('.//w:ilvl/@w:val')
            if ilvl_nodes:
                ilvl = int(ilvl_nodes[0])
            if not is_bullet:
                is_number = True
        elif 'Number' in style_name or 'List' in style_name:
            if not is_bullet:
                is_number = True
            if p.paragraph_format.left_indent:
                ilvl = int(p.paragraph_format.left_indent.pt // 18)

        indent_space = "    " * ilvl
        
        if re.match(r'^([a-zA-Z0-9]+[\.\)]|[-•])\s', text):
            if re.match(r'^1[\.\)]\s', text):
                list_counters = {}
            cell_text_parts.append(f"{indent_space}{text}")
            continue
            
        if is_bullet:
            cell_text_parts.append(f"{indent_space}• {text}")
        elif is_number:
            if ilvl < last_ilvl:
                for level in list(list_counters.keys()):
                    if level > ilvl:
                        list_counters[level] = 0
                        
            list_counters[ilvl] = list_counters.get(ilvl, 0) + 1
            last_ilvl = ilvl
            val = list_counters[ilvl]
            
            if ilvl == 0:
                prefix = f"{val}."
            elif ilvl == 1:
                prefix = f"{chr(96 + val) if val <= 26 else val}."
            elif ilvl == 2:
                roman = ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"]
                prefix = f"{roman[val]}." if val <= 10 else f"{val}."
            else:
                prefix = f"{val}."
                
            cell_text_parts.append(f"{indent_space}{prefix} {text}")
        else:
            cell_text_parts.append(text)
                
    return "\n".join(cell_text_parts)


# --- Extractor ---
def extract_field_data(doc):
    data = {
        "id_permohonan": "",
        "nama_premis": "",
        "alamat_premis": "",
        "nama_oyb": "",
        "kp_oyb": "",
        "kelayakan_oyb": "",
        "tarikh_permohonan": "",
        "tarikh_luput": "",
        "no_fail": "",
        "pegawai_proses": "",
        "isu_ulasan": ""
    }

    # Generate sequential text view of document while preserving table structures
    sequential_lines = []
    for element in doc.element.body.iterchildren():
        if isinstance(element, CT_P):
            p = Paragraph(element, doc)
            if p.text.strip():
                sequential_lines.append(p.text.strip())
        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            for row in table.rows:
                row_cells = []
                seen_cells = set()
                for cell in row.cells:
                    if cell._tc not in seen_cells:
                        seen_cells.add(cell._tc)
                        cell_text = extract_cell_text(cell)
                        if cell_text.strip():
                            row_cells.append(cell_text.strip())
                if row_cells:
                    sequential_lines.append(" | ".join(row_cells))
                    
    full_text = "\n".join(sequential_lines)

    # Regex Helper
    def extract(pattern, text, default=""):
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        return match.group(1).strip() if match else default

    # Extraction Logic 
    data["id_permohonan"] = extract(r"Nombor ID Permohonan\s*(?:[:|]\s*)([^\n]+)", full_text)
    
    data["nama_premis"] = extract(r"Nama Premis\s*\|\s*([^\n]+)", full_text)
    if not data["nama_premis"]:
        data["nama_premis"] = extract(r"Nama Premis\s*[:]\s*([^\n]+)", full_text)
        
    alamat_match = re.search(r"Alamat Premis\s*(?:[:|]\s*)(.*?)(?=\n\s*(?:Emel|No\. Telefon))", full_text, re.DOTALL | re.IGNORECASE)
    if alamat_match:
        addr_clean = alamat_match.group(1).strip().replace('\n', ' ')
        data["alamat_premis"] = re.sub(r'\s+', ' ', addr_clean)
        
    # Extract OYB details specifically from "Butiran Orang Yang Bertanggungjawab:"
    butiran_oyb_section = re.search(r'Butiran Orang Yang Bertanggungjawab:(.*?)(?=Tarikh lesen bermula|No\.\s*Fail|Pegawai Proses|\Z)', full_text, re.DOTALL | re.IGNORECASE)
    
    if butiran_oyb_section:
        oyb_text = butiran_oyb_section.group(1)
        
        # [UPDATED] Nama: Stops reading when it hits "No. K/P:"
        raw_nama = extract(r'Nama:\s*(.*?)(?=\s*No\.\s*K/?P:|\n|$)', oyb_text)
        data["nama_oyb"] = re.sub(r'\s*\(=.*?\)', '', raw_nama).strip()
        
        # [UPDATED] No. K/P: Stops reading when it hits "Kelayakan:"
        data["kp_oyb"] = extract(r'No\.\s*K/?P:\s*(.*?)(?=\s*Kelayakan:|\n|$)', oyb_text)
        
        # [UPDATED] Kelayakan: Stops reading when it hits "Tarikh" or "No. Fail"
        raw_kelayakan = extract(r'Kelayakan:\s*(.*?)(?=\s*Tarikh|\s*No\.\s*Fail|\n|$)', oyb_text)
        data["kelayakan_oyb"] = re.sub(r'\s*\(=.*?\)', '', raw_kelayakan).strip()
    else:
        # Fallback to table if Butiran section is missing
        oyb_section = re.search(r'Orang Yang Bertanggungjawab \(OYB\)(.*?)(?:Ketua Jabatan|Isu / Ulasan)', full_text, re.DOTALL | re.IGNORECASE)
        if oyb_section:
            oyb_table_text = oyb_section.group(1)
            data["nama_oyb"] = extract(r'\nNama\s*(?:[:|]\s*)(.*?)(?=\s*No\. Kad Pengenalan|\s*No\.\s*K/?P|\n|$)', oyb_table_text)
            data["kp_oyb"] = extract(r'No\. Kad Pengenalan\s*(?:[:|]\s*)(.*?)(?=\s*Jantina|\s*Bangsa|\n|$)', oyb_table_text)
        data["kelayakan_oyb"] = extract(r'Kelayakan:\s*(.*?)(?=\s*Tarikh\s+lesen\s+bermula|\s*Tarikh\s+permohonan|\s*No\.\s+Fail|\n|$)', full_text)
        
    data["tarikh_permohonan"] = extract(r"Tarikh terima fi pemprosesan:\s*([^\n(]+)", full_text).strip()
    
    # Strictly extract the date inside the [] for Tarikh Luput
    data["tarikh_luput"] = extract(r"luput lesen sedia ada iaitu pada\s*\[([^\]]+)\]", full_text)
    if not data["tarikh_luput"]:
         data["tarikh_luput"] = extract(r"Tarikh Sah Lesen\s*\|\s*[^-]+-\s*([^\n]+)", full_text)
         
    # No. Fail: Stops reading when it hits "Pegawai Proses" or "Tarikh"
    data["no_fail"] = extract(r"No\.\s*Fail:\s*(.*?)(?=\s*Pegawai\s+Proses:|\s*Tarikh|\n|$)", full_text)
    
    # Pegawai Proses: Stops reading when it hits "Tarikh"
    data["pegawai_proses"] = extract(r"Pegawai\s*Proses:\s*(.*?)(?=\s*Tarikh|\n|$)", full_text)
    
    # Extract only the specific "untuk kapasiti... diperakui" phrase
    kapasiti_match = re.search(r'Permohonan pembaharuan lesen hospital swasta.*?(untuk kapasiti.*?diperakui)', full_text, re.IGNORECASE)
    if kapasiti_match:
        kapasiti_text = kapasiti_match.group(1).strip()
        data["isu_ulasan"] = f"{kapasiti_text.capitalize()}\nTarikh lesen bermula dari tarikh luput lesen sedia ada."
    else:
        data["isu_ulasan"] = "Tarikh lesen bermula dari tarikh luput lesen sedia ada."

    return data


# --- Formatters ---
def format_styled_text_run(paragraph, text):
    if not text:
        return

    bold_pattern = f"({'|'.join(BOLD_TERMS)})"
    tokens = re.split(bold_pattern, text, flags=re.IGNORECASE)

    for token in tokens:
        if not token:
            continue
            
        is_bold = any(re.fullmatch(term, token, re.IGNORECASE) for term in BOLD_TERMS)
        sub_tokens = re.split(r'(\s+|[.,;:]+)', token)
        
        for sub in sub_tokens:
            if not sub:
                continue
            
            run = paragraph.add_run(sub)
            run.font.name = 'Arial'
            run.font.size = Pt(9)
            if is_bold:
                run.bold = True
            
            clean_word = re.sub(r'[^\w\-\']', '', sub).lower()
            if clean_word in CUSTOM_ENGLISH_WORDS and not is_bold:
                run.italic = True

def write_formatted_cell(cell, text):
    """Writes formatted string data into a cell while preserving hanging indents."""
    if not text:
        return
    lines = text.split('\n')
    first_paragraph_used = False
    
    for line in lines:
        if not line.strip():
            if first_paragraph_used:
                cell.add_paragraph()
            continue
        
        if not first_paragraph_used:
            p = cell.paragraphs[0]
            first_paragraph_used = True
        else:
            p = cell.add_paragraph()
        
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        leading_spaces = len(line) - len(line.lstrip())
        indent_level = leading_spaces // 4
        clean_text = line.lstrip() 
        
        list_match = re.match(r'^([a-zA-Z0-9]+[\.\)]|[•\-])\s', clean_text)
        
        if list_match:
            p.paragraph_format.left_indent = Pt(18 * indent_level + 18)
            p.paragraph_format.first_line_indent = Pt(-18)
        elif indent_level > 0:
            p.paragraph_format.left_indent = Pt(18 * indent_level)
        
        format_styled_text_run(p, clean_text)


def add_field(p, label, value):
    """Helper to cleanly inject bolded label + normal text into a paragraph."""
    if label:
        run_lbl = p.add_run(label)
        run_lbl.font.name = 'Arial'
        run_lbl.font.size = Pt(9)
        run_lbl.bold = True
    if value:
        run_val = p.add_run(f"{value}\n")
        run_val.font.name = 'Arial'
        run_val.font.size = Pt(9)


def build_summary_doc(records):
    doc = Document()
    
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_titles = ["SENARAI PERMOHONAN", "ISU/ULASAN"]
    
    for i, title in enumerate(hdr_titles):
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(title)
        run.font.name = 'Arial'
        run.font.size = Pt(9)
        run.bold = True

    for data in records:
        row_cells = table.add_row().cells
        
        # --- Left Column: Maklumat Permohonan ---
        p_left = row_cells[0].paragraphs[0]
        
        # Premis Nama & Alamat
        run_premis = p_left.add_run(f"{data.get('nama_premis', '').upper()}\n")
        run_premis.font.name = 'Arial'
        run_premis.font.size = Pt(9)
        run_premis.bold = True

        run_alamat = p_left.add_run(f"{data.get('alamat_premis', '')}\n")
        run_alamat.font.name = 'Arial'
        run_alamat.font.size = Pt(9)

        # Field injections
        add_field(p_left, "No. ID Permohonan: ", data.get("id_permohonan", ""))
        add_field(p_left, "Nama OYB: ", data.get("nama_oyb", ""))
        add_field(p_left, "No. K/P: ", data.get("kp_oyb", ""))
        add_field(p_left, "Kelayakan: ", data.get("kelayakan_oyb", ""))
        add_field(p_left, "Tarikh permohonan dikemukakan: ", data.get("tarikh_permohonan", ""))
        add_field(p_left, "Tarikh luput lesen: ", data.get("tarikh_luput", ""))
        add_field(p_left, "No. Fail: ", data.get("no_fail", ""))
        add_field(p_left, "Pegawai Proses: ", data.get("pegawai_proses", ""))

        # --- Right Column: Keputusan (Isu/Ulasan) ---
        write_formatted_cell(row_cells[1], data.get('isu_ulasan', ''))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit UI ---
st.set_page_config(page_title="Document Extractor", page_icon="📄", layout="centered")

st.title("📄 Pindaan to Lampiran 2 Extractor")
st.write("Upload your Word documents (.docx) below to extract data into a 2-column Word table.")

uploaded_files = st.file_uploader(
    "Choose Word files (.docx only)",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Process Documents", type="primary"):
        with st.spinner("Extracting data from files..."):
            extracted_records = []
            
            for file in uploaded_files:
                doc = Document(file)
                record = extract_field_data(doc)
                extracted_records.append(record)

            output_buffer = build_summary_doc(extracted_records)

        st.success("Extraction Complete!")
        
        st.download_button(
            label="📥 Download Extracted Summary (.docx)",
            data=output_buffer,
            file_name="Lampiran_2_MyBRP_raw.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
