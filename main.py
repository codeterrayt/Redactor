import os
import re
import json
import spacy
from pptx import Presentation
from rapidfuzz import process
from cleanco import basename

# Load the Transformer model
print("Loading Deep Intelligence Model...")
nlp = spacy.load("en_core_web_trf")

class GlobalClientRedactor:
    def __init__(self, mapping_file="client_mapping.json"):
        self.mapping_file = mapping_file
        if os.path.exists(mapping_file):
            with open(mapping_file, 'r') as f:
                data = json.load(f)
                self.client_map = data['map']
                self.counter = data['counter']
        else:
            self.client_map = {} 
            self.counter = 1

    def normalize(self, text):
        base = basename(text)
        base = base.lower().strip()
        base = re.sub(r'[^\w\s]', '', base) # amazon.com -> amazoncom
        return base.strip()

    def get_client_id(self, raw_name):
        norm_name = self.normalize(raw_name)
        if not norm_name or len(norm_name) < 2: return None

        if norm_name in self.client_map:
            return self.client_map[norm_name]

        sorted_keys = sorted(self.client_map.keys(), key=len, reverse=True)
        for existing_norm in sorted_keys:
            if existing_norm in norm_name or norm_name in existing_norm:
                target_id = self.client_map[existing_norm]
                self.client_map[norm_name] = target_id
                return target_id

        if sorted_keys:
            best_match = process.extractOne(norm_name, sorted_keys, score_cutoff=92)
            if best_match:
                target_id = self.client_map[best_match[0]]
                self.client_map[norm_name] = target_id
                return target_id

        client_id = f"[client{self.counter}]"
        self.client_map[norm_name] = client_id
        self.counter += 1
        return client_id

    def redact_pptx(self, file_path, output_dir):
        print(f"Redacting: {os.path.basename(file_path)}")
        prs = Presentation(file_path)
        
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame: continue
                for paragraph in shape.text_frame.paragraphs:
                    # HEAL THE TEXT: Combine all runs to ensure spaCy sees the full context
                    full_para_text = "".join(run.text for run in paragraph.runs)
                    if not full_para_text.strip(): continue

                    doc = nlp(full_para_text)
                    entities = sorted([e for e in doc.ents if e.label_ == "ORG"], 
                                   key=lambda x: x.start_char, reverse=True)

                    for ent in entities:
                        cid = self.get_client_id(ent.text)
                        if cid:
                            # STRATEGY: Replace name in the healed text, then re-distribute to runs
                            # This bypasses the issue of names being split across formatting blocks
                            full_para_text = full_para_text[:ent.start_char] + cid + full_para_text[ent.end_char:]
                    
                    # Wipe existing runs and replace with the redacted full text
                    # (This preserves paragraph structure but may reset run-level formatting like mid-word bolding)
                    if entities:
                        p_text = full_para_text
                        for i, run in enumerate(paragraph.runs):
                            if i == 0:
                                run.text = p_text # Put all text in the first run
                            else:
                                run.text = "" # Clear other runs to avoid duplicates

        output_path = os.path.join(output_dir, f"Redacted_{os.path.basename(file_path)}")
        prs.save(output_path)

    def save_state(self):
        with open(self.mapping_file, 'w') as f:
            json.dump({'map': self.client_map, 'counter': self.counter}, f, indent=4)

# --- Main Execution ---
input_folder = "./source_ppts"
output_folder = "./sanitized_ppts"
os.makedirs(output_folder, exist_ok=True)
os.makedirs(input_folder, exist_ok=True)

redactor = GlobalClientRedactor()
files = sorted([f for f in os.listdir(input_folder) if f.endswith(".pptx")])

for file in files:
    redactor.redact_pptx(os.path.join(input_folder, file), output_folder)

redactor.save_state()
print(f"\nRedaction Complete. Mapping saved.")