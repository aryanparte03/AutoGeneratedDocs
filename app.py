from flask import Flask, render_template, request, send_file
from docx import Document
import os
from datetime import datetime
import zipfile
import tempfile
app = Flask(__name__)

from docx.oxml.ns import qn
from docx.shared import Pt

def set_aptos_font(paragraphs):
    for para in paragraphs:
        for run in para.runs:
            run.font.name = 'Aptos'
            run.font.size = Pt(12)  # You can adjust default size
            # For compatibility (Word will save both):
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), 'Aptos')

def update_word_file(promoter_name, project_name, registration_no):
    from docx import Document
    import os
    from datetime import datetime

    template_path = os.path.join('resources', 'complaintDeclaration.docx')
    output_folder = os.path.join('generated', 'complaintDeclaration')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{registration_no}}': registration_no
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}complaintDeclaration.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path

def generate_no_complaint_file(promoter_name, project_name, registration_no, date):
    from docx import Document
    import os
    from datetime import datetime

    template_path = os.path.join('resources', 'noComplaintsDeclaration.docx')
    output_folder = os.path.join('generated', 'noComplaintsDeclaration')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)
    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{registration_no}}': registration_no,
        '{{date}}': date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_noComplaintsDeclaration.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path

def update_extension_application(date, registration_no, extension_date, promoter_name):
    from datetime import datetime
    from docx import Document
    import os

    # Define template and output folder
    template_path = os.path.join('resources', 'Extension Application under Section 7(3).docx')
    output_folder = os.path.join('generated', 'extension')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)
    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")
    extension_date = datetime.strptime(extension_date, "%Y-%m-%d").strftime("%d-%m-%Y")

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{extension_date}}': extension_date,
        '{{registration_no}}': registration_no,
        '{{date}}': date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Extension Application.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path

def update_project_pert_file(project_name, registration_no, extension_date, promoter_name):
    from docx import Document
    import os
    from datetime import datetime

    # Template and output folder
    template_path = os.path.join('resources', 'projectPert.docx')
    output_folder = os.path.join('generated', 'project_pert')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)
    extension_date = datetime.strptime(extension_date, "%Y-%m-%d").strftime("%d-%m-%Y")

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{extension_date}}': extension_date,
        '{{registration_no}}': registration_no,
        '{{project_name}}': project_name
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_projectPert.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path


def update_cersai_file(promoter_name, project_name, office_address, project_address, date):
    from docx import Document
    import os
    from datetime import datetime

    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")
    # Template and output folder
    template_path = os.path.join('resources','CERSAI Declaration.docx')
    output_folder = os.path.join('generated','CERSAI_Declaration')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{office_address}}': office_address,
        '{{project_address}}': project_address,
        '{{date}}': date
    }
    # Helper function to replace placeholders even if broken into runs
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    # Save file
    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_CERSAI DECLARATION.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path


def update_authorization_letter(promoter_name, project_name, date):
    from docx import Document
    import os
    from datetime import datetime

    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")

    template_path = os.path.join('resources', 'AUTHORIZATION LETTER.docx')
    output_folder = os.path.join('generated', 'Authorization')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{date}}': date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Authorization_Letter.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path

def update_annexure_a(promoter_name, date):
    from docx import Document
    import os
    from datetime import datetime

    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")

    template_path = os.path.join('resources', 'Annexure A.docx')
    output_folder = os.path.join('generated', 'Annexure_A')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{date}}': date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Annexure_a.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path



def update_AffidavitReason_for_Extension(promoter_name, project_name, registration_no):
    from docx import Document
    import os

    template_path = os.path.join('resources', 'Affidavit Reason for Extension.docx')
    output_folder = os.path.join('generated', 'Affidavit')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{registration_no}}': registration_no
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_AffidavitReason_for_Extension.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path


def update_consent_Extension_Tabular(promoter_name, project_name, registration_no, extension_date):
    from docx import Document
    import os
    from datetime import datetime

    template_path = os.path.join('resources', 'Consent for Extension-Tabular.docx')
    output_folder = os.path.join('generated', 'Consent for Extension')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)
    extension_date = datetime.strptime(extension_date, "%Y-%m-%d").strftime("%d-%m-%Y")
    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{registration_no}}': registration_no,
        '{{extension_date}}': extension_date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Consent for Extension.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path


def update_Declaration_For_Extension(promoter_name, project_name, registration_no, extension_date):
    from docx import Document
    import os
    from datetime import datetime

    template_path = os.path.join('resources', 'Declaration For Extension.docx')
    output_folder = os.path.join('generated', 'DeclarationExtension')
    os.makedirs(output_folder, exist_ok=True)

    extension_date = datetime.strptime(extension_date, "%Y-%m-%d").strftime("%d-%m-%Y")

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{registration_no}}': registration_no,
        '{{extension_date}}': extension_date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Declaration For Extension.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path


def update_FormB_File(promoter_name, project_name, office_address, extension_date, project_address):
    from docx import Document
    import os
    from datetime import datetime

    template_path = os.path.join('resources', 'FORM-B.docx')
    output_folder = os.path.join('generated', 'FORM_B')
    os.makedirs(output_folder, exist_ok=True)

    extension_date = datetime.strptime(extension_date, "%Y-%m-%d").strftime("%d-%m-%Y")
    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{office_address}}': office_address,
        '{{extension_date}}': extension_date,
        '{{project_address}}': project_address
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Form_B.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path

def update_FormatA_File(promoter_name, project_name, project_address, account_name, account_number, bank_name, branch_name, ifsc_code, date):
    from docx import Document
    import os
    from datetime import datetime

    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")

    template_path = os.path.join('resources', 'Format A.docx')
    output_folder = os.path.join('generated', 'Format_A')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{project_address}}': project_address,
        '{{account_name}}': account_name,
        '{{account_number}}': account_number,
        '{{bank_name}}': bank_name,
        '{{branch_name}}': branch_name,
        '{{ifsc_code}}': ifsc_code,
        '{{date}}': date

    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Format_A.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path


def update_FormatD_File(promoter_name, project_name, project_address, planning_authority, date):
    from docx import Document
    import os
    from datetime import datetime

    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")

    template_path = os.path.join('resources', 'FORMAT D.docx')
    output_folder = os.path.join('generated', 'Format_D')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{project_name}}': project_name,
        '{{project_address}}': project_address,
        '{{planning_authority}}': planning_authority,
        '{{date}}': date

    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}_Format_D.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path

def update_Consent_Letter(promoter_name, office_address, project_name, registration_no, project_address, extension_date, date):
    from docx import Document
    import os
    from datetime import datetime

    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")
    extension_date = datetime.strptime(extension_date, "%Y-%m-%d").strftime("%d-%m-%Y")
    template_path = os.path.join('resources', 'Consent Letter.docx')
    output_folder = os.path.join('generated', 'ConsentLetter')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{office_address}}': office_address,
        '{{project_name}}': project_name,
        '{{registration_no}}': registration_no,
        '{{project_address}}': project_address,
        '{{extension_date}}': extension_date,
        '{{date}}': date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name}Consent Letter.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path

def update_form1(promoter_name, office_address, project_name, registration_no, as_on_date, date):
    from docx import Document
    import os
    from datetime import datetime

    date = datetime.strptime(date, "%Y-%m-%d").strftime("%d-%m-%Y")

    template_path = os.path.join('resources', 'Form 1.docx')
    output_folder = os.path.join('generated', 'Form1')
    os.makedirs(output_folder, exist_ok=True)

    doc = Document(template_path)

    # Map of placeholders to replacement values
    replacements = {
        '{{promoter_name}}': promoter_name,
        '{{office_address}}': office_address,
        '{{project_name}}': project_name,
        '{{registration_no}}': registration_no,
        '{{as_on_date}}': as_on_date,
        '{{date}}': date
    }

    # Helper function to replace with bold for values only
    def replace_text_with_bold(paragraphs, replacements):
        for para in paragraphs:
            full_text = ''.join(run.text for run in para.runs)
            if not any(ph in full_text for ph in replacements.keys()):
                continue

            new_runs = [(full_text, False)]

            for placeholder, value in replacements.items():
                updated_runs = []
                for text, bold in new_runs:
                    if placeholder in text:
                        parts = text.split(placeholder)
                        for i, part in enumerate(parts):
                            updated_runs.append((part, bold))
                            if i < len(parts) - 1:
                                updated_runs.append((value, True))  # replacement in bold
                    else:
                        updated_runs.append((text, bold))
                new_runs = updated_runs

            para.clear()
            for text, bold in new_runs:
                run = para.add_run(text)
                run.bold = bold

    # Replace in all paragraphs
    replace_text_with_bold(doc.paragraphs, replacements)

    # Replace in all table cells too
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_with_bold(cell.paragraphs, replacements)

    # Set Aptos font everywhere
    set_aptos_font(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_aptos_font(cell.paragraphs)

    sanitized_name = promoter_name.replace(" ", "_")
    filename = f"{sanitized_name} Form1.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    return output_path




from datetime import datetime

def update_form2a(promoter_name, office_address, project_name, registration_no, base_date):
    template_path = os.path.join('resources', 'Form 2A.docx')
    output_folder = os.path.join('generated', 'form2a')
    os.makedirs(output_folder, exist_ok=True)

    # Convert input date to datetime
    base_date = datetime.strptime(base_date, "%Y-%m-%d")
    today = datetime.today()

    # ✅ Determine start financial year based on base_date
    if base_date.month < 4:
        start_year = base_date.year - 1
    else:
        start_year = base_date.year

    # ✅ Determine end financial year based on today's date
    if today.month < 4:
        end_year = today.year - 2
    else:
        end_year = today.year - 1

    generated_files = []

    for i, year in enumerate(range(start_year, end_year + 1), start=1):
        doc = Document(template_path)
        financial_year_str = f"{year}–{year + 1}"  # en dash

        cert_number = f"{i:02d}"  # 2-digit format like 01, 02, ...

        replacements = {
            '{{promoter_name}}': promoter_name,
            '{{office_address}}': office_address,
            '{{project_name}}': project_name,
            '{{registration_no}}': registration_no,
            '{{date}}': today.strftime("%d-%m-%Y"),
            '{{cert_no_upd}}': cert_number,
            '{{financial_year}}': financial_year_str,
            '{{financial_year_date}}': financial_year_str
        }

        def replace_text_with_bold(paragraphs, replacements):
            for para in paragraphs:
                full_text = ''.join(run.text for run in para.runs)
                if not any(ph in full_text for ph in replacements.keys()):
                    continue

                new_runs = [(full_text, False)]

                for placeholder, value in replacements.items():
                    updated_runs = []
                    for text, bold in new_runs:
                        if placeholder in text:
                            parts = text.split(placeholder)
                            for j, part in enumerate(parts):
                                updated_runs.append((part, bold))
                                if j < len(parts) - 1:
                                    updated_runs.append((value, True))
                        else:
                            updated_runs.append((text, bold))
                    new_runs = updated_runs

                para.clear()
                for text, bold in new_runs:
                    run = para.add_run(text)
                    run.bold = bold

        replace_text_with_bold(doc.paragraphs, replacements)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_text_with_bold(cell.paragraphs, replacements)

        # Set Aptos font everywhere
        set_aptos_font(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    set_aptos_font(cell.paragraphs)

        sanitized_name = promoter_name.replace(" ", "_")
        filename = f"{sanitized_name}_Form2A_{financial_year_str}.docx"
        output_path = os.path.join(output_folder, filename)
        doc.save(output_path)
        generated_files.append(output_path)

    return generated_files
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':

        promoter = request.form['promoter_name']
        project = request.form['project_name']
        planning_authority = request.form['planning_authority']
        reg_no = request.form['registration_no']
        # Use today's date instead of taking from form
        date = datetime.today().strftime('%Y-%m-%d')
        extension_date = request.form['extension_date']
        office_add = request.form['office_address']
        project_add = request.form['project_address']
        acc_name = request.form['account_name']
        acc_number = request.form['account_number']
        bank = request.form['bank_name']
        branch = request.form['branch_name']
        ifsc = request.form['ifsc_code']
        as_on_date_raw = request.form['as_on_date']
        as_on_date = datetime.strptime(as_on_date_raw, "%Y-%m-%d").strftime("%d-%m-%Y")
        financial_year_date = request.form['financial_year_date']


        # Generate all 3 files
        file1 = update_word_file(promoter, project, reg_no)
        file2 = generate_no_complaint_file(promoter, project, reg_no, date)
        file3 = update_extension_application(date, reg_no, extension_date, promoter)
        file4 = update_project_pert_file(project, reg_no, extension_date, promoter)
        file5 = update_cersai_file(promoter, project, office_add, project_add, date)
        file6 = update_authorization_letter(promoter, project, date)
        file7 = update_annexure_a(promoter, date)
        file8 = update_AffidavitReason_for_Extension(promoter, project, reg_no)
        file9 = update_consent_Extension_Tabular(promoter, project, reg_no, extension_date)
        file10 = update_Declaration_For_Extension(promoter, project, reg_no, extension_date)
        file11 = update_FormB_File(promoter, project, office_add, extension_date, project_add)
        file12 = update_FormatA_File(promoter, project, project_add, acc_name, acc_number, bank, branch, ifsc, date)
        file13 = update_FormatD_File(promoter, project, project_add, planning_authority, date)
        file14 = update_Consent_Letter(promoter, office_add, project, reg_no, project_add, extension_date, date)
        file15 = update_form1(promoter, office_add, project, reg_no, as_on_date, date)
        form2a_files = update_form2a(promoter, office_add, project, reg_no, financial_year_date)



        # Create zip
        temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
        with zipfile.ZipFile(temp_zip.name, 'w') as zipf:
            zipf.write(file1, os.path.basename(file1))
            zipf.write(file2, os.path.basename(file2))
            zipf.write(file3, os.path.basename(file3))
            zipf.write(file4, os.path.basename(file4))
            zipf.write(file5, os.path.basename(file5))
            zipf.write(file6, os.path.basename(file6))
            zipf.write(file7, os.path.basename(file7))
            zipf.write(file8, os.path.basename(file8))
            zipf.write(file9, os.path.basename(file9))
            zipf.write(file10, os.path.basename(file10))
            zipf.write(file11, os.path.basename(file11))
            zipf.write(file12, os.path.basename(file12))
            zipf.write(file13, os.path.basename(file13))
            zipf.write(file14, os.path.basename(file14))
            zipf.write(file15, os.path.basename(file15))
            for file16 in form2a_files:
                zipf.write(file16, os.path.basename(file16))

        # Use promoter name in zip filename
        sanitized_name = promoter.replace(" ", "_")
        zip_filename = f"{sanitized_name}_Documents.zip"

        return send_file(temp_zip.name, as_attachment=True, download_name=zip_filename)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
