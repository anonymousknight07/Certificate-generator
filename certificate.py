# Supporting functions
import re

# Replace placeholder text in docx files
def docx_replace_regex(doc_obj, regex, replace):
    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex, replace)

def replace_info(doc, placeholder, replace_text):
    reg = re.compile(re.escape(placeholder))
    docx_replace_regex(doc, reg, replace_text)

def replace_participant_name(doc, name):
    placeholder = "{Name Surname}"
    replace_info(doc, placeholder, name)

def replace_event_name(doc, event):
    placeholder = "{EVENT NAME}"
    replace_info(doc, placeholder, event)

def replace_ambassador_name(doc, ambassador):
    placeholder = "{AMBASSADOR NAME}"
    replace_info(doc, placeholder, ambassador)
