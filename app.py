import os
import uuid
import re
import tempfile
from flask import Flask, render_template, request, send_file, jsonify, url_for
import fitz
import docx

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.secret_key = 'pdftoxml-secret-key'

UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'pdftoxml-uploads')

def ensure_upload_dir():
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

def extract_text_from_docx(docx_path):
    doc = docx.Document(docx_path)
    text = "\n".join(p.text for p in doc.paragraphs)
    return text

def guess_title(text):
    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
    for line in lines[:20]:
        if len(line) > 20 and len(line) < 300:
            return line
    return ""

def guess_abstract(text):
    patterns = [r'(?:resumen|abstract)[:\s]*\n*(.*?)(?:\n\n|\n(?:introducci[óo]n|palabras clave|keywords|introduction))', r'(?:resumen)[:\s]*(.*?)(?:\n\n)']
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE | re.DOTALL)
        if m:
            return m.group(1).strip()[:2000]
    return ""

def guess_keywords(text, lang='es'):
    patterns_es = [r'palabras clave[:\s]*(.*?)(?:\n\n|\n(?:keywords|introducci[óo]n|abstract))', r'palabras clave[:\s]*(.*?)$']
    patterns_en = [r'keywords[:\s]*(.*?)(?:\n\n|\n(?:palabras clave|introduction|abstract))', r'keywords[:\s]*(.*?)$']
    patterns = patterns_es if lang == 'es' else patterns_en
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE | re.DOTALL)
        if m:
            kw = m.group(1).strip()
            return [k.strip() for k in re.split(r'[;,\n]', kw) if k.strip()]
    return []

def guess_sections(text):
    sections = []
    heading_pattern = re.compile(r'^(introducci[óo]n|justificaci[óo]n|planteamiento\s+del\s+problema|objetivos|metodolog[íi]a|resultados|conclusiones|discusi[óo]n|referencias|bibliograf[íi]a|m[eé]todos|material\s+y\s+m[eé]todos)\s*$', re.IGNORECASE | re.MULTILINE)
    parts = heading_pattern.split(text)
    if len(parts) <= 1:
        return [{"title": "", "content": text[:5000]}]
    sections.append({"title": "Introducción", "content": parts[0].strip()[:2000]})
    i = 1
    while i < len(parts) - 1:
        title = parts[i].strip().capitalize()
        content = parts[i+1].strip()[:5000] if i+1 < len(parts) else ""
        sections.append({"title": title, "content": content})
        i += 2
    return sections

def guess_references(text):
    refs = []
    ref_markers = [
        r'(?:^|\n)\s*(?:referencias|bibliograf[íi]a|references|bibliography)\s*[\:\n]',
        r'(?:^|\n)\s*(?:referencias|bibliograf[íi]a|references|bibliography)\s*$'
    ]
    ref_section = ""
    for marker in ref_markers:
        m = re.search(marker, text, re.IGNORECASE)
        if m:
            ref_section = text[m.end():]
            break

    if not ref_section or len(ref_section) < 20:
        return refs

    ref_section = ref_section.strip()
    lines = ref_section.split('\n')
    buffer = ""
    for line in lines:
        line = line.strip()
        if not line:
            continue
        starts_with_num = re.match(r'^\[?\d+[\].]?\s+', line)
        starts_with_ref = re.match(r'^[A-Z][A-Za-zÁÉÍÓÚÑáéíóúñ]+[,]?\s', line)
        if starts_with_num or (starts_with_ref and buffer and len(buffer) > 30):
            if buffer and len(buffer) > 30:
                refs.append(buffer.strip())
            buffer = line
        else:
            if buffer:
                buffer += " " + line
            else:
                buffer = line

    if buffer and len(buffer) > 30:
        refs.append(buffer.strip())

    seen = set()
    unique = []
    for r in refs:
        r_clean = re.sub(r'^\[?\d+[\].]?\s*', '', r).strip()[:80]
        if r_clean and r_clean not in seen:
            seen.add(r_clean)
            unique.append(r)

    return unique[:20]

def guess_authors(text):
    lines = text.strip().split('\n')
    for i, line in enumerate(lines[:30]):
        line_clean = line.strip()
        if re.search(r'Clark.*Valenzuela|Zayas.*Campas|autor|author', line_clean, re.IGNORECASE):
            continue
        if re.match(r'^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,4}$', line_clean) and len(line_clean) < 80:
            if i > 0 and len(line_clean) > 10:
                continue
        possible = re.findall(r'([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,3}(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)?)', text[:2000])
        if possible:
            return list(set(possible))[:3]
    return []

def generate_jats_xml(data):
    article_id = data.get('article_id', '1')
    doi = data.get('doi', f'10.46589/riasf.v1i43.{article_id}')
    title_es = data.get('title_es', '')
    title_en = data.get('title_en', '')
    abstract_es = data.get('abstract_es', '')
    abstract_en = data.get('abstract_en', '')
    keywords_es = data.get('keywords_es', '').split(',')
    keywords_en = data.get('keywords_en', '').split(',')
    body_html = data.get('body', '')
    authors_count = int(data.get('authors_count', '0'))
    refs_count = int(data.get('refs_count', '0'))

    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<!DOCTYPE article\n'
    xml += '  PUBLIC "-//NLM//DTD JATS (Z39.96) Journal Publishing DTD v1.1 20151215//EN" "https://jats.nlm.nih.gov/publishing/1.1/JATS-journalpublishing1.dtd">\n'
    xml += '<article article-type="research-article" dtd-version="1.1" specific-use="sps-1.9" xml:lang="es" xmlns:mml="http://www.w3.org/1998/Math/MathML" xmlns:xlink="http://www.w3.org/1999/xlink">\n'
    xml += '\t<front>\n'
    xml += '\t\t<journal-meta>\n'
    xml += '\t\t\t<journal-id journal-id-type="publisher-id">riasf</journal-id>\n'
    xml += '\t\t\t<journal-title-group>\n'
    xml += '\t\t\t\t<journal-title>Revista de Investigación Académica sin Frontera</journal-title>\n'
    xml += '\t\t\t\t<abbrev-journal-title abbrev-type="publisher">riasf</abbrev-journal-title>\n'
    xml += '\t\t\t</journal-title-group>\n'
    xml += '\t\t\t<issn pub-type="ppub">2007-8870</issn>\n'
    xml += '\t\t\t<publisher>\n'
    xml += '\t\t\t\t<publisher-name>Universidad de Sonora</publisher-name>\n'
    xml += '\t\t\t</publisher>\n'
    xml += '\t\t</journal-meta>\n'
    xml += '\n'
    xml += '\t\t<article-meta>\n'
    xml += f'\t\t\t<article-id pub-id-type="doi">{doi}</article-id>\n'
    xml += f'\t\t\t<article-id pub-id-type="other">{article_id}</article-id>\n'
    xml += '\n'
    xml += '\t\t\t<article-categories>\n'
    xml += '\t\t\t\t<subj-group subj-group-type="heading">\n'
    xml += '\t\t\t\t\t<subject>Artículos</subject>\n'
    xml += '\t\t\t\t</subj-group>\n'
    xml += '\t\t\t</article-categories>\n'
    xml += '\n'
    xml += '\t\t\t<title-group>\n'
    xml += f'\t\t\t\t<article-title>{escape_xml(title_es)}</article-title>\n'
    if title_en:
        xml += '\t\t\t\t<trans-title-group xml:lang="en">\n'
        xml += f'\t\t\t\t\t<trans-title>{escape_xml(title_en)}</trans-title>\n'
        xml += '\t\t\t\t</trans-title-group>\n'
    xml += '\t\t\t</title-group>\n'

    if authors_count > 0:
        xml += '\n'
        xml += '\t\t\t<contrib-group>\n'
        for i in range(1, authors_count + 1):
            name = data.get(f'author_name_{i}', '')
            surname = data.get(f'author_surname_{i}', '')
            given = data.get(f'author_given_{i}', '')
            orcid = data.get(f'author_orcid_{i}', '')
            email = data.get(f'author_email_{i}', '')
            aff_id = data.get(f'author_aff_{i}', '1')
            if not name and not surname:
                continue
            xml += '\t\t\t\t<contrib contrib-type="author">\n'
            if orcid:
                xml += f'\t\t\t\t\t<contrib-id contrib-id-type="orcid">{escape_xml(orcid)}</contrib-id>\n'
            xml += '\t\t\t\t\t<name>\n'
            if surname:
                xml += f'\t\t\t\t\t\t<surname>{escape_xml(surname)}</surname>\n'
            else:
                parts = name.strip().split(' ', 1)
                if len(parts) > 1:
                    xml += f'\t\t\t\t\t\t<surname>{escape_xml(parts[0])}</surname>\n'
                    xml += f'\t\t\t\t\t\t<given-names>{escape_xml(parts[1])}</given-names>\n'
                else:
                    xml += f'\t\t\t\t\t\t<given-names>{escape_xml(name)}</given-names>\n'
            if not surname and not given:
                pass
            elif surname and given:
                pass
            else:
                pass
            if given and surname:
                xml += f'\t\t\t\t\t\t<surname>{escape_xml(surname)}</surname>\n'
                xml += f'\t\t\t\t\t\t<given-names>{escape_xml(given)}</given-names>\n'
            xml += '\t\t\t\t\t</name>\n'
            xml += f'\t\t\t\t\t<xref ref-type="aff" rid="aff{aff_id}"><sup>{aff_id}</sup></xref>\n'
            xml += '\t\t\t\t</contrib>\n'
        xml += '\t\t\t</contrib-group>\n'

    aff_count = int(data.get('aff_count', '0'))
    if aff_count > 0:
        for i in range(1, aff_count + 1):
            aff_text = data.get(f'aff_text_{i}', '')
            if not aff_text:
                continue
            xml += '\n'
            xml += f'\t\t\t<aff id="aff{i}">\n'
            xml += f'\t\t\t\t<label>{i}</label>\n'
            xml += f'\t\t\t\t<institution content-type="original">{escape_xml(aff_text)}</institution>\n'
            xml += f'\t\t\t\t<institution content-type="normalized">{escape_xml(aff_text)}</institution>\n'
            xml += f'\t\t\t\t<institution content-type="orgname">{escape_xml(aff_text)}</institution>\n'
            xml += '\t\t\t\t<country country="MX">Mexico</country>\n'
            email_aff = data.get(f'aff_email_{i}', '')
            if email_aff:
                xml += f'\t\t\t\t<email>{escape_xml(email_aff)}</email>\n'
            xml += '\t\t\t</aff>\n'

    pub_day = data.get('pub_day', '30')
    pub_month = data.get('pub_month', '06')
    pub_year = data.get('pub_year', '2025')
    volume = data.get('volume', '1')
    issue = data.get('issue', '43')
    elocation = data.get('elocation', f'e{article_id}')

    xml += '\n'
    xml += '\t\t\t<pub-date date-type="pub" publication-format="electronic">\n'
    xml += f'\t\t\t\t<day>{pub_day}</day>\n'
    xml += f'\t\t\t\t<month>{pub_month}</month>\n'
    xml += f'\t\t\t\t<year>{pub_year}</year>\n'
    xml += '\t\t\t</pub-date>\n'
    xml += '\n'
    xml += f'\t\t\t<volume>{volume}</volume>\n'
    xml += f'\t\t\t<issue>{issue}</issue>\n'
    xml += f'\t\t\t<elocation-id>{escape_xml(elocation)}</elocation-id>\n'

    rec_day = data.get('rec_day', '16')
    rec_month = data.get('rec_month', '03')
    rec_year = data.get('rec_year', '2025')
    acc_day = data.get('acc_day', '16')
    acc_month = data.get('acc_month', '06')
    acc_year = data.get('acc_year', '2025')

    xml += '\n'
    xml += '\t\t\t<history>\n'
    xml += '\t\t\t\t<date date-type="received">\n'
    xml += f'\t\t\t\t\t<day>{rec_day}</day>\n'
    xml += f'\t\t\t\t\t<month>{rec_month}</month>\n'
    xml += f'\t\t\t\t\t<year>{rec_year}</year>\n'
    xml += '\t\t\t\t</date>\n'
    xml += '\t\t\t\t<date date-type="accepted">\n'
    xml += f'\t\t\t\t\t<day>{acc_day}</day>\n'
    xml += f'\t\t\t\t\t<month>{acc_month}</month>\n'
    xml += f'\t\t\t\t\t<year>{acc_year}</year>\n'
    xml += '\t\t\t\t</date>\n'
    xml += '\t\t\t\t<date date-type="pub">\n'
    xml += f'\t\t\t\t\t<day>{pub_day}</day>\n'
    xml += f'\t\t\t\t\t<month>{pub_month}</month>\n'
    xml += f'\t\t\t\t\t<year>{pub_year}</year>\n'
    xml += '\t\t\t\t</date>\n'
    xml += '\t\t\t</history>\n'

    xml += '\n'
    xml += '\t\t\t<permissions>\n'
    xml += '\t\t\t\t<license license-type="open-access" xlink:href="https://creativecommons.org/licenses/by-nc/4.0/" xml:lang="es">\n'
    xml += '\t\t\t\t\t<license-p>Este es un artículo publicado en acceso abierto bajo una licencia Creative Commons.</license-p>\n'
    xml += '\t\t\t\t</license>\n'
    xml += '\t\t\t</permissions>\n'

    if abstract_es:
        xml += '\n'
        xml += '\t\t\t<abstract>\n'
        xml += '\t\t\t\t<title>Resumen</title>\n'
        xml += f'\t\t\t\t<p>{escape_xml(abstract_es)}</p>\n'
        xml += '\t\t\t</abstract>\n'

    if abstract_en:
        xml += '\n'
        xml += '\t\t\t<trans-abstract xml:lang="en">\n'
        xml += '\t\t\t\t<title>Abstract</title>\n'
        xml += f'\t\t\t\t<p>{escape_xml(abstract_en)}</p>\n'
        xml += '\t\t\t</trans-abstract>\n'

    if keywords_es and keywords_es[0]:
        xml += '\n'
        xml += '\t\t\t<kwd-group xml:lang="es">\n'
        xml += '\t\t\t\t<title>Palabras clave:</title>\n'
        for kw in keywords_es:
            kw = kw.strip()
            if kw:
                xml += f'\t\t\t\t<kwd>{escape_xml(kw)}</kwd>\n'
        xml += '\t\t\t</kwd-group>\n'

    if keywords_en and keywords_en[0]:
        xml += '\n'
        xml += '\t\t\t<kwd-group xml:lang="en">\n'
        xml += '\t\t\t\t<title>Keywords:</title>\n'
        for kw in keywords_en:
            kw = kw.strip()
            if kw:
                xml += f'\t\t\t\t<kwd>{escape_xml(kw)}</kwd>\n'
        xml += '\t\t\t</kwd-group>\n'

    xml += '\n'
    xml += '\t\t\t<counts>\n'
    xml += '\t\t\t\t<fig-count count="0"/>\n'
    xml += '\t\t\t\t<table-count count="0"/>\n'
    xml += '\t\t\t\t<equation-count count="0"/>\n'
    xml += f'\t\t\t\t<ref-count count="{refs_count}"/>\n'
    xml += '\t\t\t\t<page-count count="0"/>\n'
    xml += '\t\t\t</counts>\n'
    xml += '\t\t</article-meta>\n'
    xml += '\t</front>\n'

    xml += '\n'
    xml += '\t<body>\n'

    if body_html:
        body_paragraphs = re.split(r'\n\s*\n', body_html.strip())
        in_refs = False
        for para in body_paragraphs:
            para = para.strip()
            if not para:
                continue
            para_lower = para.lower()
            if para_lower.startswith('introducción') or para_lower.startswith('introduccion'):
                xml += '\n\t\t<sec sec-type="intro">\n'
                xml += '\t\t\t<title>Introducción</title>\n'
                content = para[12:].strip() if len(para) > 12 else ""
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
                in_refs = False
            elif para_lower.startswith('metodología') or para_lower.startswith('metodologia') or para_lower.startswith('métodos') or para_lower.startswith('metodos') or para_lower.startswith('material'):
                xml += '\n\t\t<sec sec-type="methods">\n'
                xml += '\t\t\t<title>Metodología</title>\n'
                content = para[11:].strip()
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
            elif para_lower.startswith('resultados'):
                xml += '\n\t\t<sec sec-type="results">\n'
                xml += '\t\t\t<title>Resultados</title>\n'
                content = para[10:].strip()
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
            elif para_lower.startswith('discusión') or para_lower.startswith('discusion'):
                xml += '\n\t\t<sec sec-type="discussion">\n'
                xml += '\t\t\t<title>Discusión</title>\n'
                content = para[9:].strip()
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
            elif para_lower.startswith('conclusiones') or para_lower.startswith('conclusión') or para_lower.startswith('conclusion'):
                xml += '\n\t\t<sec sec-type="conclusions">\n'
                xml += '\t\t\t<title>Conclusiones</title>\n'
                content = para[12:].strip() if para_lower.startswith('conclusiones') else para[10:].strip()
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
            elif para_lower.startswith('justificación') or para_lower.startswith('justificacion'):
                xml += '\n\t\t<sec>\n'
                xml += '\t\t\t<title>Justificación</title>\n'
                content = para[13:].strip() if para_lower.startswith('justificación') else para[12:].strip()
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
            elif para_lower.startswith('objetivos'):
                xml += '\n\t\t<sec>\n'
                xml += '\t\t\t<title>Objetivos</title>\n'
                content = para[9:].strip()
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
            elif para_lower.startswith('referencias') or para_lower.startswith('bibliografía') or para_lower.startswith('bibliografia'):
                in_refs = True
            elif in_refs:
                pass
            else:
                xml += f'\n\t\t<sec>\n'
                first_word = para.split()[0] if para.split() else ""
                if para_lower.startswith('planteamiento'):
                    xml += '\t\t\t<title>Planteamiento del problema</title>\n'
                    content = para[23:].strip()
                else:
                    xml += f'\t\t\t<title>{escape_xml(first_word)}</title>\n'
                    content = para[len(first_word):].strip()
                if content:
                    xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'

        for sec_match in re.finditer(r'</sec>', xml):
            pass
    else:
        xml += '\t\t<sec>\n'
        xml += '\t\t\t<title>Contenido</title>\n'
        xml += '\t\t\t<p>Sin contenido disponible.</p>\n'
        xml += '\t\t</sec>\n'

    xml += '\n\t</body>\n'

    xml += '\n\t<back>\n'
    xml += '\t\t<ref-list>\n'
    xml += '\t\t\t<title>Referencias</title>\n'

    if refs_count > 0:
        for i in range(1, refs_count + 1):
            ref_text = data.get(f'ref_text_{i}', '')
            if ref_text:
                xml += '\n'
                xml += f'\t\t\t<ref id="B{i}">\n'
                xml += f'\t\t\t\t<mixed-citation>{escape_xml(ref_text)}</mixed-citation>\n'
                xml += '\t\t\t</ref>\n'
    else:
        xml += '\n'
        xml += '\t\t\t<ref id="B1">\n'
        xml += '\t\t\t\t<mixed-citation>Referencia 1</mixed-citation>\n'
        xml += '\t\t\t</ref>\n'

    xml += '\t\t</ref-list>\n'
    xml += '\t</back>\n'
    xml += '</article>'

    return xml

def escape_xml(text):
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    text = text.replace('"', '&quot;')
    text = text.replace("'", '&apos;')
    return text

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No se encontró el archivo'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No se seleccionó ningún archivo'}), 400
    ext = os.path.splitext(file.filename.lower())[1]
    if ext not in ('.pdf', '.docx'):
        return jsonify({'error': 'El archivo debe ser PDF o DOCX'}), 400

    file_id = str(uuid.uuid4())
    filename = f'{file_id}{ext}'
    ensure_upload_dir()
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        if ext == '.pdf':
            texto = extract_text_from_pdf(filepath)
        else:
            texto = extract_text_from_docx(filepath)
    except Exception as e:
        os.remove(filepath)
        return jsonify({'error': f'Error al leer el archivo: {str(e)}'}), 400

    if not texto.strip():
        os.remove(filepath)
        return jsonify({'error': 'No se pudo extraer texto del archivo. Verifica que no esté dañado o protegido.'}), 400

    title = guess_title(texto)
    abstract = guess_abstract(texto)
    keywords = guess_keywords(texto, 'es')
    refs = guess_references(texto)

    lines = texto.strip().split('\n')
    total_lines = len(lines)

    original_name = os.path.splitext(os.path.basename(file.filename))[0]

    return jsonify({
        'file_id': file_id,
        'texto': texto[:50000],
        'title_es': title,
        'abstract_es': abstract,
        'keywords_es': ', '.join(keywords),
        'references': refs,
        'original_name': original_name,
        'total_lines': total_lines
    })

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.form.to_dict()
        xml = generate_jats_xml(data)
        file_id = data.get('file_id', 'result')
        original_name = data.get('original_name', 'articulo')
        output_filename = f'{file_id}.xml'
        ensure_upload_dir()
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(xml)
        from urllib.parse import quote
        download_url = url_for('download', filename=output_filename) + '?name=' + quote(original_name + '.xml')
        return jsonify({'download_url': download_url})
    except Exception as e:
        return jsonify({'error': f'Error al generar XML: {str(e)}'}), 500

@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(filepath):
        return 'Archivo no encontrado', 404
    dl_name = request.args.get('name') or os.path.splitext(os.path.basename(filename))[0] + '.xml'
    return send_file(filepath, as_attachment=True, download_name=dl_name)

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'El archivo es demasiado grande. Máximo 50 MB.'}), 413

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': 'Error interno del servidor. Intenta de nuevo.'}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
