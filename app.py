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

# Palabras que indican que la línea es una cabecera de revista, no el título del artículo
HEADER_SKIP_PATTERNS = re.compile(
    r'(revista|journal|issn|vol\.?\s*\d|núm|number|pp\.?\s*\d|doi\s*:|http|universidad de sonora'
    r'|investigaci[oó]n acad[eé]mica|sin frontera|editorial|editor|published|publicado en)',
    re.IGNORECASE
)

def guess_title(text):
    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
    candidates = []
    for line in lines[:40]:
        if len(line) < 20 or len(line) > 350:
            continue
        if HEADER_SKIP_PATTERNS.search(line):
            continue
        # Saltar líneas que parecen datos de autoría o institución (contienen @ o son muy cortas)
        if '@' in line:
            continue
        # Evitar líneas que son solo números/fechas
        if re.match(r'^[\d\s\/\-\.,:]+$', line):
            continue
        candidates.append(line)
    # El título suele ser la línea más larga de los primeros candidatos o la primera válida
    if candidates:
        # Preferir candidatos con palabras mayúsculas o que parecen oraciones completas
        for c in candidates[:5]:
            # Si tiene más de 4 palabras y no es solo mayúsculas (que sería un encabezado de sección)
            words = c.split()
            if len(words) >= 4:
                return c
        return candidates[0]
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
    """
    Intenta detectar autores en las primeras líneas del documento.
    Devuelve lista de dicts con {'surname': ..., 'given': ...}.
    Si no detecta nada, devuelve un autor por defecto.
    """
    # Patron: línea que contiene solo nombres propios (2-5 palabras capitalizadas)
    # típicamente entre la línea 3 y 25 del documento
    name_pattern = re.compile(
        r'^([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,4})$'
    )
    # Palabras que indican que no es un autor
    skip_words = re.compile(
        r'(revista|journal|resumen|abstract|introduc|metodolog|resultado|conclus|referenc'
        r'|universidad|instituto|departamento|palabras\s+clave|keywords|issn|vol|http|doi)',
        re.IGNORECASE
    )
    authors = []
    lines = text.strip().split('\n')
    for line in lines[2:40]:   # Saltar la primera línea (suele ser el título)
        line_clean = line.strip()
        if not line_clean or len(line_clean) < 6 or len(line_clean) > 120:
            continue
        if skip_words.search(line_clean):
            continue
        if '@' in line_clean:   # Email → probablemente línea de autor
            continue
        m = name_pattern.match(line_clean)
        if m:
            full = m.group(1).strip()
            parts = full.rsplit(' ', 1)   # Separar apellido (último) y nombre
            if len(parts) == 2:
                authors.append({'given': parts[0], 'surname': parts[1]})
            else:
                authors.append({'given': '', 'surname': full})
            if len(authors) >= 5:
                break

    # Fallback: buscar con regex en las primeras 2000 chars
    if not authors:
        possible = re.findall(
            r'(?:^|\n)[ \t]*([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,3})[ \t]*(?:\n|$)',
            text[:2500], re.MULTILINE
        )
        seen = set()
        for p in possible:
            p = p.strip()
            if p in seen or skip_words.search(p):
                continue
            seen.add(p)
            parts = p.rsplit(' ', 1)
            if len(parts) == 2:
                authors.append({'given': parts[0], 'surname': parts[1]})
            else:
                authors.append({'given': '', 'surname': p})
            if len(authors) >= 3:
                break

    return authors  # Puede ser lista vacía; el front-end pondrá el default

def parse_reference(ref_text):
    year_match = re.search(r'\b(19\d\d|20\d\d)\b', ref_text)
    year = year_match.group(1) if year_match else ""
    authors = ""
    title = ""
    source = ""
    if year_match:
        parts = ref_text.split(year_match.group(0), 1)
        if len(parts) == 2:
            authors_part = parts[0].strip(' (.,;')
            rest = parts[1].strip(' ).,;')
            authors = authors_part
            rest_parts = re.split(r'[\.\?]\s+', rest, 1)
            if len(rest_parts) == 2:
                title = rest_parts[0].strip()
                source_part = rest_parts[1].strip()
                source_match = re.match(r'^([^,]+)', source_part)
                source = source_match.group(1).strip() if source_match else source_part
            else:
                title = rest
    return {
        "year": year,
        "authors": authors,
        "title": title,
        "source": source
    }

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
    # NOTA: No se incluye <!DOCTYPE> para evitar que OJS3 intente resolver el DTD
    # externo en línea, lo que causa que el galley se muestre en blanco.
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
            if not name and not surname and not given:
                continue
            xml += '\t\t\t\t<contrib contrib-type="author">\n'
            if orcid:
                xml += f'\t\t\t\t\t<contrib-id contrib-id-type="orcid">{escape_xml(orcid)}</contrib-id>\n'
            xml += '\t\t\t\t\t<name>\n'
            if surname and given:
                xml += f'\t\t\t\t\t\t<surname>{escape_xml(surname)}</surname>\n'
                xml += f'\t\t\t\t\t\t<given-names>{escape_xml(given)}</given-names>\n'
            elif surname:
                xml += f'\t\t\t\t\t\t<surname>{escape_xml(surname)}</surname>\n'
            elif given:
                xml += f'\t\t\t\t\t\t<given-names>{escape_xml(given)}</given-names>\n'
            elif name:
                parts = name.strip().split(' ', 1)
                if len(parts) > 1:
                    xml += f'\t\t\t\t\t\t<surname>{escape_xml(parts[0])}</surname>\n'
                    xml += f'\t\t\t\t\t\t<given-names>{escape_xml(parts[1])}</given-names>\n'
                else:
                    xml += f'\t\t\t\t\t\t<given-names>{escape_xml(name)}</given-names>\n'
            xml += '\t\t\t\t\t</name>\n'
            xml += f'\t\t\t\t\t<xref ref-type="aff" rid="aff{aff_id}">{aff_id}</xref>\n'
            if email:
                xml += f'\t\t\t\t\t<email>{escape_xml(email)}</email>\n'
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
        paragraphs = [p.strip() for p in body_html.split('\n') if p.strip()]
        sec_open = False
        for para in paragraphs:
            para_lower = para.lower()
            is_heading = False
            sec_type = None
            title = ""

            if para_lower.startswith('introducción') or para_lower.startswith('introduccion'):
                is_heading = True
                sec_type = 'intro'
                title = 'Introducción'
            elif para_lower.startswith('metodología') or para_lower.startswith('metodologia') or para_lower.startswith('métodos') or para_lower.startswith('metodos') or para_lower.startswith('material'):
                is_heading = True
                sec_type = 'methods'
                title = 'Metodología'
            elif para_lower.startswith('resultados'):
                is_heading = True
                sec_type = 'results'
                title = 'Resultados'
            elif para_lower.startswith('discusión') or para_lower.startswith('discusion'):
                is_heading = True
                sec_type = 'discussion'
                title = 'Discusión'
            elif para_lower.startswith('conclusiones') or para_lower.startswith('conclusión') or para_lower.startswith('conclusion'):
                is_heading = True
                sec_type = 'conclusions'
                title = 'Conclusiones'
            elif para_lower.startswith('justificación') or para_lower.startswith('justificacion'):
                is_heading = True
                title = 'Justificación'
            elif para_lower.startswith('objetivos'):
                is_heading = True
                title = 'Objetivos'
            elif para_lower.startswith('planteamiento'):
                is_heading = True
                title = 'Planteamiento del problema'
            elif para_lower.startswith('referencias') or para_lower.startswith('bibliografía') or para_lower.startswith('bibliografia'):
                continue

            if is_heading:
                if sec_open:
                    xml += '\t\t</sec>\n'
                if sec_type:
                    xml += f'\n\t\t<sec sec-type="{sec_type}">\n'
                else:
                    xml += '\n\t\t<sec>\n'
                xml += f'\t\t\t<title>{escape_xml(title)}</title>\n'
                sec_open = True

                if len(para) > len(title) + 5:
                    content = re.sub(rf'^{title}[:\s\-\.]*', '', para, flags=re.IGNORECASE).strip()
                    if content:
                        xml += f'\t\t\t<p>{escape_xml(content)}</p>\n'
            else:
                if not sec_open:
                    xml += '\n\t\t<sec>\n'
                    xml += '\t\t\t<title>Introducción</title>\n'
                    sec_open = True
                xml += f'\t\t\t<p>{escape_xml(para)}</p>\n'

        if sec_open:
            xml += '\t\t</sec>\n'
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
                ref_info = parse_reference(ref_text)
                xml += '\n'
                xml += f'\t\t\t<ref id="B{i}">\n'
                xml += f'\t\t\t\t<label>{i}</label>\n'
                xml += f'\t\t\t\t<mixed-citation>{escape_xml(ref_text)}</mixed-citation>\n'
                xml += '\t\t\t\t<element-citation publication-type="journal">\n'
                if ref_info['authors']:
                    xml += '\t\t\t\t\t<person-group person-group-type="author">\n'
                    xml += f'\t\t\t\t\t\t<name>\n\t\t\t\t\t\t\t<surname>{escape_xml(ref_info["authors"])}</surname>\n\t\t\t\t\t\t</name>\n'
                    xml += '\t\t\t\t\t</person-group>\n'
                if ref_info['title']:
                    xml += f'\t\t\t\t\t<article-title>{escape_xml(ref_info["title"])}</article-title>\n'
                if ref_info['source']:
                    xml += f'\t\t\t\t\t<source>{escape_xml(ref_info["source"])}</source>\n'
                if ref_info['year']:
                    xml += f'\t\t\t\t\t<year>{escape_xml(ref_info["year"])}</year>\n'
                xml += '\t\t\t\t</element-citation>\n'
                xml += '\t\t\t</ref>\n'
    else:
        xml += '\n'
        xml += '\t\t\t<ref id="B1">\n'
        xml += '\t\t\t\t<label>1</label>\n'
        xml += '\t\t\t\t<mixed-citation>Referencia 1</mixed-citation>\n'
        xml += '\t\t\t\t<element-citation publication-type="journal">\n'
        xml += '\t\t\t\t\t<article-title>Referencia 1</article-title>\n'
        xml += '\t\t\t\t</element-citation>\n'
        xml += '\t\t\t</ref>\n'

    xml += '\t\t</ref-list>\n'
    xml += '\t</back>\n'
    xml += '</article>'

    return xml

def escape_xml(text):
    if not isinstance(text, str):
        text = str(text or '')
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

    base = os.path.splitext(os.path.basename(file.filename))[0]
    m = re.match(r'^(\d+)', base)
    original_name = m.group(1) if m else base

    authors = guess_authors(texto)

    return jsonify({
        'file_id': file_id,
        'texto': texto[:50000],
        'title_es': title,
        'abstract_es': abstract,
        'keywords_es': ', '.join(keywords),
        'references': refs,
        'authors': authors,
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
