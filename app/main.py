from __future__ import annotations

import asyncio
import io
import math
import os
import re
import unicodedata
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import httpx
import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse, StreamingResponse
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

APP_NAME = "academic-search-api"
APP_VERSION = "0.9.1"

OPENALEX_BASE_URL = "https://api.openalex.org/works"
OPENALEX_EMAIL = os.getenv("OPENALEX_EMAIL", "").strip()

SEMANTIC_SCHOLAR_BASE_URL = "https://api.semanticscholar.org/graph/v1"
SEMANTIC_SCHOLAR_API_KEY = os.getenv("SEMANTIC_SCHOLAR_API_KEY", "").strip()

SCOPUS_SEARCH_URL = "https://api.elsevier.com/content/search/scopus"
SCOPUS_API_KEY = os.getenv("SCOPUS_API_KEY", "").strip()
SCOPUS_PAGE_SIZE = 25

CROSSREF_BASE_URL = "https://api.crossref.org/works"
CROSSREF_MAILTO = os.getenv("CROSSREF_MAILTO", "").strip()

RECENT_WINDOW_YEARS = 5

MODOS_VALIDOS = {"mixto", "recientes", "clasicos"}
FUENTES_VALIDAS = {"openalex", "semanticscholar", "scopus", "crossref", "mixto"}

SEMANTIC_SCHOLAR_FIELDS = ",".join(
    [
        "paperId",
        "url",
        "title",
        "abstract",
        "year",
        "citationCount",
        "authors",
        "venue",
        "publicationDate",
        "publicationTypes",
        "externalIds",
        "openAccessPdf",
    ]
)

DOCUMENT_TYPE_MAP = {
    "article": "Artículo",
    "dissertation": "Tesis",
    "preprint": "Preprint",
    "book": "Libro",
    "book-chapter": "Capítulo de libro",
    "report": "Informe",
    "dataset": "Conjunto de datos",
    "editorial": "Editorial",
    "letter": "Carta",
    "review": "Revisión",
    "paratext": "Paratexto",
    "reference-entry": "Entrada de referencia",
    "journal-article": "Artículo",
    "proceedings-article": "Conferencia",
    "book-section": "Capítulo de libro",
    "monograph": "Libro",
    "edited-book": "Libro",
    "reference-book": "Libro",
    "book-set": "Libro",
    "posted-content": "Preprint",
}

SEMANTIC_PUBLICATION_TYPE_MAP = {
    "journalarticle": "Artículo",
    "review": "Revisión",
    "conference": "Conferencia",
    "conferencepaper": "Conferencia",
    "book": "Libro",
    "booksection": "Capítulo de libro",
    "editorial": "Editorial",
    "letter": "Carta",
    "case report": "Informe",
    "clinical trial": "Artículo",
    "meta-analysis": "Revisión",
}

SOURCE_TYPE_MAP = {
    "journal": "Revista",
    "repository": "Repositorio",
    "conference": "Conferencia",
    "ebook-platform": "Plataforma de eBook",
    "book-series": "Serie de libros",
    "metadata": "Metadatos",
    "other": "Otra fuente",
    "unknown": "Desconocido",
}

SPECIAL_QUERY_KEYWORDS = {
    "telemedicina": [
        "telemedicina",
        "telemedicine",
        "telehealth",
        "virtual care",
        "remote care",
    ],
    "telemedicine": [
        "telemedicina",
        "telemedicine",
        "telehealth",
        "virtual care",
        "remote care",
    ],
}

PROVIDER_PRIORITY = {
    "openalex": 2.0,
    "scopus": 1.8,
    "semanticscholar": 1.5,
    "crossref": 1.4,
}

LANGUAGE_ALIASES = {
    "es": "es",
    "espanol": "es",
    "español": "es",
    "spanish": "es",
    "en": "en",
    "ingles": "en",
    "inglés": "en",
    "english": "en",
    "pt": "pt",
    "portugues": "pt",
    "portugués": "pt",
    "portuguese": "pt",
    "fr": "fr",
    "frances": "fr",
    "francés": "fr",
    "french": "fr",
    "de": "de",
    "aleman": "de",
    "alemán": "de",
    "german": "de",
    "it": "it",
    "italiano": "it",
    "italian": "it",
}

DOCUMENT_FILTER_ALIASES = {
    "articulo": [
        "articulo",
        "artículos",
        "articulos",
        "article",
        "articles",
        "journalarticle",
        "journal article",
        "journal-article",
    ],
    "revision": [
        "revision",
        "revisión",
        "revisiones",
        "review",
        "reviews",
        "systematic review",
        "meta-analysis",
        "metaanalysis",
        "scoping review",
    ],
    "tesis": [
        "tesis",
        "thesis",
        "dissertation",
        "dissertations",
    ],
    "libro": [
        "libro",
        "libros",
        "book",
        "books",
        "monograph",
        "edited-book",
        "reference-book",
        "book-set",
    ],
    "capitulo_libro": [
        "capitulo_libro",
        "capitulo de libro",
        "capítulos de libro",
        "capitulos de libro",
        "book chapter",
        "book chapters",
        "book-chapter",
        "booksection",
        "book section",
        "book-section",
    ],
    "preprint": [
        "preprint",
        "preprints",
        "posted-content",
    ],
}

app = FastAPI(
    title=APP_NAME,
    version=APP_VERSION,
    description="API académica para búsqueda bibliográfica con OpenAlex, Semantic Scholar, Scopus, Crossref y exportación a Excel.",
)


# =========================================================
# Utilidades generales
# =========================================================

def normalize_text(text: Optional[str]) -> str:
    if not text:
        return ""
    text = text.strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_title(text: Optional[str]) -> str:
    text = normalize_text(text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_doi(doi: Optional[str]) -> str:
    if not doi:
        return ""
    doi = doi.strip()
    doi = doi.replace("https://doi.org/", "").replace("http://doi.org/", "")
    doi = doi.replace("doi.org/", "")
    return doi.strip().lower()


def safe_filename(text: str) -> str:
    text = normalize_text(text)
    text = re.sub(r"[^a-z0-9\s_-]", "", text)
    text = re.sub(r"\s+", "_", text).strip("_")
    return text or "resultados"


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except Exception:
        return default


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def interleave_multiple_lists(*lists: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not lists:
        return []

    out: List[Dict[str, Any]] = []
    max_len = max((len(lst) for lst in lists), default=0)

    for i in range(max_len):
        for lst in lists:
            if i < len(lst):
                out.append(lst[i])

    return out


def deduplicate_results(results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen_titles = set()
    seen_dois = set()
    deduped = []

    for item in results:
        normalized_title = normalize_title(item.get("title"))
        normalized_doi = normalize_doi(item.get("doi"))

        title_dup = normalized_title and normalized_title in seen_titles
        doi_dup = normalized_doi and normalized_doi in seen_dois

        if title_dup or doi_dup:
            continue

        if normalized_title:
            seen_titles.add(normalized_title)
        if normalized_doi:
            seen_dois.add(normalized_doi)

        deduped.append(item)

    return deduped


def split_targets(max_results: int, modo: str) -> Tuple[int, int]:
    if modo == "mixto":
        target_recent = math.ceil(max_results / 2)
        target_classic = max_results - target_recent
    elif modo == "recientes":
        target_recent = max_results
        target_classic = 0
    else:
        target_recent = 0
        target_classic = max_results
    return target_recent, target_classic


def classify_source_type(source_type: Optional[str]) -> str:
    if not source_type:
        return SOURCE_TYPE_MAP["unknown"]
    return SOURCE_TYPE_MAP.get(source_type, source_type.capitalize())


def build_title_keywords(query: str) -> List[str]:
    q_norm = normalize_text(query)
    if q_norm in SPECIAL_QUERY_KEYWORDS:
        return SPECIAL_QUERY_KEYWORDS[q_norm]

    keywords = [q_norm]
    tokens = [tok for tok in re.split(r"\W+", q_norm) if len(tok) >= 4]
    keywords.extend(tokens)

    return list(dict.fromkeys([k for k in keywords if k]))


def get_primary_query_tokens(query: str) -> List[str]:
    q_norm = normalize_text(query)
    return [tok for tok in re.split(r"\W+", q_norm) if len(tok) >= 4]


def title_matches_query_keywords(title: str, query: str) -> bool:
    normalized_title = normalize_text(title)
    keywords = build_title_keywords(query)
    return any(keyword in normalized_title for keyword in keywords)


def apply_strict_title_relevance(
    results: List[Dict[str, Any]],
    query: str,
    strict_relevance: bool,
) -> List[Dict[str, Any]]:
    if not strict_relevance:
        return results
    return [item for item in results if title_matches_query_keywords(item.get("title", ""), query)]


def get_doi_url(doi: str) -> str:
    return f"https://doi.org/{doi}" if doi else ""


# =========================================================
# Política documental por defecto
# =========================================================

def document_type_policy_passes(
    item: Dict[str, Any],
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> bool:
    raw_type = normalize_text(item.get("type") or "")
    tipo_documento = normalize_text(item.get("tipo_documento") or "")
    publication_types = normalize_text(item.get("publication_types") or "")

    bag = " ".join([raw_type, tipo_documento, publication_types])

    article_markers = [
        "article",
        "articulo",
        "journalarticle",
        "journal-article",
        "clinical trial",
    ]
    thesis_markers = [
        "tesis",
        "thesis",
        "dissertation",
    ]
    review_markers = [
        "review",
        "revision",
        "revisión",
        "meta-analysis",
        "systematic review",
        "scoping review",
    ]
    book_markers = [
        "book",
        "libro",
        "book-chapter",
        "book chapter",
        "booksection",
        "book-section",
        "chapter",
        "capitulo de libro",
        "capítulo de libro",
        "monograph",
        "edited-book",
        "reference-book",
        "book-set",
    ]
    preprint_markers = [
        "preprint",
        "posted-content",
    ]

    if any(marker in bag for marker in article_markers):
        return True

    if any(marker in bag for marker in thesis_markers):
        return True

    if any(marker in bag for marker in review_markers):
        return True

    if incluir_libros and any(marker in bag for marker in book_markers):
        return True

    if incluir_preprints and any(marker in bag for marker in preprint_markers):
        return True

    return False


# =========================================================
# Filtros avanzados opcionales
# =========================================================

def normalize_language_code(value: Optional[str]) -> str:
    if not value:
        return ""
    text = normalize_text(value)
    if text in LANGUAGE_ALIASES:
        return LANGUAGE_ALIASES[text]
    return text[:2] if len(text) >= 2 else text


def parse_tipo_documento_filter(tipo_documento: Optional[str]) -> List[str]:
    if not tipo_documento:
        return []

    parsed: List[str] = []

    for raw in str(tipo_documento).split(","):
        token = normalize_text(raw)
        if not token:
            continue

        matched = None
        for canonical, aliases in DOCUMENT_FILTER_ALIASES.items():
            normalized_aliases = [normalize_text(a) for a in aliases]
            if token == canonical or token in normalized_aliases:
                matched = canonical
                break

        parsed.append(matched or token)

    return list(dict.fromkeys(parsed))


def normalize_tipo_documento_canonical(tipo_documento: Optional[str]) -> Optional[str]:
    parsed = parse_tipo_documento_filter(tipo_documento)
    if not parsed:
        return None
    return parsed[0]


def get_requested_openalex_types(
    tipo_documento: Optional[str],
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> List[str]:
    td = normalize_tipo_documento_canonical(tipo_documento)

    if td == "articulo":
        return ["article"]
    if td == "revision":
        return ["review"]
    if td == "tesis":
        return ["dissertation"]
    if td == "libro":
        return ["book", "book-chapter"]
    if td == "capitulo_libro":
        return ["book-chapter"]
    if td == "preprint":
        return ["preprint"]

    allowed = ["article", "dissertation", "review"]

    if incluir_libros:
        allowed.extend(["book", "book-chapter"])

    if incluir_preprints:
        allowed.append("preprint")

    return allowed


def tipo_documento_is_book_like(
    tipo_documento: Optional[str],
    incluir_libros: bool = False,
) -> bool:
    td = normalize_tipo_documento_canonical(tipo_documento)
    return td in {"libro", "capitulo_libro"} or incluir_libros


def get_result_document_bag(item: Dict[str, Any]) -> str:
    parts: List[str] = []

    for key in ["type", "tipo_documento", "source_type"]:
        value = item.get(key)
        if value:
            parts.append(str(value))

    publication_types = item.get("publication_types")
    if isinstance(publication_types, list):
        parts.extend([str(x) for x in publication_types if x])
    elif publication_types:
        parts.append(str(publication_types))

    return " | ".join(normalize_text(p) for p in parts if p)


def item_matches_tipo_documento(item: Dict[str, Any], requested_types: List[str]) -> bool:
    if not requested_types:
        return True

    bag = get_result_document_bag(item)
    if not bag:
        return False

    for requested in requested_types:
        aliases = DOCUMENT_FILTER_ALIASES.get(requested, [requested])
        normalized_aliases = [normalize_text(a) for a in aliases]
        if any(alias in bag for alias in normalized_aliases):
            return True

    return False


def extract_year_from_item(item: Dict[str, Any]) -> int:
    year = safe_int(item.get("year"), 0)
    if year > 0:
        return year

    publication_date = str(item.get("publication_date") or "")
    match = re.search(r"(19|20)\d{2}", publication_date)
    if match:
        return safe_int(match.group(0), 0)

    return 0


def item_matches_year_range(
    item: Dict[str, Any],
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
) -> bool:
    if year_from is None and year_to is None:
        return True

    year = extract_year_from_item(item)
    if year <= 0:
        return False

    if year_from is not None and year < year_from:
        return False

    if year_to is not None and year > year_to:
        return False

    return True


def item_matches_language(item: Dict[str, Any], idioma: Optional[str] = None) -> bool:
    if not idioma:
        return True

    wanted = normalize_language_code(idioma)
    result_lang = normalize_language_code(item.get("language"))

    if not result_lang:
        return False

    return result_lang == wanted


def apply_advanced_filters(
    results: List[Dict[str, Any]],
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    idioma: Optional[str] = None,
    tipo_documento: Optional[str] = None,
) -> List[Dict[str, Any]]:
    requested_types = parse_tipo_documento_filter(tipo_documento)

    filtered: List[Dict[str, Any]] = []
    for item in results:
        if not item_matches_year_range(item, year_from=year_from, year_to=year_to):
            continue
        if not item_matches_language(item, idioma=idioma):
            continue
        if not item_matches_tipo_documento(item, requested_types=requested_types):
            continue
        filtered.append(item)

    return filtered


def tipo_documento_requires_books(tipo_documento: Optional[str]) -> bool:
    requested_types = parse_tipo_documento_filter(tipo_documento)
    return any(t in {"libro", "capitulo_libro"} for t in requested_types)


def tipo_documento_requires_preprints(tipo_documento: Optional[str]) -> bool:
    requested_types = parse_tipo_documento_filter(tipo_documento)
    return "preprint" in requested_types


# =========================================================
# Ranking de relevancia
# =========================================================

def count_title_keyword_matches(title: str, query: str) -> int:
    normalized_title = normalize_text(title)
    keywords = build_title_keywords(query)
    return sum(1 for keyword in keywords if keyword and keyword in normalized_title)


def has_exact_query_match(title: str, query: str) -> bool:
    normalized_title = normalize_text(title)
    normalized_query = normalize_text(query)
    return bool(normalized_query and normalized_query in normalized_title)


def has_all_primary_tokens(title: str, query: str) -> bool:
    normalized_title = normalize_text(title)
    tokens = get_primary_query_tokens(query)
    if not tokens:
        return False
    return all(tok in normalized_title for tok in tokens)


def get_document_type_bonus(item: Dict[str, Any]) -> float:
    doc = normalize_text(item.get("tipo_documento") or "")
    if "revision" in doc or "review" in doc or "revisión" in doc:
        return 6.0
    if "articulo" in doc or "article" in doc:
        return 5.0
    if "tesis" in doc or "dissertation" in doc or "thesis" in doc:
        return 4.0
    if "libro" in doc or "book" in doc:
        return 3.0
    if "preprint" in doc or "posted-content" in doc:
        return 2.0
    return 1.0


def get_source_type_bonus(item: Dict[str, Any]) -> float:
    source_type = normalize_text(item.get("source_type") or "")
    if source_type == "journal":
        return 3.0
    if source_type == "repository":
        return 1.5
    if source_type == "conference":
        return 1.0
    if source_type in {"ebook-platform", "book-series"}:
        return 1.0
    return 0.5


def get_provider_bonus(item: Dict[str, Any]) -> float:
    provider = normalize_text(item.get("provider") or "")
    return float(PROVIDER_PRIORITY.get(provider, 0.0))


def get_year_bonus(item: Dict[str, Any]) -> float:
    year = safe_int(item.get("year"), 0)
    if year <= 0:
        return 0.0
    current_year = datetime.utcnow().year
    age = max(0, current_year - year)
    return max(0.0, 6.0 - min(age, 6))


def get_citations_bonus(item: Dict[str, Any]) -> float:
    citations = max(0, safe_int(item.get("citations"), 0))
    return min(15.0, math.log1p(citations) * 3.0)


def compute_relevance_score(item: Dict[str, Any], query: str) -> float:
    title = item.get("title", "")
    keyword_matches = count_title_keyword_matches(title, query)
    exact_match = has_exact_query_match(title, query)
    all_tokens = has_all_primary_tokens(title, query)

    score = 0.0
    score += min(30.0, keyword_matches * 6.0)
    score += 30.0 if exact_match else 0.0
    score += 15.0 if all_tokens else 0.0
    score += get_document_type_bonus(item)
    score += get_source_type_bonus(item)
    score += get_provider_bonus(item)
    score += get_year_bonus(item)
    score += get_citations_bonus(item)
    score += 1.0 if bool(item.get("open_access")) else 0.0

    return round(score, 2)


def annotate_and_rank_results(results: List[Dict[str, Any]], query: str) -> List[Dict[str, Any]]:
    annotated = []
    for item in results:
        ranked_item = dict(item)
        ranked_item["score_relevancia"] = compute_relevance_score(ranked_item, query)
        ranked_item["match_keywords_title"] = count_title_keyword_matches(
            ranked_item.get("title", ""), query
        )
        annotated.append(ranked_item)

    annotated.sort(
        key=lambda x: (
            safe_float(x.get("score_relevancia"), 0.0),
            safe_int(x.get("citations"), 0),
            safe_int(x.get("year"), 0),
        ),
        reverse=True,
    )
    return annotated


def balance_ranked_results_by_provider(results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not results:
        return []

    buckets: Dict[str, List[Dict[str, Any]]] = {}
    for item in results:
        provider = normalize_text(item.get("provider") or "unknown")
        buckets.setdefault(provider, []).append(item)

    def top_score(provider_name: str) -> float:
        bucket = buckets.get(provider_name, [])
        if not bucket:
            return -1.0
        return safe_float(bucket[0].get("score_relevancia"), 0.0)

    ordered_providers = sorted(buckets.keys(), key=top_score, reverse=True)
    balanced: List[Dict[str, Any]] = []

    while True:
        active_providers = [p for p in ordered_providers if buckets.get(p)]
        if not active_providers:
            break

        active_providers = sorted(active_providers, key=top_score, reverse=True)

        moved = False
        for provider in active_providers:
            if buckets[provider]:
                balanced.append(buckets[provider].pop(0))
                moved = True

        if not moved:
            break

    return balanced


# =========================================================
# Scopus
# =========================================================

def build_scopus_test_query(query: str) -> str:
    q = (query or "").strip()
    if not q:
        return q

    q_upper = q.upper()
    advanced_markers = [
        "TITLE-ABS-KEY(",
        "TITLE(",
        "ABS(",
        "AUTHKEY(",
        "DOCTYPE(",
        "SRCTYPE(",
        "PUBYEAR",
    ]

    if any(marker in q_upper for marker in advanced_markers):
        return q

    return f"TITLE-ABS-KEY({q})"


def parse_scopus_entry(entry: Dict[str, Any], segmento_temporal: str = "") -> Dict[str, Any]:
    doi = normalize_doi(entry.get("prism:doi") or "")
    cover_date = entry.get("prism:coverDate") or ""
    aggregation_type = entry.get("prism:aggregationType") or ""
    source_type = (
        "journal"
        if normalize_text(aggregation_type) == "journal"
        else normalize_text(aggregation_type) or "unknown"
    )

    return {
        "title": entry.get("dc:title") or "",
        "authors": entry.get("dc:creator") or "",
        "year": safe_int(cover_date[:4], 0) if cover_date else None,
        "type": entry.get("subtype") or aggregation_type,
        "tipo_documento": entry.get("subtypeDescription") or aggregation_type or "Desconocido",
        "source": "Scopus",
        "provider": "scopus",
        "source_name": entry.get("prism:publicationName") or "Sin fuente identificada",
        "source_type": source_type,
        "tipo_fuente": classify_source_type(source_type),
        "revista_o_fuente": entry.get("prism:publicationName") or "Sin fuente identificada",
        "doi": doi,
        "doi_url": get_doi_url(doi),
        "citations": safe_int(entry.get("citedby-count"), 0),
        "url": entry.get("prism:url") or "",
        "landing_page_url": entry.get("prism:url") or "",
        "pdf_url": "",
        "publication_date": cover_date,
        "language": "",
        "open_access": str(entry.get("openaccess", "0")) == "1",
        "openalex_id": "",
        "paper_id": entry.get("eid") or "",
        "publication_types": entry.get("subtypeDescription") or "",
        "abstract": "",
        "segmento_temporal": segmento_temporal,
        "eid": entry.get("eid") or "",
    }


def scopus_item_passes_filters(
    item: Dict[str, Any],
    solo_revistas: bool = False,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> bool:
    source_type = normalize_text(item.get("source_type") or "")
    if solo_revistas and source_type != "journal":
        return False

    return document_type_policy_passes(
        item,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
    )


def build_scopus_search_query(
    query: str,
    recent: bool,
    cutoff_year: int,
    solo_revistas: bool = False,
) -> str:
    base_query = build_scopus_test_query(query)
    clauses = [base_query]

    if recent:
        clauses.append(f"PUBYEAR > {cutoff_year - 1}")
    else:
        clauses.append(f"PUBYEAR < {cutoff_year}")

    if solo_revistas:
        clauses.append("SRCTYPE(j)")

    return " AND ".join(clauses)


async def fetch_scopus_segment(
    client: httpx.AsyncClient,
    query: str,
    recent: bool,
    target_results: int,
    solo_revistas: bool = False,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> Tuple[int, List[Dict[str, Any]]]:
    current_year = datetime.utcnow().year
    cutoff_year = current_year - RECENT_WINDOW_YEARS

    scopus_query = build_scopus_search_query(
        query=query,
        recent=recent,
        cutoff_year=cutoff_year,
        solo_revistas=solo_revistas,
    )

    desired_pool = min(max(target_results * 4, SCOPUS_PAGE_SIZE), 100)
    total_found = 0
    collected: List[Dict[str, Any]] = []
    start = 0
    max_loops = 5
    segmento = "reciente" if recent else "clasico"

    for _ in range(max_loops):
        params = {
            "query": scopus_query,
            "count": SCOPUS_PAGE_SIZE,
            "start": start,
            "view": "STANDARD",
        }

        response = await client.get(
            SCOPUS_SEARCH_URL,
            params=params,
        )
        response.raise_for_status()

        payload = response.json()
        search_results = payload.get("search-results", {}) if isinstance(payload, dict) else {}

        if start == 0:
            total_found = safe_int(search_results.get("opensearch:totalResults"), 0)

        entries = search_results.get("entry", []) or []
        if not entries:
            break

        parsed = [parse_scopus_entry(entry, segmento_temporal=segmento) for entry in entries]
        parsed = [
            item
            for item in parsed
            if scopus_item_passes_filters(
                item,
                solo_revistas=solo_revistas,
                incluir_libros=incluir_libros,
                incluir_preprints=incluir_preprints,
            )
        ]

        collected.extend(parsed)
        collected = deduplicate_results(collected)

        if len(collected) >= desired_pool:
            break

        if len(entries) < SCOPUS_PAGE_SIZE:
            break

        start += SCOPUS_PAGE_SIZE
        if start >= total_found:
            break

    return total_found, collected[:desired_pool]


async def search_scopus_provider(
    query: str,
    target_recent: int,
    target_classic: int,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    strict_relevance: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> Dict[str, Any]:
    if not SCOPUS_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="SCOPUS_API_KEY no está configurada en variables de entorno."
        )

    total_recent = 0
    total_classic = 0
    recent_results: List[Dict[str, Any]] = []
    classic_results: List[Dict[str, Any]] = []

    headers = {
        "Accept": "application/json",
        "X-ELS-APIKey": SCOPUS_API_KEY,
    }

    async with httpx.AsyncClient(timeout=45.0, headers=headers) as client:
        try:
            tasks = []
            if target_recent > 0:
                tasks.append(
                    fetch_scopus_segment(
                        client,
                        query=query,
                        recent=True,
                        target_results=target_recent,
                        solo_revistas=solo_revistas,
                        incluir_libros=incluir_libros,
                        incluir_preprints=incluir_preprints,
                    )
                )

            if target_classic > 0:
                tasks.append(
                    fetch_scopus_segment(
                        client,
                        query=query,
                        recent=False,
                        target_results=target_classic,
                        solo_revistas=solo_revistas,
                        incluir_libros=incluir_libros,
                        incluir_preprints=incluir_preprints,
                    )
                )

            task_results = await asyncio.gather(*tasks)

        except httpx.HTTPError as e:
            raise HTTPException(
                status_code=502,
                detail=f"Error al consultar Scopus: {str(e)}"
            ) from e

    idx = 0
    if target_recent > 0:
        total_recent, recent_results = task_results[idx]
        idx += 1
        recent_results = apply_strict_title_relevance(
            recent_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        recent_results = deduplicate_results(recent_results)

    if target_classic > 0:
        total_classic, classic_results = task_results[idx]
        classic_results = apply_strict_title_relevance(
            classic_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        classic_results = deduplicate_results(classic_results)

    return {
        "provider": "scopus",
        "source_label": "Scopus",
        "total_recent": total_recent,
        "total_classic": total_classic,
        "total_found": total_recent + total_classic,
        "total_is_estimated": False,
        "recent_results": recent_results,
        "classic_results": classic_results,
    }


# =========================================================
# OpenAlex
# =========================================================

def classify_openalex_document_type(doc_type: Optional[str]) -> str:
    if not doc_type:
        return "Desconocido"
    return DOCUMENT_TYPE_MAP.get(doc_type, doc_type.capitalize())


def extract_openalex_authors(authorships: Optional[List[Dict[str, Any]]]) -> str:
    if not authorships:
        return ""
    names = []
    for authorship in authorships:
        author = authorship.get("author") or {}
        display_name = author.get("display_name")
        if display_name:
            names.append(display_name)
    return "; ".join(names)


def get_openalex_primary_source(work: Dict[str, Any]) -> Dict[str, Any]:
    primary_location = work.get("primary_location") or {}
    source = primary_location.get("source") or {}
    return source


def get_openalex_source_name(work: Dict[str, Any]) -> str:
    source = get_openalex_primary_source(work)
    return source.get("display_name") or "Sin fuente identificada"


def get_openalex_source_type(work: Dict[str, Any]) -> str:
    source = get_openalex_primary_source(work)
    return source.get("type") or "unknown"


def get_openalex_doi(work: Dict[str, Any]) -> str:
    ids = work.get("ids") or {}
    doi = ids.get("doi") or work.get("doi")
    if doi:
        return normalize_doi(doi)
    return ""


def get_openalex_landing_page_url(work: Dict[str, Any]) -> str:
    primary_location = work.get("primary_location") or {}
    best_oa_location = work.get("best_oa_location") or {}
    return (
        primary_location.get("landing_page_url")
        or best_oa_location.get("landing_page_url")
        or ""
    )


def get_openalex_pdf_url(work: Dict[str, Any]) -> str:
    primary_location = work.get("primary_location") or {}
    best_oa_location = work.get("best_oa_location") or {}
    return (
        primary_location.get("pdf_url")
        or best_oa_location.get("pdf_url")
        or ""
    )


def is_openalex_open_access(work: Dict[str, Any]) -> bool:
    open_access = work.get("open_access") or {}
    return bool(open_access.get("is_oa", False))


def parse_openalex_work(work: Dict[str, Any], segmento_temporal: str) -> Dict[str, Any]:
    raw_type = work.get("type") or ""
    source_type = get_openalex_source_type(work)
    doi = get_openalex_doi(work)

    return {
        "title": work.get("display_name") or "",
        "authors": extract_openalex_authors(work.get("authorships")),
        "year": work.get("publication_year"),
        "type": raw_type,
        "tipo_documento": classify_openalex_document_type(raw_type),
        "source": "OpenAlex",
        "provider": "openalex",
        "source_name": get_openalex_source_name(work),
        "source_type": source_type,
        "tipo_fuente": classify_source_type(source_type),
        "revista_o_fuente": get_openalex_source_name(work),
        "doi": doi,
        "doi_url": get_doi_url(doi),
        "citations": work.get("cited_by_count", 0),
        "url": work.get("id") or "",
        "landing_page_url": get_openalex_landing_page_url(work),
        "pdf_url": get_openalex_pdf_url(work),
        "publication_date": work.get("publication_date") or "",
        "language": work.get("language") or "",
        "open_access": is_openalex_open_access(work),
        "openalex_id": work.get("id") or "",
        "paper_id": "",
        "publication_types": "",
        "abstract": "",
        "segmento_temporal": segmento_temporal,
    }


def source_passes_filters_openalex(
    item: Dict[str, Any],
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> bool:
    source_type = normalize_text(item.get("source_type") or "")

    if solo_revistas and source_type != "journal":
        return False

    if not incluir_repositorios and source_type == "repository":
        return False

    return document_type_policy_passes(
        item,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
    )


def build_openalex_filter(
    recent: bool,
    cutoff_year: int,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tipo_documento: Optional[str] = None,
) -> str:
    requested_types = get_requested_openalex_types(
        tipo_documento=tipo_documento,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
    )

    book_like = tipo_documento_is_book_like(
        tipo_documento=tipo_documento,
        incluir_libros=incluir_libros,
    )

    filters: List[str] = [
        f"type:{'|'.join(requested_types)}",
        "is_paratext:false",
    ]

    if not book_like:
        filters.append("has_abstract:true")

    effective_year_from = year_from
    effective_year_to = year_to

    if effective_year_from is None and effective_year_to is None:
        if recent:
            effective_year_from = cutoff_year
        else:
            effective_year_to = cutoff_year - 1

    if effective_year_from is not None:
        filters.append(f"from_publication_date:{effective_year_from}-01-01")

    if effective_year_to is not None:
        filters.append(f"to_publication_date:{effective_year_to}-12-31")

    if solo_revistas and not book_like:
        filters.append("primary_location.source.type:journal")
    else:
        if book_like:
            allowed_source_types = ["ebook-platform", "book-series", "repository", "journal"]
        else:
            allowed_source_types = ["journal"]
            if incluir_repositorios:
                allowed_source_types.append("repository")

        filters.append(f"primary_location.source.type:{'|'.join(allowed_source_types)}")

    return ",".join(filters)


async def fetch_openalex_count_segment(
    client: httpx.AsyncClient,
    query: str,
    recent: bool,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tipo_documento: Optional[str] = None,
) -> int:
    current_year = datetime.utcnow().year
    cutoff_year = current_year - RECENT_WINDOW_YEARS

    params = {
        "search": query,
        "filter": build_openalex_filter(
            recent=recent,
            cutoff_year=cutoff_year,
            solo_revistas=solo_revistas,
            incluir_repositorios=incluir_repositorios,
            incluir_libros=incluir_libros,
            incluir_preprints=incluir_preprints,
            year_from=year_from,
            year_to=year_to,
            tipo_documento=tipo_documento,
        ),
        "per-page": 1,
    }

    response = await client.get(OPENALEX_BASE_URL, params=params)
    response.raise_for_status()
    data = response.json()
    return int((data.get("meta") or {}).get("count", 0))


async def fetch_openalex_results_segment(
    client: httpx.AsyncClient,
    query: str,
    fetch_size: int,
    segmento: str,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tipo_documento: Optional[str] = None,
) -> List[Dict[str, Any]]:
    current_year = datetime.utcnow().year
    cutoff_year = current_year - RECENT_WINDOW_YEARS
    recent = segmento == "reciente"

    params = {
        "search": query,
        "filter": build_openalex_filter(
            recent=recent,
            cutoff_year=cutoff_year,
            solo_revistas=solo_revistas,
            incluir_repositorios=incluir_repositorios,
            incluir_libros=incluir_libros,
            incluir_preprints=incluir_preprints,
            year_from=year_from,
            year_to=year_to,
            tipo_documento=tipo_documento,
        ),
        "per-page": fetch_size,
        "sort": "relevance_score:desc",
    }

    response = await client.get(OPENALEX_BASE_URL, params=params)
    response.raise_for_status()
    data = response.json()

    works = data.get("results") or []
    parsed = [parse_openalex_work(work, segmento_temporal=segmento) for work in works]
    parsed = [
        item
        for item in parsed
        if source_passes_filters_openalex(
            item,
            solo_revistas=solo_revistas,
            incluir_repositorios=incluir_repositorios,
            incluir_libros=incluir_libros,
            incluir_preprints=incluir_preprints,
        )
    ]
    return parsed


async def search_openalex_provider(
    query: str,
    target_recent: int,
    target_classic: int,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    strict_relevance: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tipo_documento: Optional[str] = None,
) -> Dict[str, Any]:
    book_like = tipo_documento_is_book_like(
        tipo_documento=tipo_documento,
        incluir_libros=incluir_libros,
    )

    recent_fetch_size = max(target_recent * (8 if book_like else 6), 40 if book_like else 30) if target_recent > 0 else 0
    classic_fetch_size = max(target_classic * (8 if book_like else 4), 40 if book_like else 20) if target_classic > 0 else 0

    headers = {"User-Agent": f"{APP_NAME}/{APP_VERSION}"}
    params = {}
    if OPENALEX_EMAIL:
        params["mailto"] = OPENALEX_EMAIL

    total_recent = 0
    total_classic = 0
    recent_results: List[Dict[str, Any]] = []
    classic_results: List[Dict[str, Any]] = []

    async with httpx.AsyncClient(timeout=40.0, headers=headers, params=params) as client:
        try:
            tasks = []
            if target_recent > 0:
                tasks.append(
                    asyncio.gather(
                        fetch_openalex_count_segment(
                            client,
                            query=query,
                            recent=True,
                            solo_revistas=solo_revistas,
                            incluir_repositorios=incluir_repositorios,
                            incluir_libros=incluir_libros,
                            incluir_preprints=incluir_preprints,
                            year_from=year_from,
                            year_to=year_to,
                            tipo_documento=tipo_documento,
                        ),
                        fetch_openalex_results_segment(
                            client,
                            query=query,
                            fetch_size=recent_fetch_size,
                            segmento="reciente",
                            solo_revistas=solo_revistas,
                            incluir_repositorios=incluir_repositorios,
                            incluir_libros=incluir_libros,
                            incluir_preprints=incluir_preprints,
                            year_from=year_from,
                            year_to=year_to,
                            tipo_documento=tipo_documento,
                        ),
                    )
                )
            if target_classic > 0:
                tasks.append(
                    asyncio.gather(
                        fetch_openalex_count_segment(
                            client,
                            query=query,
                            recent=False,
                            solo_revistas=solo_revistas,
                            incluir_repositorios=incluir_repositorios,
                            incluir_libros=incluir_libros,
                            incluir_preprints=incluir_preprints,
                            year_from=year_from,
                            year_to=year_to,
                            tipo_documento=tipo_documento,
                        ),
                        fetch_openalex_results_segment(
                            client,
                            query=query,
                            fetch_size=classic_fetch_size,
                            segmento="clasico",
                            solo_revistas=solo_revistas,
                            incluir_repositorios=incluir_repositorios,
                            incluir_libros=incluir_libros,
                            incluir_preprints=incluir_preprints,
                            year_from=year_from,
                            year_to=year_to,
                            tipo_documento=tipo_documento,
                        ),
                    )
                )

            task_results = await asyncio.gather(*tasks)

        except httpx.HTTPError as e:
            raise HTTPException(
                status_code=502,
                detail=f"Error al consultar OpenAlex: {str(e)}"
            ) from e

    idx = 0
    if target_recent > 0:
        total_recent, recent_results = task_results[idx]
        idx += 1
        recent_results = apply_strict_title_relevance(
            recent_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        recent_results = deduplicate_results(recent_results)

    if target_classic > 0:
        total_classic, classic_results = task_results[idx]
        classic_results = apply_strict_title_relevance(
            classic_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        classic_results = deduplicate_results(classic_results)

    return {
        "provider": "openalex",
        "source_label": "OpenAlex",
        "total_recent": total_recent,
        "total_classic": total_classic,
        "total_found": total_recent + total_classic,
        "total_is_estimated": False,
        "recent_results": recent_results,
        "classic_results": classic_results,
    }


# =========================================================
# Crossref
# =========================================================

def classify_crossref_document_type(doc_type: Optional[str]) -> str:
    if not doc_type:
        return "Desconocido"
    doc_type_norm = normalize_text(doc_type)
    return DOCUMENT_TYPE_MAP.get(doc_type_norm, DOCUMENT_TYPE_MAP.get(doc_type, doc_type.replace("-", " ").capitalize()))


def get_crossref_title(item: Dict[str, Any]) -> str:
    titles = item.get("title") or []
    if isinstance(titles, list) and titles:
        return str(titles[0])
    if isinstance(titles, str):
        return titles
    return ""


def get_crossref_authors(item: Dict[str, Any]) -> str:
    authors = item.get("author") or []
    names: List[str] = []
    for author in authors:
        given = str(author.get("given") or "").strip()
        family = str(author.get("family") or "").strip()
        full = " ".join([part for part in [given, family] if part]).strip()
        if full:
            names.append(full)
    return "; ".join(names)


def get_crossref_source_name(item: Dict[str, Any]) -> str:
    container = item.get("container-title") or []
    if isinstance(container, list) and container:
        return str(container[0])
    if isinstance(container, str):
        return container

    publisher = item.get("publisher") or ""
    return publisher or "Sin fuente identificada"


def infer_crossref_source_type(item: Dict[str, Any]) -> str:
    doc_type = normalize_text(item.get("type") or "")
    container = get_crossref_source_name(item)

    if doc_type in {"journal-article", "review-article"}:
        return "journal"
    if doc_type in {"proceedings-article"}:
        return "conference"
    if doc_type in {"book", "monograph", "edited-book", "reference-book", "book-set"}:
        return "book-series"
    if doc_type in {"book-chapter", "book-section"}:
        return "book-series"
    if container:
        return "journal"
    return "unknown"


def get_crossref_year(item: Dict[str, Any]) -> Optional[int]:
    date_fields = [
        item.get("published-print"),
        item.get("published-online"),
        item.get("issued"),
        item.get("created"),
    ]

    for field in date_fields:
        if isinstance(field, dict):
            parts = field.get("date-parts") or []
            if parts and isinstance(parts, list) and isinstance(parts[0], list) and parts[0]:
                year = safe_int(parts[0][0], 0)
                if year > 0:
                    return year
    return None


def get_crossref_publication_date(item: Dict[str, Any]) -> str:
    date_fields = [
        item.get("published-print"),
        item.get("published-online"),
        item.get("issued"),
    ]

    for field in date_fields:
        if isinstance(field, dict):
            parts = field.get("date-parts") or []
            if parts and isinstance(parts, list) and isinstance(parts[0], list) and parts[0]:
                nums = [str(x) for x in parts[0][:3]]
                return "-".join(nums)
    return ""


def parse_crossref_item(item: Dict[str, Any], segmento_temporal: str) -> Dict[str, Any]:
    raw_type = item.get("type") or ""
    source_type = infer_crossref_source_type(item)
    doi = normalize_doi(item.get("DOI") or "")
    doi_url = get_doi_url(doi)
    landing_page = ""
    resource = item.get("resource") or {}
    if isinstance(resource, dict):
        landing_page = resource.get("primary", {}).get("URL") or ""

    if not landing_page:
        landing_page = doi_url

    language = str(item.get("language") or "").strip().lower()

    return {
        "title": get_crossref_title(item),
        "authors": get_crossref_authors(item),
        "year": get_crossref_year(item),
        "type": raw_type,
        "tipo_documento": classify_crossref_document_type(raw_type),
        "source": "Crossref",
        "provider": "crossref",
        "source_name": get_crossref_source_name(item),
        "source_type": source_type,
        "tipo_fuente": classify_source_type(source_type),
        "revista_o_fuente": get_crossref_source_name(item),
        "doi": doi,
        "doi_url": doi_url,
        "citations": safe_int(item.get("is-referenced-by-count"), 0),
        "url": doi_url or landing_page,
        "landing_page_url": landing_page or doi_url,
        "pdf_url": "",
        "publication_date": get_crossref_publication_date(item),
        "language": language,
        "open_access": False,
        "openalex_id": "",
        "paper_id": "",
        "publication_types": raw_type,
        "abstract": "",
        "segmento_temporal": segmento_temporal,
    }


def crossref_item_passes_filters(
    item: Dict[str, Any],
    solo_revistas: bool = False,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> bool:
    source_type = normalize_text(item.get("source_type") or "")
    if solo_revistas and source_type != "journal":
        return False

    return document_type_policy_passes(
        item,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
    )


def get_requested_crossref_types(
    tipo_documento: Optional[str],
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> List[str]:
    td = normalize_tipo_documento_canonical(tipo_documento)

    if td == "articulo":
        return ["journal-article"]
    if td == "revision":
        return ["journal-article"]
    if td == "tesis":
        return ["dissertation"]
    if td == "libro":
        return ["book", "monograph", "edited-book", "reference-book", "book-set", "book-chapter", "book-section"]
    if td == "capitulo_libro":
        return ["book-chapter", "book-section"]
    if td == "preprint":
        return ["posted-content"]

    allowed = ["journal-article", "dissertation"]

    if incluir_libros:
        allowed.extend(["book", "monograph", "edited-book", "reference-book", "book-set", "book-chapter", "book-section"])

    if incluir_preprints:
        allowed.append("posted-content")

    return allowed


async def fetch_crossref_segment(
    client: httpx.AsyncClient,
    query: str,
    recent: bool,
    target_results: int,
    solo_revistas: bool = False,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tipo_documento: Optional[str] = None,
) -> Tuple[int, List[Dict[str, Any]]]:
    current_year = datetime.utcnow().year
    cutoff_year = current_year - RECENT_WINDOW_YEARS

    effective_year_from = year_from
    effective_year_to = year_to

    if effective_year_from is None and effective_year_to is None:
        if recent:
            effective_year_from = cutoff_year
        else:
            effective_year_to = cutoff_year - 1

    filters: List[str] = []
    if effective_year_from is not None:
        filters.append(f"from-pub-date:{effective_year_from}")
    if effective_year_to is not None:
        filters.append(f"until-pub-date:{effective_year_to}")

    requested_types = get_requested_crossref_types(
        tipo_documento=tipo_documento,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
    )
    for t in requested_types:
        filters.append(f"type:{t}")

    filter_value = ",".join(filters)
    desired_pool = min(max(target_results * 6, 30), 100)
    rows = min(desired_pool, 100)
    offset = 0
    total_found = 0
    collected: List[Dict[str, Any]] = []
    segmento = "reciente" if recent else "clasico"
    max_loops = 3

    for loop in range(max_loops):
        params: Dict[str, Any] = {
            "query.bibliographic": query,
            "rows": rows,
            "offset": offset,
            "sort": "relevance",
            "order": "desc",
        }
        if filter_value:
            params["filter"] = filter_value
        if CROSSREF_MAILTO:
            params["mailto"] = CROSSREF_MAILTO

        response = await client.get(CROSSREF_BASE_URL, params=params)
        response.raise_for_status()
        payload = response.json()
        message = payload.get("message", {}) if isinstance(payload, dict) else {}

        if loop == 0:
            total_found = safe_int(message.get("total-results"), 0)

        items = message.get("items") or []
        if not items:
            break

        parsed = [parse_crossref_item(item, segmento_temporal=segmento) for item in items]
        parsed = [
            item
            for item in parsed
            if crossref_item_passes_filters(
                item,
                solo_revistas=solo_revistas,
                incluir_libros=incluir_libros,
                incluir_preprints=incluir_preprints,
            )
        ]

        collected.extend(parsed)
        collected = deduplicate_results(collected)

        if len(collected) >= desired_pool:
            break

        if len(items) < rows:
            break

        offset += rows
        if offset >= total_found:
            break

    return total_found, collected[:desired_pool]


async def search_crossref_provider(
    query: str,
    target_recent: int,
    target_classic: int,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    strict_relevance: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tipo_documento: Optional[str] = None,
) -> Dict[str, Any]:
    headers = {
        "User-Agent": f"{APP_NAME}/{APP_VERSION} (mailto:{CROSSREF_MAILTO or 'anonymous@example.com'})"
    }

    total_recent = 0
    total_classic = 0
    recent_results: List[Dict[str, Any]] = []
    classic_results: List[Dict[str, Any]] = []

    async with httpx.AsyncClient(timeout=40.0, headers=headers) as client:
        try:
            tasks = []
            if target_recent > 0:
                tasks.append(
                    fetch_crossref_segment(
                        client,
                        query=query,
                        recent=True,
                        target_results=target_recent,
                        solo_revistas=solo_revistas,
                        incluir_libros=incluir_libros,
                        incluir_preprints=incluir_preprints,
                        year_from=year_from,
                        year_to=year_to,
                        tipo_documento=tipo_documento,
                    )
                )
            if target_classic > 0:
                tasks.append(
                    fetch_crossref_segment(
                        client,
                        query=query,
                        recent=False,
                        target_results=target_classic,
                        solo_revistas=solo_revistas,
                        incluir_libros=incluir_libros,
                        incluir_preprints=incluir_preprints,
                        year_from=year_from,
                        year_to=year_to,
                        tipo_documento=tipo_documento,
                    )
                )

            task_results = await asyncio.gather(*tasks)

        except httpx.HTTPError as e:
            raise HTTPException(
                status_code=502,
                detail=f"Error al consultar Crossref: {str(e)}"
            ) from e

    idx = 0
    if target_recent > 0:
        total_recent, recent_results = task_results[idx]
        idx += 1
        recent_results = apply_strict_title_relevance(
            recent_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        recent_results = deduplicate_results(recent_results)

    if target_classic > 0:
        total_classic, classic_results = task_results[idx]
        classic_results = apply_strict_title_relevance(
            classic_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        classic_results = deduplicate_results(classic_results)

    return {
        "provider": "crossref",
        "source_label": "Crossref",
        "total_recent": total_recent,
        "total_classic": total_classic,
        "total_found": total_recent + total_classic,
        "total_is_estimated": False,
        "recent_results": recent_results,
        "classic_results": classic_results,
    }


# =========================================================
# Semantic Scholar
# =========================================================

def classify_semantic_document_type(publication_types: Optional[List[str]]) -> str:
    if not publication_types:
        return "Desconocido"

    normalized = [normalize_text(p) for p in publication_types]
    for item in normalized:
        if item in SEMANTIC_PUBLICATION_TYPE_MAP:
            return SEMANTIC_PUBLICATION_TYPE_MAP[item]

    return publication_types[0]


def semantic_publication_types_to_string(publication_types: Optional[List[str]]) -> str:
    if not publication_types:
        return ""
    return "; ".join(publication_types)


def extract_semantic_authors(authors: Optional[List[Dict[str, Any]]]) -> str:
    if not authors:
        return ""
    names = []
    for author in authors:
        name = author.get("name")
        if name:
            names.append(name)
    return "; ".join(names)


def get_semantic_doi(external_ids: Optional[Dict[str, Any]]) -> str:
    if not external_ids:
        return ""
    for key, value in external_ids.items():
        if normalize_text(key) == "doi" and value:
            return normalize_doi(str(value))
    return ""


def get_semantic_pdf_url(open_access_pdf: Any) -> str:
    if isinstance(open_access_pdf, dict):
        return open_access_pdf.get("url") or ""
    return ""


def infer_semantic_source_type(publication_types: Optional[List[str]], venue: str) -> str:
    normalized = [normalize_text(p) for p in (publication_types or [])]
    if "journalarticle" in normalized or "review" in normalized:
        return "journal"
    if "conference" in normalized or "conferencepaper" in normalized:
        return "conference"
    if venue:
        return "journal"
    return "unknown"


def parse_semantic_paper(paper: Dict[str, Any], segmento_temporal: str) -> Dict[str, Any]:
    publication_types = paper.get("publicationTypes") or []
    venue = paper.get("venue") or "Sin fuente identificada"
    doi = get_semantic_doi(paper.get("externalIds"))
    pdf_url = get_semantic_pdf_url(paper.get("openAccessPdf"))
    source_type = infer_semantic_source_type(publication_types, venue)

    return {
        "title": paper.get("title") or "",
        "authors": extract_semantic_authors(paper.get("authors")),
        "year": paper.get("year"),
        "type": semantic_publication_types_to_string(publication_types),
        "tipo_documento": classify_semantic_document_type(publication_types),
        "source": "SemanticScholar",
        "provider": "semanticscholar",
        "source_name": venue,
        "source_type": source_type,
        "tipo_fuente": classify_source_type(source_type),
        "revista_o_fuente": venue,
        "doi": doi,
        "doi_url": get_doi_url(doi),
        "citations": paper.get("citationCount", 0),
        "url": paper.get("url") or "",
        "landing_page_url": paper.get("url") or "",
        "pdf_url": pdf_url,
        "publication_date": paper.get("publicationDate") or "",
        "language": "",
        "open_access": bool(pdf_url),
        "openalex_id": "",
        "paper_id": paper.get("paperId") or "",
        "publication_types": semantic_publication_types_to_string(publication_types),
        "abstract": paper.get("abstract") or "",
        "segmento_temporal": segmento_temporal,
    }


def build_semantic_year_filter(recent: bool, cutoff_year: int) -> str:
    if recent:
        return f"{cutoff_year}-"
    return f"1900-{cutoff_year - 1}"


def semantic_item_passes_filters(
    item: Dict[str, Any],
    solo_revistas: bool = False,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> bool:
    source_type = normalize_text(item.get("source_type") or "")
    if solo_revistas and source_type != "journal":
        return False

    return document_type_policy_passes(
        item,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
    )


async def fetch_semantic_segment(
    client: httpx.AsyncClient,
    query: str,
    recent: bool,
    target_results: int,
    solo_revistas: bool = False,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> Tuple[int, List[Dict[str, Any]]]:
    current_year = datetime.utcnow().year
    cutoff_year = current_year - RECENT_WINDOW_YEARS
    year_filter = build_semantic_year_filter(recent=recent, cutoff_year=cutoff_year)
    segmento = "reciente" if recent else "clasico"

    params_base = {
        "query": query,
        "fields": SEMANTIC_SCHOLAR_FIELDS,
        "year": year_filter,
    }

    headers = {}
    if SEMANTIC_SCHOLAR_API_KEY:
        headers["x-api-key"] = SEMANTIC_SCHOLAR_API_KEY

    collected: List[Dict[str, Any]] = []
    token: Optional[str] = None
    total_estimated = 0
    max_loops = 5
    desired_pool = min(max(target_results * 5, 50), 500)

    for loop in range(max_loops):
        params = dict(params_base)
        if token:
            params["token"] = token

        response = await client.get(
            f"{SEMANTIC_SCHOLAR_BASE_URL}/paper/search/bulk",
            params=params,
            headers=headers,
        )
        response.raise_for_status()
        payload = response.json()

        if loop == 0:
            total_estimated = int(payload.get("total", 0) or 0)

        papers = payload.get("data") or []
        if not papers:
            break

        parsed = [parse_semantic_paper(paper, segmento_temporal=segmento) for paper in papers]
        parsed = [
            item
            for item in parsed
            if semantic_item_passes_filters(
                item,
                solo_revistas=solo_revistas,
                incluir_libros=incluir_libros,
                incluir_preprints=incluir_preprints,
            )
        ]
        collected.extend(parsed)
        collected = deduplicate_results(collected)

        if len(collected) >= desired_pool:
            break

        token = payload.get("token")
        if not token:
            break

    return total_estimated, collected[:desired_pool]


async def search_semantic_provider(
    query: str,
    target_recent: int,
    target_classic: int,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    strict_relevance: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
) -> Dict[str, Any]:
    total_recent = 0
    total_classic = 0
    recent_results: List[Dict[str, Any]] = []
    classic_results: List[Dict[str, Any]] = []

    async with httpx.AsyncClient(timeout=40.0) as client:
        try:
            tasks = []
            if target_recent > 0:
                tasks.append(
                    fetch_semantic_segment(
                        client,
                        query=query,
                        recent=True,
                        target_results=target_recent,
                        solo_revistas=solo_revistas,
                        incluir_libros=incluir_libros,
                        incluir_preprints=incluir_preprints,
                    )
                )
            if target_classic > 0:
                tasks.append(
                    fetch_semantic_segment(
                        client,
                        query=query,
                        recent=False,
                        target_results=target_classic,
                        solo_revistas=solo_revistas,
                        incluir_libros=incluir_libros,
                        incluir_preprints=incluir_preprints,
                    )
                )

            task_results = await asyncio.gather(*tasks)

        except httpx.HTTPError as e:
            raise HTTPException(
                status_code=502,
                detail=f"Error al consultar Semantic Scholar: {str(e)}"
            ) from e

    idx = 0
    if target_recent > 0:
        total_recent, recent_results = task_results[idx]
        idx += 1
        recent_results = apply_strict_title_relevance(
            recent_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        recent_results = deduplicate_results(recent_results)

    if target_classic > 0:
        total_classic, classic_results = task_results[idx]
        classic_results = apply_strict_title_relevance(
            classic_results,
            query=query,
            strict_relevance=strict_relevance,
        )
        classic_results = deduplicate_results(classic_results)

    return {
        "provider": "semanticscholar",
        "source_label": "SemanticScholar",
        "total_recent": total_recent,
        "total_classic": total_classic,
        "total_found": total_recent + total_classic,
        "total_is_estimated": True,
        "recent_results": recent_results,
        "classic_results": classic_results,
    }


# =========================================================
# Orquestación multi-fuente
# =========================================================

def finalize_single_source_results(
    source_payload: Dict[str, Any],
    modo: str,
    target_recent: int,
    target_classic: int,
    fuente: str,
    query: str,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    idioma: Optional[str] = None,
    tipo_documento: Optional[str] = None,
) -> Dict[str, Any]:
    recent_results = annotate_and_rank_results(source_payload["recent_results"], query=query)
    classic_results = annotate_and_rank_results(source_payload["classic_results"], query=query)

    recent_results = apply_advanced_filters(
        recent_results,
        year_from=year_from,
        year_to=year_to,
        idioma=idioma,
        tipo_documento=tipo_documento,
    )
    classic_results = apply_advanced_filters(
        classic_results,
        year_from=year_from,
        year_to=year_to,
        idioma=idioma,
        tipo_documento=tipo_documento,
    )

    if modo == "mixto":
        final_recent = recent_results[:target_recent]
        final_classic = classic_results[:target_classic]
        final_results = final_recent + final_classic
    elif modo == "recientes":
        final_recent = recent_results[:target_recent]
        final_classic = []
        final_results = final_recent
    else:
        final_recent = []
        final_classic = classic_results[:target_classic]
        final_results = final_classic

    payload = {
        "source": source_payload["source_label"],
        "fuente": fuente,
        "total_found": source_payload["total_found"],
        "returned": len(final_results),
        "recent_count": len(final_recent),
        "classic_count": len(final_classic),
        "results": final_results,
    }

    if source_payload.get("total_is_estimated"):
        payload["total_found_note"] = "El total_found de Semantic Scholar es estimado por su endpoint bulk search."
    return payload


def finalize_mixed_results(
    openalex_payload: Dict[str, Any],
    semantic_payload: Dict[str, Any],
    scopus_payload: Dict[str, Any],
    modo: str,
    target_recent: int,
    target_classic: int,
    query: str,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    idioma: Optional[str] = None,
    tipo_documento: Optional[str] = None,
) -> Dict[str, Any]:
    mixed_recent = deduplicate_results(
        interleave_multiple_lists(
            openalex_payload["recent_results"],
            semantic_payload["recent_results"],
            scopus_payload["recent_results"],
        )
    )

    mixed_classic = deduplicate_results(
        interleave_multiple_lists(
            openalex_payload["classic_results"],
            semantic_payload["classic_results"],
            scopus_payload["classic_results"],
        )
    )

    mixed_recent = annotate_and_rank_results(mixed_recent, query=query)
    mixed_classic = annotate_and_rank_results(mixed_classic, query=query)

    mixed_recent = apply_advanced_filters(
        mixed_recent,
        year_from=year_from,
        year_to=year_to,
        idioma=idioma,
        tipo_documento=tipo_documento,
    )
    mixed_classic = apply_advanced_filters(
        mixed_classic,
        year_from=year_from,
        year_to=year_to,
        idioma=idioma,
        tipo_documento=tipo_documento,
    )

    mixed_recent = balance_ranked_results_by_provider(mixed_recent)
    mixed_classic = balance_ranked_results_by_provider(mixed_classic)

    if modo == "mixto":
        final_recent = mixed_recent[:target_recent]
        final_classic = mixed_classic[:target_classic]
        final_results = final_recent + final_classic
    elif modo == "recientes":
        final_recent = mixed_recent[:target_recent]
        final_classic = []
        final_results = final_recent
    else:
        final_recent = []
        final_classic = mixed_classic[:target_classic]
        final_results = final_classic

    total_found = (
        openalex_payload["total_found"]
        + semantic_payload["total_found"]
        + scopus_payload["total_found"]
    )

    return {
        "source": "OpenAlex+SemanticScholar+Scopus",
        "fuente": "mixto",
        "total_found": total_found,
        "total_found_note": "En modo mixto, total_found suma OpenAlex, Scopus y el total estimado de Semantic Scholar antes de la deduplicación cruzada.",
        "returned": len(final_results),
        "recent_count": len(final_recent),
        "classic_count": len(final_classic),
        "source_breakdown": {
            "openalex_total_found": openalex_payload["total_found"],
            "semanticscholar_total_found_estimated": semantic_payload["total_found"],
            "scopus_total_found": scopus_payload["total_found"],
        },
        "results": final_results,
    }


async def run_search(
    query: str,
    max_results: int = 10,
    solo_revistas: bool = False,
    incluir_repositorios: bool = True,
    modo: str = "recientes",
    fuente: str = "openalex",
    strict_relevance: bool = True,
    incluir_libros: bool = False,
    incluir_preprints: bool = False,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    idioma: Optional[str] = None,
    tipo_documento: Optional[str] = None,
) -> Dict[str, Any]:
    if not query or not query.strip():
        raise HTTPException(status_code=400, detail="El parámetro 'query' es obligatorio.")

    if max_results < 1 or max_results > 100:
        raise HTTPException(status_code=400, detail="max_results debe estar entre 1 y 100.")

    if year_from is not None and (year_from < 1900 or year_from > 2100):
        raise HTTPException(status_code=400, detail="year_from debe estar entre 1900 y 2100.")

    if year_to is not None and (year_to < 1900 or year_to > 2100):
        raise HTTPException(status_code=400, detail="year_to debe estar entre 1900 y 2100.")

    if year_from is not None and year_to is not None and year_from > year_to:
        raise HTTPException(status_code=400, detail="year_from no puede ser mayor que year_to.")

    modo = normalize_text(modo)
    if modo not in MODOS_VALIDOS:
        raise HTTPException(
            status_code=400,
            detail="modo debe ser uno de: mixto, recientes, clasicos."
        )

    fuente = normalize_text(fuente)
    if fuente not in FUENTES_VALIDAS:
        raise HTTPException(
            status_code=400,
            detail="fuente debe ser uno de: openalex, semanticscholar, scopus, crossref, mixto."
        )

    effective_incluir_libros = incluir_libros or tipo_documento_requires_books(tipo_documento)
    effective_incluir_preprints = incluir_preprints or tipo_documento_requires_preprints(tipo_documento)

    target_recent, target_classic = split_targets(max_results=max_results, modo=modo)

    if fuente == "openalex":
        openalex_payload = await search_openalex_provider(
            query=query,
            target_recent=target_recent,
            target_classic=target_classic,
            solo_revistas=solo_revistas,
            incluir_repositorios=incluir_repositorios,
            strict_relevance=strict_relevance,
            incluir_libros=effective_incluir_libros,
            incluir_preprints=effective_incluir_preprints,
            year_from=year_from,
            year_to=year_to,
            tipo_documento=tipo_documento,
        )

        base = finalize_single_source_results(
            source_payload=openalex_payload,
            modo=modo,
            target_recent=target_recent,
            target_classic=target_classic,
            fuente=fuente,
            query=query,
            year_from=year_from,
            year_to=year_to,
            idioma=idioma,
            tipo_documento=tipo_documento,
        )

    elif fuente == "semanticscholar":
        semantic_payload = await search_semantic_provider(
            query=query,
            target_recent=target_recent,
            target_classic=target_classic,
            solo_revistas=solo_revistas,
            incluir_repositorios=incluir_repositorios,
            strict_relevance=strict_relevance,
            incluir_libros=effective_incluir_libros,
            incluir_preprints=effective_incluir_preprints,
        )

        base = finalize_single_source_results(
            source_payload=semantic_payload,
            modo=modo,
            target_recent=target_recent,
            target_classic=target_classic,
            fuente=fuente,
            query=query,
            year_from=year_from,
            year_to=year_to,
            idioma=idioma,
            tipo_documento=tipo_documento,
        )

    elif fuente == "scopus":
        scopus_payload = await search_scopus_provider(
            query=query,
            target_recent=target_recent,
            target_classic=target_classic,
            solo_revistas=solo_revistas,
            incluir_repositorios=incluir_repositorios,
            strict_relevance=strict_relevance,
            incluir_libros=effective_incluir_libros,
            incluir_preprints=effective_incluir_preprints,
        )

        base = finalize_single_source_results(
            source_payload=scopus_payload,
            modo=modo,
            target_recent=target_recent,
            target_classic=target_classic,
            fuente=fuente,
            query=query,
            year_from=year_from,
            year_to=year_to,
            idioma=idioma,
            tipo_documento=tipo_documento,
        )

    elif fuente == "crossref":
        crossref_payload = await search_crossref_provider(
            query=query,
            target_recent=target_recent,
            target_classic=target_classic,
            solo_revistas=solo_revistas,
            incluir_repositorios=incluir_repositorios,
            strict_relevance=strict_relevance,
            incluir_libros=effective_incluir_libros,
            incluir_preprints=effective_incluir_preprints,
            year_from=year_from,
            year_to=year_to,
            tipo_documento=tipo_documento,
        )

        base = finalize_single_source_results(
            source_payload=crossref_payload,
            modo=modo,
            target_recent=target_recent,
            target_classic=target_classic,
            fuente=fuente,
            query=query,
            year_from=year_from,
            year_to=year_to,
            idioma=idioma,
            tipo_documento=tipo_documento,
        )

    else:
        openalex_payload, semantic_payload, scopus_payload = await asyncio.gather(
            search_openalex_provider(
                query=query,
                target_recent=target_recent,
                target_classic=target_classic,
                solo_revistas=solo_revistas,
                incluir_repositorios=incluir_repositorios,
                strict_relevance=strict_relevance,
                incluir_libros=effective_incluir_libros,
                incluir_preprints=effective_incluir_preprints,
                year_from=year_from,
                year_to=year_to,
                tipo_documento=tipo_documento,
            ),
            search_semantic_provider(
                query=query,
                target_recent=target_recent,
                target_classic=target_classic,
                solo_revistas=solo_revistas,
                incluir_repositorios=incluir_repositorios,
                strict_relevance=strict_relevance,
                incluir_libros=effective_incluir_libros,
                incluir_preprints=effective_incluir_preprints,
            ),
            search_scopus_provider(
                query=query,
                target_recent=target_recent,
                target_classic=target_classic,
                solo_revistas=solo_revistas,
                incluir_repositorios=incluir_repositorios,
                strict_relevance=strict_relevance,
                incluir_libros=effective_incluir_libros,
                incluir_preprints=effective_incluir_preprints,
            ),
        )

        base = finalize_mixed_results(
            openalex_payload=openalex_payload,
            semantic_payload=semantic_payload,
            scopus_payload=scopus_payload,
            modo=modo,
            target_recent=target_recent,
            target_classic=target_classic,
            query=query,
            year_from=year_from,
            year_to=year_to,
            idioma=idioma,
            tipo_documento=tipo_documento,
        )

    return {
        "query": query,
        "solo_revistas": solo_revistas,
        "incluir_repositorios": incluir_repositorios,
        "incluir_libros": effective_incluir_libros,
        "incluir_preprints": effective_incluir_preprints,
        "modo": modo,
        "strict_relevance": strict_relevance,
        "year_from": year_from,
        "year_to": year_to,
        "idioma": idioma,
        "tipo_documento": tipo_documento,
        "filters_applied": {
            "solo_revistas": solo_revistas,
            "incluir_repositorios": incluir_repositorios,
            "incluir_libros": effective_incluir_libros,
            "incluir_preprints": effective_incluir_preprints,
            "modo": modo,
            "fuente": fuente,
            "strict_relevance": strict_relevance,
            "year_from": year_from,
            "year_to": year_to,
            "idioma": idioma,
            "tipo_documento": tipo_documento,
        },
        "policy_note": "Por defecto, la API prioriza artículos, tesis y revisiones de los últimos 5 años. Libros, preprints y clásicos solo se incluyen cuando el usuario lo solicita.",
        **base,
    }


# =========================================================
# Excel
# =========================================================

RESULTS_COLUMNS = [
    "title",
    "authors",
    "year",
    "tipo_documento",
    "tipo_fuente",
    "source",
    "source_name",
    "revista_o_fuente",
    "segmento_temporal",
    "score_relevancia",
    "match_keywords_title",
    "doi",
    "citations",
    "landing_page_url",
    "pdf_url",
    "url",
    "open_access",
]

FULL_METADATA_COLUMNS = [
    "title",
    "authors",
    "year",
    "type",
    "tipo_documento",
    "source",
    "provider",
    "source_name",
    "source_type",
    "tipo_fuente",
    "revista_o_fuente",
    "score_relevancia",
    "match_keywords_title",
    "doi",
    "doi_url",
    "citations",
    "url",
    "landing_page_url",
    "pdf_url",
    "publication_date",
    "language",
    "open_access",
    "openalex_id",
    "paper_id",
    "publication_types",
    "abstract",
    "segmento_temporal",
]


def autosize_worksheet(ws) -> None:
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            try:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))
            except Exception:
                pass
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 60)


def format_header(ws) -> None:
    header_fill = PatternFill(fill_type="solid", start_color="1F4E78", end_color="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def add_clickable_links(ws) -> None:
    header_map = {cell.value: cell.column for cell in ws[1]}
    link_columns = ["landing_page_url", "pdf_url", "url", "doi_url"]

    for col_name in link_columns:
        col_idx = header_map.get(col_name)
        if not col_idx:
            continue

        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value and str(cell.value).startswith(("http://", "https://")):
                cell.hyperlink = str(cell.value)
                cell.style = "Hyperlink"


def build_statistics_df(payload: Dict[str, Any]) -> pd.DataFrame:
    results = payload.get("results", [])
    by_segment = {}
    by_doc_type = {}
    by_source_type = {}
    by_provider = {}

    for item in results:
        seg = item.get("segmento_temporal", "desconocido")
        by_segment[seg] = by_segment.get(seg, 0) + 1

        doc = item.get("tipo_documento", "Desconocido")
        by_doc_type[doc] = by_doc_type.get(doc, 0) + 1

        source = item.get("tipo_fuente", "Desconocido")
        by_source_type[source] = by_source_type.get(source, 0) + 1

        provider = item.get("provider", "desconocido")
        by_provider[provider] = by_provider.get(provider, 0) + 1

    rows = [
        {"categoria": "consulta", "clave": "query", "valor": payload.get("query")},
        {"categoria": "fuente", "clave": "source", "valor": payload.get("source")},
        {"categoria": "fuente", "clave": "fuente", "valor": payload.get("fuente")},
        {"categoria": "conteo", "clave": "total_found", "valor": payload.get("total_found")},
        {"categoria": "conteo", "clave": "returned", "valor": payload.get("returned")},
        {"categoria": "conteo", "clave": "recent_count", "valor": payload.get("recent_count")},
        {"categoria": "conteo", "clave": "classic_count", "valor": payload.get("classic_count")},
        {"categoria": "filtros", "clave": "solo_revistas", "valor": payload.get("solo_revistas")},
        {"categoria": "filtros", "clave": "incluir_repositorios", "valor": payload.get("incluir_repositorios")},
        {"categoria": "filtros", "clave": "incluir_libros", "valor": payload.get("incluir_libros")},
        {"categoria": "filtros", "clave": "incluir_preprints", "valor": payload.get("incluir_preprints")},
        {"categoria": "filtros", "clave": "modo", "valor": payload.get("modo")},
        {"categoria": "filtros", "clave": "strict_relevance", "valor": payload.get("strict_relevance")},
        {"categoria": "filtros", "clave": "year_from", "valor": payload.get("year_from")},
        {"categoria": "filtros", "clave": "year_to", "valor": payload.get("year_to")},
        {"categoria": "filtros", "clave": "idioma", "valor": payload.get("idioma")},
        {"categoria": "filtros", "clave": "tipo_documento", "valor": payload.get("tipo_documento")},
    ]

    if payload.get("policy_note"):
        rows.append({"categoria": "nota", "clave": "policy_note", "valor": payload.get("policy_note")})

    if payload.get("total_found_note"):
        rows.append({"categoria": "nota", "clave": "total_found_note", "valor": payload.get("total_found_note")})

    source_breakdown = payload.get("source_breakdown") or {}
    for key, value in source_breakdown.items():
        rows.append({"categoria": "source_breakdown", "clave": key, "valor": value})

    for key, value in by_segment.items():
        rows.append({"categoria": "segmento_temporal", "clave": key, "valor": value})

    for key, value in by_doc_type.items():
        rows.append({"categoria": "tipo_documento", "clave": key, "valor": value})

    for key, value in by_source_type.items():
        rows.append({"categoria": "tipo_fuente", "clave": key, "valor": value})

    for key, value in by_provider.items():
        rows.append({"categoria": "provider", "clave": key, "valor": value})

    return pd.DataFrame(rows)


def build_excel_response(payload: Dict[str, Any]) -> StreamingResponse:
    results = payload.get("results", [])
    results_df = pd.DataFrame(results)

    if results_df.empty:
        results_df = pd.DataFrame(columns=RESULTS_COLUMNS)
        full_df = pd.DataFrame(columns=FULL_METADATA_COLUMNS)
    else:
        full_df = results_df.copy()

        for col in RESULTS_COLUMNS:
            if col not in results_df.columns:
                results_df[col] = None

        for col in FULL_METADATA_COLUMNS:
            if col not in full_df.columns:
                full_df[col] = None

        results_df = results_df[RESULTS_COLUMNS]
        full_df = full_df[FULL_METADATA_COLUMNS]

    stats_df = build_statistics_df(payload)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        results_df.to_excel(writer, sheet_name="Resultados", index=False)
        full_df.to_excel(writer, sheet_name="Metadatos completos", index=False)
        stats_df.to_excel(writer, sheet_name="Estadísticas", index=False)

        wb = writer.book
        for sheet_name in ["Resultados", "Metadatos completos", "Estadísticas"]:
            ws = wb[sheet_name]
            format_header(ws)
            add_clickable_links(ws)
            autosize_worksheet(ws)

    output.seek(0)
    filename = f"{safe_filename(payload.get('query', 'consulta'))}_resultados.xlsx"

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# =========================================================
# Endpoints
# =========================================================

@app.get("/")
async def root() -> Dict[str, Any]:
    return {
        "message": "academic-search-api funcionando correctamente",
        "version": APP_VERSION,
        "policy_defaults": {
            "modo_default": "recientes",
            "ventana_reciente": f"últimos {RECENT_WINDOW_YEARS} años",
            "tipos_por_defecto": "artículos, tesis y revisiones",
            "idioma_por_defecto": "sin filtro",
            "libros_por_defecto": False,
            "preprints_por_defecto": False,
        },
        "endpoints": {
            "health": "/health",
            "search_openalex": "/search?query=telemedicina&max_results=10&fuente=openalex",
            "search_semanticscholar": "/search?query=telemedicina&max_results=10&fuente=semanticscholar",
            "search_scopus": "/search?query=telemedicina&max_results=10&fuente=scopus",
            "search_crossref": "/search?query=telemedicina&max_results=10&fuente=crossref",
            "search_mixto": "/search?query=telemedicina&max_results=10&fuente=mixto",
            "search_filtrado": "/search?query=telemedicina&max_results=10&fuente=mixto&tipo_documento=revision&year_from=2021&year_to=2026",
            "excel": "/export.xlsx?query=telemedicina&max_results=10",
            "scopus_test": "/search_scopus_test?query=telemedicine&max_results=10",
        },
        "params_disponibles": {
            "max_results": "1 a 100",
            "solo_revistas": "true | false",
            "incluir_repositorios": "true | false",
            "incluir_libros": "true | false",
            "incluir_preprints": "true | false",
            "modo": "recientes | clasicos | mixto",
            "fuente": "openalex | semanticscholar | scopus | crossref | mixto",
            "strict_relevance": "true | false",
            "year_from": "año mínimo opcional",
            "year_to": "año máximo opcional",
            "idioma": "es | en | pt | fr | de | it o equivalente textual",
            "tipo_documento": "articulo | revision | tesis | libro | capitulo_libro | preprint",
        },
    }


@app.get("/health")
async def health() -> Dict[str, Any]:
    return {
        "status": "ok",
        "app": APP_NAME,
        "version": APP_VERSION,
        "timestamp_utc": datetime.utcnow().isoformat(),
    }


@app.get("/search")
async def search(
    query: str = Query(..., description="Término de búsqueda"),
    max_results: int = Query(10, ge=1, le=100, description="Máximo total de resultados"),
    solo_revistas: bool = Query(False, description="Devuelve solo resultados de revistas científicas"),
    incluir_repositorios: bool = Query(True, description="Incluye resultados de repositorios"),
    incluir_libros: bool = Query(False, description="Incluye libros y capítulos de libro"),
    incluir_preprints: bool = Query(False, description="Incluye preprints"),
    modo: str = Query("recientes", description="recientes | clasicos | mixto"),
    fuente: str = Query("openalex", description="openalex | semanticscholar | scopus | crossref | mixto"),
    strict_relevance: bool = Query(True, description="Filtra resultados por coincidencia estricta en el título"),
    year_from: Optional[int] = Query(None, description="Año mínimo opcional"),
    year_to: Optional[int] = Query(None, description="Año máximo opcional"),
    idioma: Optional[str] = Query(None, description="Idioma opcional, por ejemplo: es, en, español, english"),
    tipo_documento: Optional[str] = Query(None, description="articulo, revision, tesis, libro, capitulo_libro, preprint"),
):
    payload = await run_search(
        query=query,
        max_results=max_results,
        solo_revistas=solo_revistas,
        incluir_repositorios=incluir_repositorios,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
        modo=modo,
        fuente=fuente,
        strict_relevance=strict_relevance,
        year_from=year_from,
        year_to=year_to,
        idioma=idioma,
        tipo_documento=tipo_documento,
    )
    return JSONResponse(content=payload)


@app.get("/export.xlsx")
async def export_excel(
    query: str = Query(..., description="Término de búsqueda"),
    max_results: int = Query(10, ge=1, le=100, description="Máximo total de resultados"),
    solo_revistas: bool = Query(False, description="Devuelve solo resultados de revistas científicas"),
    incluir_repositorios: bool = Query(True, description="Incluye resultados de repositorios"),
    incluir_libros: bool = Query(False, description="Incluye libros y capítulos de libro"),
    incluir_preprints: bool = Query(False, description="Incluye preprints"),
    modo: str = Query("recientes", description="recientes | clasicos | mixto"),
    fuente: str = Query("openalex", description="openalex | semanticscholar | scopus | crossref | mixto"),
    strict_relevance: bool = Query(True, description="Filtra resultados por coincidencia estricta en el título"),
    year_from: Optional[int] = Query(None, description="Año mínimo opcional"),
    year_to: Optional[int] = Query(None, description="Año máximo opcional"),
    idioma: Optional[str] = Query(None, description="Idioma opcional, por ejemplo: es, en, español, english"),
    tipo_documento: Optional[str] = Query(None, description="articulo, revision, tesis, libro, capitulo_libro, preprint"),
):
    payload = await run_search(
        query=query,
        max_results=max_results,
        solo_revistas=solo_revistas,
        incluir_repositorios=incluir_repositorios,
        incluir_libros=incluir_libros,
        incluir_preprints=incluir_preprints,
        modo=modo,
        fuente=fuente,
        strict_relevance=strict_relevance,
        year_from=year_from,
        year_to=year_to,
        idioma=idioma,
        tipo_documento=tipo_documento,
    )
    return build_excel_response(payload)


@app.get("/search_scopus_test")
async def search_scopus_test(
    query: str = Query(..., description="Texto a buscar en Scopus"),
    max_results: int = Query(10, ge=1, le=25),
    start: int = Query(0, ge=0),
):
    if not SCOPUS_API_KEY:
        return JSONResponse(
            status_code=500,
            content={
                "ok": False,
                "status_code": 500,
                "message": "No existe SCOPUS_API_KEY en variables de entorno.",
            },
        )

    headers = {
        "Accept": "application/json",
        "X-ELS-APIKey": SCOPUS_API_KEY,
    }

    params = {
        "query": build_scopus_test_query(query),
        "count": max_results,
        "start": start,
        "view": "STANDARD",
    }

    try:
        async with httpx.AsyncClient(timeout=45.0, headers=headers) as client:
            response = await client.get(
                SCOPUS_SEARCH_URL,
                params=params,
            )

        try:
            data = response.json()
        except Exception:
            data = {"raw_text": response.text}

        if response.status_code >= 400:
            return JSONResponse(
                status_code=response.status_code,
                content={
                    "ok": False,
                    "status_code": response.status_code,
                    "query_received": query,
                    "query_sent": params["query"],
                    "message": "Scopus devolvió un error.",
                    "response": data,
                },
            )

        search_results = data.get("search-results", {}) if isinstance(data, dict) else {}
        entries = search_results.get("entry", []) or []
        parsed_entries = [parse_scopus_entry(entry) for entry in entries]

        payload = {
            "ok": True,
            "status_code": response.status_code,
            "query_received": query,
            "query_sent": params["query"],
            "source": "Scopus",
            "fuente": "scopus_test",
            "total_found": safe_int(search_results.get("opensearch:totalResults"), 0),
            "returned": len(parsed_entries),
            "start": safe_int(search_results.get("opensearch:startIndex"), start),
            "items_per_page": safe_int(search_results.get("opensearch:itemsPerPage"), len(parsed_entries)),
            "results": parsed_entries,
            "raw_search_keys": list(search_results.keys()),
        }

        return JSONResponse(content=payload)

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "ok": False,
                "status_code": 500,
                "query_received": query,
                "query_sent": params["query"],
                "message": f"Error al consultar Scopus: {str(e)}",
            },
        )
