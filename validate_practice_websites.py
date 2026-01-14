#!/usr/bin/env python3
"""
Validierung von Arztpraxen-Websites basierend auf einer Excel-Datei

Requirements:
  pip install openpyxl pandas requests beautifulsoup4 rapidfuzz

Beispiel Nutzung:
  python validate_practice_websites.py --input input.xlsx --output output_validated.xlsx
"""

import argparse
import csv
import logging
import re
import time
import unicodedata
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from rapidfuzz import fuzz


# Configure logging
logging.basicConfig(
    filename='validation.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class RateLimiter:
    def __init__(self, rate_per_second: float = 1.0):
        self.rate = rate_per_second
        self.last_call = 0
        self.min_interval = 1.0 / rate_per_second if rate_per_second > 0 else 0
    
    def wait(self):
        now = time.time()
        elapsed = now - self.last_call
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self.last_call = time.time()


class ValidationResult:
    def __init__(self):
        self.validation_status: str = ""
        self.validation_reason: str = ""
        self.final_url: str = ""
        self.http_status: str = ""
        self.match_score_firstname: float = 0.0
        self.match_score_lastname: float = 0.0
        self.error_details: str = ""


def normalize_name(name: str) -> str:
    """Normalisiere Namen für besseres Matching"""
    if pd.isna(name) or not isinstance(name, str):
        return ""
    
    # Konvertiere zu String und lowercase
    name = str(name).lower().strip()
    
    # Unicode Normalisierung (NFKD) und Diakritika entfernen
    name = unicodedata.normalize('NFKD', name)
    name = ''.join(c for c in name if not unicodedata.combining(c))
    
    # Umlaute ersetzen
    replacements = {
        'ä': 'ae', 'ö': 'oe', 'ü': 'ue', 'ß': 'ss',
        'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'å': 'a',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o',
        'ú': 'u', 'ù': 'u', 'û': 'u',
        'ý': 'y', 'ÿ': 'y',
        'ñ': 'n', 'ç': 'c'
    }
    
    for src, dst in replacements.items():
        name = name.replace(src, dst)
    
    # Entferne nicht-alphanumerische Zeichen (außer Leerzeichen)
    name = re.sub(r'[^a-z0-9\s]', '', name)
    
    # Normalisiere Whitespace
    name = re.sub(r'\s+', ' ', name).strip()
    
    return name


def normalize_url(url: str) -> Optional[str]:
    """Normalisiere URL und ergänze Schema wenn nötig"""
    if pd.isna(url) or not isinstance(url, str) or url.strip() == "":
        return None
    
    url = url.strip()
    
    # Prüfe ob URL offensichtlich kaputt ist
    if not url or len(url) < 5 or ' ' in url:
        return None
    
    parsed = urlparse(url)
    
    # Sche das Schema hinzu wenn fehlt
    if not parsed.scheme:
        # Versuche zuerst https
        test_url = f"https://{url}"
        try:
            parsed_test = urlparse(test_url)
            if parsed_test.netloc:
                return test_url
        except:
            pass
        
        # Fallback auf http
        test_url = f"http://{url}"
        try:
            parsed_test = urlparse(test_url)
            if parsed_test.netloc:
                return test_url
        except:
            pass
        
        return None
    
    return url


def fetch_page(url: str, timeout: int = 15) -> Tuple[Optional[requests.Response], str]:
    """Abrufen der Webseite mit Fehlerbehandlung"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(
            url,
            timeout=timeout,
            headers=headers,
            allow_redirects=True,
            stream=True
        )
        
        # Prüfe Content-Type
        content_type = response.headers.get('content-type', '').lower()
        if 'text/html' not in content_type:
            return response, f"Content-Type ist kein HTML: {content_type}"
        
        # Limitiere Download-Größe (max 4 MB)
        content_length = response.headers.get('content-length')
        if content_length and int(content_length) > 4 * 1024 * 1024:
            return response, "Datei zu groß (> 4 MB)"
        
        return response, ""
        
    except requests.exceptions.Timeout:
        return None, "Timeout"
    except requests.exceptions.SSLError:
        return None, "SSL Error"
    except requests.exceptions.ConnectionError:
        return None, "Connection Error"
    except requests.exceptions.HTTPError as e:
        return None, f"HTTP Error: {str(e)}"
    except requests.exceptions.RequestException as e:
        return None, f"Request Error: {str(e)}"
    except Exception as e:
        return None, f"Unexpected Error: {str(e)}"


def extract_text(html_content: str) -> str:
    """Extrahiere sichtbaren Text aus HTML"""
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Entferne nicht-sichtbare Inhalte
        for tag in soup(['script', 'style', 'noscript', 'meta', 'link', 'head']):
            tag.decompose()
        
        # Extrahiere Text
        text = soup.get_text(separator=' ')
        
        # Normalisiere Whitespace
        text = re.sub(r'\s+', ' ', text)
        text = text.strip()
        
        return text.lower()
        
    except Exception as e:
        logger.error(f"Fehler beim Text-Extrahieren: {str(e)}")
        return ""


def compute_match(normalized_name: str, page_text: str) -> float:
    """Berechne Fuzzy-Match-Score"""
    if not normalized_name or not page_text:
        return 0.0
    
    # Token-Suche als erst
    tokens = normalized_name.split()
    if len(tokens) == 1:
        # Exakte Token-Suche
        if normalized_name in page_text:
            return 100.0
    else:
        # Prüfe ob alle Tokens enthalten sind
        all_tokens_found = all(token in page_text for token in tokens)
        if all_tokens_found:
            return 100.0
    
    # Fuzzy-Matching mit rapidfuzz
    try:
        score = fuzz.partial_ratio(normalized_name, page_text, score_cutoff=70.0)
        if score is None:
            score = 0.0
        return float(score)
    except Exception as e:
        logger.error(f"Fehler beim Fuzzy-Matching: {str(e)}")
        return 0.0


def validate_row(row: Dict, threshold: float = 95.0) -> ValidationResult:
    """Validiere eine einzelne Zeile"""
    result = ValidationResult()
    
    try:
        # Extrahiere Daten
        website = row.get('Website', '') if isinstance(row, dict) else row.get('Website')
        firstname = row.get('Vorname', '') if isinstance(row, dict) else row.get('Vorname')
        lastname = row.get('Nachname', '') if isinstance(row, dict) else row.get('Nachname')
        
        # Normalisiere URL
        normalized_url = normalize_url(website)
        if not normalized_url:
            result.validation_status = "unvalide"
            result.validation_reason = "Ungültige URL"
            return result
        
        result.final_url = normalized_url
        
        # Abrufen der Seite
        response, error_msg = fetch_page(normalized_url)
        
        if error_msg:
            result.validation_status = "nicht erreichbar"
            result.validation_reason = error_msg
            result.error_details = error_msg
            if response:
                result.http_status = str(response.status_code)
            return result
        
        if response is None:
            result.validation_status = "nicht erreichbar"
            result.validation_reason = "Unbekannter Fehler"
            return result
        
        result.http_status = str(response.status_code)
        
        if response.status_code != 200:
            result.validation_status = "nicht erreichbar"
            result.validation_reason = f"HTTP {response.status_code}"
            return result
        
        # HTML-Content lesen
        try:
            html_content = response.content.decode('utf-8', errors='ignore')
        except:
            html_content = response.text
        
        # Text extrahieren
        page_text = extract_text(html_content)
        
        if not page_text:
            result.validation_status = "unvalide"
            result.validation_reason = "Kein Text extrahiert"
            return result
        
        # Namen normalisieren
        normalized_firstname = normalize_name(firstname)
        normalized_lastname = normalize_name(lastname)
        
        if not normalized_lastname:
            result.validation_status = "unvalide"
            result.validation_reason = "Kein Nachname vorhanden"
            return result
        
        # Match-Scores berechnen
        result.match_score_lastname = compute_match(normalized_lastname, page_text)
        result.match_score_firstname = compute_match(normalized_firstname, page_text) if normalized_firstname else 0.0
        
        # Entscheidungslogik
        if result.match_score_lastname >= threshold:
            result.validation_status = "valide"
            result.validation_reason = f"Nachname gefunden (Score: {result.match_score_lastname:.1f}%)"
        elif result.match_score_firstname >= threshold and result.match_score_lastname >= threshold * 0.8:
            result.validation_status = "valide"
            result.validation_reason = f"Vorname und Nachname gefunden (Scores: {result.match_score_firstname:.1f}%, {result.match_score_lastname:.1f}%)"
        else:
            result.validation_status = "unvalide"
            result.validation_reason = f"Kein Name gefunden (Vorname: {result.match_score_firstname:.1f}%, Nachname: {result.match_score_lastname:.1f}%)"
        
        return result
        
    except Exception as e:
        logger.error(f"Fehler bei der Validierung: {str(e)}")
        result.validation_status = "unvalide"
        result.validation_reason = f"Validierungsfehler: {str(e)}"
        result.error_details = str(e)
        return result


def process_file(
    input_file: str,
    output_file: str,
    output_csv: str,
    errors_csv: str,
    sheet_name: str = 'Sheet1',
    threshold: float = 95.0,
    rate: float = 1.0,
    timeout: int = 15
):
    """Verarbeite die gesamte Excel-Datei"""
    logger.info(f"Starte Verarbeitung von {input_file}")
    
    try:
        # Lade Excel-Datei
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        logger.info(f"{len(df)} Zeilen geladen")
        
        # Rate Limiter
        rate_limiter = RateLimiter(rate)
        
        # Ergebnis-Listen
        results = []
        errors = []
        
        # Verarbeite jede Zeile
        for idx, row in df.iterrows():
            logger.info(f"Verarbeite Zeile {idx + 1}/{len(df)}")
            
            rate_limiter.wait()
            
            # Konvertiere row zu Dict falls nötig
            if hasattr(row, 'to_dict'):
                row_dict = row.to_dict()
            else:
                row_dict = dict(row)
            
            # Validierung
            result = validate_row(row_dict, threshold)
            
            # Kombiniere Original-Row mit Ergebnissen
            result_dict = row_dict.copy()
            result_dict.update({
                'validation_status': result.validation_status,
                'validation_reason': result.validation_reason,
                'final_url': result.final_url,
                'http_status': result.http_status,
                'match_score_firstname': result.match_score_firstname,
                'match_score_lastname': result.match_score_lastname,
                'error_details': result.error_details
            })
            
            results.append(result_dict)
            
            # Fehler sammeln
            if result.validation_status == "nicht erreichbar":
                error_row = row_dict.copy()
                error_row.update({
                    'final_url': result.final_url,
                    'error_reason': result.validation_reason,
                    'error_details': result.error_details
                })
                errors.append(error_row)
            
            logger.info(f"Zeile {idx + 1}: {result.validation_status} - {result.validation_reason}")
        
        # Erstelle Ergebnis DataFrame
        result_df = pd.DataFrame(results)
        
        # Exportiere Excel
        result_df.to_excel(output_file, index=False)
        logger.info(f"Ergebnis gespeichert: {output_file}")
        
        # Exportiere CSV
        result_df.to_csv(output_csv, index=False, quoting=csv.QUOTE_ALL)
        logger.info(f"Ergebnis gespeichert: {output_csv}")
        
        # Exportiere Fehler CSV
        if errors:
            error_df = pd.DataFrame(errors)
            
            # Wähle relevante Spalten für Fehler
            error_columns = [col for col in error_df.columns if col not in ['match_score_firstname', 'match_score_lastname']]
            error_df = error_df[error_columns]
            
            error_df.to_csv(errors_csv, index=False, quoting=csv.QUOTE_ALL)
            logger.info(f"Fehler gespeichert: {errors_csv} ({len(errors)} Einträge)")
        else:
            logger.info("Keine Fehler aufgetreten")
            # Leere Fehler-Datei mit Spaltennamen
            with open(errors_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['doctor_info', 'original_url', 'final_url', 'error_reason', 'error_details'])
        
        logger.info("Verarbeitung abgeschlossen")
        
    except Exception as e:
        logger.error(f"Fehler bei der Dateiverarbeitung: {str(e)}")
        raise


def main():
    """Main-Funktion mit CLI Argumenten"""
    parser = argparse.ArgumentParser(description='Validierung von Arztpraxen-Websites')
    
    parser.add_argument('--input', default='input.xlsx', help='Input Excel-Datei')
    parser.add_argument('--output', default='output_validated.xlsx', help='Output Excel-Datei')
    parser.add_argument('--sheet', default='Sheet1', help='Sheet-Name in Excel-Datei')
    parser.add_argument('--threshold', type=float, default=95.0, help='Matching-Schwellwert (0-100)')
    parser.add_argument('--rate', type=float, default=1.0, help='Rate Limit (requests pro Sekunde)')
    parser.add_argument('--timeout', type=int, default=15, help='Timeout in Sekunden')
    
    args = parser.parse_args()
    
    # Setze Output-Dateien
    output_csv = args.output.replace('.xlsx', '.csv')
    errors_csv = 'errors.csv'
    
    # Starte Verarbeitung
    process_file(
        input_file=args.input,
        output_file=args.output,
        output_csv=output_csv,
        errors_csv=errors_csv,
        sheet_name=args.sheet,
        threshold=args.threshold,
        rate=args.rate,
        timeout=args.timeout
    )


if __name__ == "__main__":
    main()
