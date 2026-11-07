# -*- coding: utf-8 -*-
# NVDA Add-on: Outlook Image Helper (Outlook - Pomocnik obrazów i OCR)
# Author: Sławomir Strugarek z pomocą Antigravity (Google DeepMind)
# Description: Detects, lists, copies, performs OCR, and generates AI descriptions/Q&A
# for inline and floating images/screenshots embedded in Microsoft Outlook Classic messages.

from __future__ import annotations
import array
import base64
import ctypes
from ctypes import byref, POINTER, c_void_p, oledll
import hashlib
import html
import json
import os
import platform
import re
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import shutil
import wave
import webbrowser
import weakref

import addonHandler
import api
import config
import controlTypes
import core
import speech
import globalPluginHandler
import globalVars
import gui
import gui.guiHelper
import gui.settingsDialogs
from logHandler import log
import nvwave
from scriptHandler import script
import ssl
import tones
import ui
import winUser
import wx
import wx.html
import zipfile
import xml.etree.ElementTree as ET

ROLE_TABLEROW = getattr(controlTypes.Role, "TABLEROW", getattr(controlTypes.Role, "ROW", None))

addonHandler.initTranslation()

ADDON_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_libDir = os.path.join(ADDON_DIR, "lib")
if _libDir not in sys.path:
	sys.path.insert(0, _libDir)

try:
	import pypdf
	log.info('outlookImageHelper: Biblioteka pypdf została pomyślnie zaimportowana.')
except Exception as _pypdfErr:
	pypdf = None
	log.error(f'outlookImageHelper: Nie udało się zaimportować biblioteki pypdf: {_pypdfErr}', exc_info=True)

CONFIG_SECTION = "outlookImageHelper"
ADDON_NAME = "Outlook Image Helper"
ADDON_TITLE = "Outlook - Pomocnik obrazów i załączników"
ADDON_AUTHOR = "Sławomir Strugarek z pomocą Antigravity (Google DeepMind)"
ADDON_BUILD_DATE = "2026-11-07"
ADDON_BUILD_ID = "20261107.2"
ADDON_COMPATIBILITY = "NVDA 2024.1 – 2026.1+ | Microsoft Outlook 2013-365 Classic"
ADDON_LICENSE = "GPL v2+ (Open Source)"
ADDON_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ADDON_VERSION = "1.6.5"
try:
	_manifest_path = os.path.join(ADDON_DIR, "manifest.ini")
	if os.path.isfile(_manifest_path):
		with open(_manifest_path, "r", encoding="utf-8") as _mf:
			for _line in _mf:
				if _line.strip().startswith("version"):
					_v = _line.split("=", 1)[1].strip().strip('"\'')
					if _v:
						ADDON_VERSION = _v
						break
except Exception:
	pass
SOUNDS_DIR = os.path.join(ADDON_DIR, "sounds")
ALERT_SOUND_PATH = os.path.join(SOUNDS_DIR, "image_alert.wav")
OCR_SCRIPT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ocr_worker.ps1")
SOUNDS_CACHE_DIR = os.path.join(tempfile.gettempdir(), "nvda_oih_sound_cache")
REPORTS_DIR_PRIMARY = None
REPORTS_DIR_MIRROR = None
MAX_REPORT_BYTES = 100 * 1024  # Maksymalny bezpieczny rozmiar raportu (100 KB)
MAX_REPORTS_RETAINED = 20      # Zachowuj maksymalnie 20 najnowszych raportów w folderze

# Bufor kołowy ostatnich zdarzeń audio (maks. 30 wpisów) oraz podręczny cache weryfikacji grafik na liście
_sound_events_history: list[dict[str, str]] = []
_mail_inline_img_cache: dict[str, bool] = {}

# ==============================================================================
# STAŁE BEZPIECZEŃSTWA I OPTYMALIZACJI
# ==============================================================================
MAX_DOWNLOAD_IMAGE_BYTES = 25 * 1024 * 1024  # Maksymalny bezpieczny rozmiar pobieranego obrazu: 25 MB
DANGEROUS_ATTACHMENT_EXTENSIONS = {
	".exe", ".bat", ".cmd", ".com", ".pif", ".scr", ".vbs", ".vbe", ".vb",
	".js", ".jse", ".wsf", ".wsh", ".wsc", ".sct", ".shb", ".shs",
	".ps1", ".ps1xml", ".ps2", ".ps2xml", ".psc1", ".psc2", ".psd1", ".psm1",
	".msi", ".msp", ".mst", ".hta", ".cpl", ".msc", ".jar", ".reg",
	".lnk", ".url", ".iso", ".img", ".vhd", ".vhdx", ".chm", ".one",
	".docm", ".dotm", ".xlsm", ".xltm", ".xlam", ".pptm", ".potm", ".ppam", ".ppsm", ".sldm",
	".appref-ms", ".application", ".gadget", ".library-ms", ".settingcontent-ms",
	".mht", ".mhtml", ".svg", ".xbap", ".xll", ".dll", ".sys", ".drv", ".ocx",
}
ALLOWED_URL_SCHEMES = ("http", "https", "mailto", "tel", "geo")

# Prekompilowane wyrażenia regularne dla maksymalnej wydajności parsowania
_RE_HTML_TOKENS = re.compile(r'(<a\s+[^>]*href=["\'](?P<href>[^"\']+)["\'][^>]*>)|(<img\s+[^>]*>)', re.IGNORECASE)
_RE_IMG_SRC = re.compile(r'src=["\']([^"\']+)["\']', re.IGNORECASE)
_RE_IMG_ALT = re.compile(r'alt=["\']([^"\']*)["\']', re.IGNORECASE)
_RE_IMG_TITLE = re.compile(r'title=["\']([^"\']*)["\']', re.IGNORECASE)
_RE_CLEAN_FILENAME_INVALID = re.compile(r'[\\/*?:"<>|]')
_RE_CLEAN_FILENAME_WHITESPACE = re.compile(r'[\r\n\t]+')
_RE_CLEAN_FILENAME_DASHES = re.compile(r'[\s\-]+')


def cleanOldTempFiles(maxAgeSeconds: int = 86400):
	"""
	Usuwa stare pliki tymczasowe utworzone przez dodatek (starsze niż 24h),
	zapobiegając gromadzeniu się osieroconych plików na dysku użytkownika.
	"""
	try:
		tempDir = tempfile.gettempdir()
		now = time.time()
		# 1. Pliki obrazów nvda_oih_* w katalogu tymczasowym
		for fname in os.listdir(tempDir):
			if fname.startswith("nvda_oih_"):
				fpath = os.path.join(tempDir, fname)
				try:
					if os.path.isfile(fpath) and (now - os.path.getmtime(fpath) > maxAgeSeconds):
						os.remove(fpath)
				except Exception:
					pass
		# 2. Katalog załączników nvda_oih_attachments (pliki i podkatalogi)
		attDir = os.path.join(tempDir, "nvda_oih_attachments")
		if os.path.isdir(attDir):
			for fname in os.listdir(attDir):
				fpath = os.path.join(attDir, fname)
				try:
					if os.path.isdir(fpath) and (now - os.path.getmtime(fpath) > maxAgeSeconds):
						shutil.rmtree(fpath, ignore_errors=True)
					elif os.path.isfile(fpath) and (now - os.path.getmtime(fpath) > maxAgeSeconds):
						os.remove(fpath)
				except Exception:
					pass
		# 3. Pamięć podręczna przeskalowanych dźwięków nvda_oih_sound_cache
		if os.path.isdir(SOUNDS_CACHE_DIR):
			for fname in os.listdir(SOUNDS_CACHE_DIR):
				fpath = os.path.join(SOUNDS_CACHE_DIR, fname)
				try:
					if os.path.isfile(fpath) and (now - os.path.getmtime(fpath) > maxAgeSeconds):
						os.remove(fpath)
				except Exception:
					pass
	except Exception as ex:
		log.debug(f"outlookImageHelper: cleanOldTempFiles error: {ex}")



def getConfigDir() -> str:
	"""Zwraca ścieżkę do katalogu konfiguracji NVDA w sposób odporny na błędy."""
	try:
		p = getattr(getattr(globalVars, "appArgs", None), "configPath", None)
		if p and os.path.isdir(p):
			return p
	except Exception:
		pass
	try:
		if hasattr(config.conf, "filename") and config.conf.filename:
			return os.path.dirname(config.conf.filename)
	except Exception:
		pass
	appData = os.environ.get("APPDATA", "")
	if appData:
		nvdaDir = os.path.join(appData, "nvda")
		if os.path.isdir(nvdaDir):
			return nvdaDir
	return os.path.dirname(os.path.abspath(__file__))


PROMPTS_FILE = os.path.join(getConfigDir(), "outlookImageHelper_prompts.json")

confspec = {
	"language": "string(default='auto')",
	"autoAnnounce": "boolean(default=True)",
	"playSound": "boolean(default=True)",
	"soundTheme": "string(default='associative')",
	"soundTriggerMode": "string(default='open_only')",
	"soundVolume": "integer(default=80, min=0, max=100)",
	"navSounds": "boolean(default=True)",
	"defaultOcrLang": "string(default='pl-PL')",
	"aiProvider": "string(default='gemini')",
	"geminiApiKey": "string(default='')",
	"geminiModel": "string(default='gemini-3.8-flash')",
	"openaiApiKey": "string(default='')",
	"openaiModel": "string(default='gpt-4o-mini')",
	"aiPrompt": "string(default='')",
	"minDimension": "integer(default=15, min=1, max=500)",
	"aiNaming": "boolean(default=False)",
	"showAdvancedSettings": "boolean(default=False)",
	"contextMenuEnabled": "boolean(default=True)",
	"autoCheckUpdates": "boolean(default=True)",
	"updateSourceUrl": r"string(default='https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/update_info.json')",
	"lastUpdateCheckTime": "float(default=0.0)",
}



def _to_bool(val, default=False) -> bool:
	"""
	Bezpieczna konwersja dowolnej wartości (bool, str, int) na typ bool.
	Chroni przed błędem wxPython: CheckBox.SetValue(): argument 1 has unexpected type 'str'.
	"""
	if isinstance(val, bool):
		return val
	if isinstance(val, str):
		v = val.strip().lower()
		if v in ("true", "1", "yes", "on", "t", "prawda"):
			return True
		if v in ("false", "0", "no", "off", "f", "fałsz", "falsz"):
			return False
	if isinstance(val, (int, float)):
		return bool(val)
	return bool(default)


def _to_int(val, default=0, min_val=None, max_val=None) -> int:
	"""
	Bezpieczna konwersja wartości na int w opcjonalnym zakresie [min_val, max_val].
	"""
	try:
		iv = int(val)
		if min_val is not None:
			iv = max(min_val, iv)
		if max_val is not None:
			iv = min(max_val, iv)
		return iv
	except (ValueError, TypeError):
		return default


# WinAPI / COM constants
OBJID_NATIVEOM = -16  # 0xFFFFFFF0
IID_IDispatch_GUID = "{00020400-0000-0000-C000-000000000046}"

OCR_LANG_OPTIONS_PL = [
	("pl-PL", "Polski (pl-PL)"),
	("en-US", "Angielski (en-US)"),
	("de-DE", "Niemiecki (de-DE)"),
	("fr-FR", "Francuski (fr-FR)"),
	("es-ES", "Hiszpański (es-ES)"),
	("uk-UA", "Ukraiński (uk-UA)"),
	("it-IT", "Włoski (it-IT)"),
	("custom", "Inny / własny kod języka..."),
]

OCR_LANG_OPTIONS_EN = [
	("pl-PL", "Polish (pl-PL)"),
	("en-US", "English (en-US)"),
	("de-DE", "German (de-DE)"),
	("fr-FR", "French (fr-FR)"),
	("es-ES", "Spanish (es-ES)"),
	("uk-UA", "Ukrainian (uk-UA)"),
	("it-IT", "Italian (it-IT)"),
	("custom", "Other / custom language code..."),
]

OCR_LANG_OPTIONS = OCR_LANG_OPTIONS_PL

STRINGS = {
	"pl": {
		"addon_title": "Outlook - Pomocnik obrazów i załączników",
		"tag_text": "Tekst",
		"tag_text_empty": "Tekst: brak tekstu",
		"tag_ocr_ready": "OCR: gotowy",
		"tag_ai": "AI",
		"tag_ai_ready": "AI: gotowy",
		"tag_inverted": "[Negatyw]",
		"detail_colors_lbl": "Kolory: {status}",
		"detail_colors_inverted": "Inwersja kolorów (negatyw) — zapamiętano dla tego maila",
		"detail_colors_normal": "Standardowe barwy (oryginał)",
		"label_divider": "Linia podziału / separator",
		"label_social_icon": "Ikona {svc}",
		"label_phone_icon": "Ikona telefonu ({svc})",
		"label_email_icon": "Ikona poczty ({svc})",
		"label_unlabeled_img": "Obraz bez opisu",
		"dlg_elements_title": "Lista grafik i załączników w wiadomości",
		"cat_real_graphics": "&1. Grafiki i zrzuty ekranu w treści ({count})",
		"cat_attachments": "&2. Załączniki do wiadomości ({count})",
		"cat_divider_graphics": "&3. Linie podziału, separatory i ozdobniki ({count})",
		"lbl_elements_category": "&Kategoria elementów wiadomości:",
		"lbl_real_list": "&Wykryte grafiki (Enter - OCR, P - podgląd, I - inwersja, Alt+C - czynności):",
		"lbl_divider_list": "&Linie podziału (Enter - OCR, P - podgląd, I - inwersja, Alt+C - czynności):",
		"lbl_att_list": "&Załączniki do wiadomości (Enter - OCR/odczyt, Alt+C - czynności):",
		"lbl_details": "&Szczegóły wybranego elementu:",
		"btn_actions": "&Czynności (menu)...",
		"btn_preview": "&Podgląd (P)...",
		"btn_invert": "&Inwersja (I)",
		"btn_ocr": "&Rozpoznaj tekst (OCR)",
		"btn_ocr_pdf": "&Przepisz tekst z PDF (OCR / AI)",
		"btn_read_doc": "&Odczytaj treść dokumentu",
		"btn_ai_describe": "&Opisz przez AI...",
		"btn_ai_chat": "&Rozmowa z AI...",
		"btn_rename": "&Nazwij (F2)...",
		"btn_open_link": "&Otwórz link...",
		"btn_open_file": "&Otwórz plik załącznika...",
		"btn_save": "&Zapisz...",
		"btn_save_one_img": "&Zapisz wybraną grafikę...",
		"btn_save_one_att": "&Zapisz ten załącznik...",
		"btn_help": "Pomo&c (F1)...",
		"btn_close": "Zam&knij",
		"menu_ocr": "&Rozpoznaj tekst (OCR)\tAlt+R",
		"menu_preview": "&Pokaż podgląd grafiki (powiększenie)...\tP",
		"menu_invert_toggle": "&Przełącz inwersję kolorów (negatyw)\tI",
		"menu_invert_email": "Zastąp grafikę w treści &maila wersją w negatywie...",
		"menu_ai_quick": "&Opisz przez AI (szybki opis)\tAlt+O",
		"menu_ai_prompt": "Opisz przez AI z &wyborem prompta...",
		"menu_ai_chat": "&Rozmowa z AI (zadaj kolejne pytania o grafikę)...",
		"menu_rename": "&Zmień nazwę / nadaj własny opis grafice...\tF2",
		"menu_clear_name": "&Usuń własną nazwę (przywróć domyślną)",
		"menu_open_link": "&Otwórz link w przeglądarce ({domain})...\tCtrl+Enter",
		"menu_open_link_simple": "&Otwórz link w przeglądarce...\tCtrl+Enter",
		"menu_select_doc": "&Przejdź do grafiki w wiadomości",
		"menu_copy_img": "&Kopiuj grafikę do schowka",
		"menu_save_file": "&Zapisz wybraną grafikę jako plik...\tAlt+Z",
		"menu_save_all": "Zapisz &wszystkie grafiki do folderu...\tAlt+W",
		"menu_gemini_key": "Kl&ucz API Google Gemini...",
		"menu_help": "Pomo&c kontekstowa (skróty klawiszowe)...\tF1",
		"preview_title": "Podgląd grafiki: {name}",
		"preview_help_lbl": "Powiększenie: {zoom}% ({w}×{h} px){inv} | +/- zoom, 0 dopasuj, 1 100%, I inwersja, O w Windows, Esc zamknij, F1 pomoc",
		"preview_btn_in": "Powiększ (+)",
		"preview_btn_out": "Pomniejsz (-)",
		"preview_btn_fit": "Dopasuj do okna (0)",
		"preview_btn_100": "Skala 100% (1)",
		"preview_btn_invert": "Inwersja kolorów (I)",
		"preview_btn_invert_on": "Inwersja: WŁĄCZONA (I)",
		"preview_btn_invert_off": "Inwersja: wyłączona (I)",
		"preview_btn_open_os": "Otwórz w przeglądarce Windows (O)...",
		"preview_btn_help": "Pomoc (F1)...",
		"preview_btn_close": "Zamknij (Esc)",
		"preview_zoom_announced": "Powiększenie {zoom}%",
		"preview_invert_on": "Inwersja kolorów włączona (tryb negatywu).",
		"preview_invert_off": "Inwersja kolorów wyłączona (standardowe barwy).",
		"preview_opened_os": "Otwarto obraz w przeglądarce systemowej.",
		"preview_only_images": "Podgląd graficzny jest dostępny wyłącznie dla plików graficznych.",
		"err_extract_att": "Nie udało się wyodrębnić załącznika do podglądu.",
		"err_extract_img": "Nie udało się wyodrębnić grafiki do podglądu.",
		"msg_extracting_preview": "Pobieranie obrazu do powiększonego podglądu...",
		"list_invert_toggled_on": "Grafika {idx}: włączono inwersję kolorów (negatyw).",
		"list_invert_toggled_off": "Grafika {idx}: wyłączono inwersję kolorów (standardowe barwy).",
		"invert_msg_not_graphic": "Inwersja kolorów jest dostępna wyłącznie dla grafik.",
		"invert_msg_no_raw": "Nie można uzyskać dostępu do obiektu grafiki w wiadomości.",
		"invert_msg_extract_failed": "Nie udało się wyodrębnić obrazu z wiadomości.",
		"invert_msg_invert_failed": "Nie udało się wygenerować obrazu z odwróconymi kolorami.",
		"invert_msg_replaced_success": "Pomyślnie zastąpiono grafikę w wiadomości wersją z odwróconymi kolorami.",
		"invert_msg_shape_floating": "Wstawiono grafikę w negatywie w punkcie kotwiczenia kształtu.",
		"invert_msg_locked_warning": "Wiadomość jest chroniona przed edycją (tryb tylko do odczytu). Aby zastąpić obraz w samej wiadomości, włącz w Outlooku: Akcje -> Edytuj wiadomość. Możesz też bez przeszkód oglądać tę grafikę w negatywie w oknie podglądu (P lub I), co jest w 100% bezpieczne dla serwera pocztowego.",
		"script_quick_invert_doc": "Szybki podgląd bieżącej grafiki w wiadomości z natychmiastową inwersją kolorów (negatyw)",
		"chat_title": "Rozmowa z AI — {target}",
		"chat_history_lbl": "Przebieg rozmowy z AI (nawigacja: H nagłówki, K linki, 2 pytania H2, 3 odpowiedzi H3, 4 wypunktowania H4, B przyciski, E pytanie):",
		"chat_question_lbl": "Zadaj kolejne pytanie (wpisz treść i naciśnij Enter lub Ctrl+Enter, Shift+Enter dla nowej linii):",
		"chat_btn_send": "Wyślij pytanie (Enter)",
		"chat_btn_copy_last": "Kopiuj ostatnią odpowiedź",
		"chat_btn_copy_all": "Kopiuj całą rozmowę",
		"chat_btn_export": "Eksportuj do pliku...",
		"chat_chk_include_img": "Dołącz grafikę źródłową do schowka / eksportu",
		"chat_btn_help": "Pomoc (F1)",
		"chat_btn_close": "Zamknij (Esc)",
		"chat_btn_copy_ans": "Kopiuj odpowiedź {num} do schowka",
		"chat_btn_save_ans": "Zapisz odpowiedź {num} do pliku...",
		"chat_user_hdr": "Twoje pytanie {num}{time}",
		"chat_model_hdr": "Odpowiedź Gemini nr {num}{time}",
		"set_lang_lbl": "Język interfejsu dodatku / Add-on language:",
		"set_ocr_lbl": "&Domyślny język silnika OCR (Windows Media OCR):",
		"set_custom_ocr_lbl": "&Własny kod języka OCR (gdy wybrano 'Inny / własny kod języka...'):",
		"set_auto_announce": "&Powiadamiaj o wklejonych grafikach po otwarciu wiadomości",
		"set_play_sound": "&Odtwarzaj dźwięk, gdy wiadomość zawiera obrazy bez opisu / zrzuty ekranu",
		"set_context_menu": "&Otwieraj menu czynności klawiszem menu kontekstowego (Aplikacje / Shift+F10) wyłącznie gdy kursor stoi na grafice",
		"set_nav_sounds": "&Odtwarzaj subtelne dźwięki rozpoznawania typu podczas nawigacji (linki, ikony, separatory)",
		"set_sound_theme": "&Motyw dźwięku powiadomienia:",
		"set_test_sound": "&Odtwórz dźwięk",
		"set_sound_trigger_mode": "&Miejsce odtwarzania dźwięków o obecności obrazów:",
		"sound_mode_open_only": "Tylko po otwarciu wiadomości (domyślnie)",
		"sound_mode_list_only": "Tylko na liście wiadomości",
		"sound_mode_both": "Na liście wiadomości i po otwarciu",
		"sound_mode_off": "Wyłączone",
		"set_sound_volume": "&Głośność dźwięków alertów i nawigacji (%):",
		"set_test_volume": "&Przetestuj dźwięk i głośność",
		"set_sound_coord_detected": "Koordynacja dźwięków: Wykryto dodatek Outlook Speech Filter — dźwięki obecności grafik na liście są koordynowane, aby zapobiec duplikacji.",
		"set_sound_coord_standalone": "Koordynacja dźwięków: Outlook Speech Filter nie jest aktywny — dźwięki są w pełni kontrolowane przez Pomocnika obrazów.",
		"set_btn_diagnostics": "&Kontakt z twórcą i zgłaszanie uwag (GitHub)...",
		"set_ai_provider": "&Dostawca sztucznej inteligencji (AI):",
		"ai_prov_gemini": "Google Gemini (zalecany, darmowy)",
		"ai_prov_openai": "OpenAI / ChatGPT (GPT-4o, GPT-4o-mini)",
		"set_ai_naming": "&Używaj sztucznej inteligencji (AI) do generowania nazw plików (wymaga klucza, wysyła obrazy do chmury)",
		"set_gemini_key": "&Klucz API Google Gemini:",
		"set_gemini_model": "Model &Gemini (domyślnie: gemini-3.8-flash):",
		"set_openai_key": "Klucz API &OpenAI (ChatGPT):",
		"set_openai_model": "Model &OpenAI (domyślnie: gpt-4o-mini):",
		"set_gemini_prompt": "&Własna domyślna instrukcja dla AI (pozostaw puste dla standardowej):",
		"set_btn_get_gemini_key": "Instrukcja: &Klucz Google Gemini...",
		"set_btn_get_openai_key": "Instrukcja: Klucz &OpenAI (ChatGPT)...",
		"set_btn_test_gemini_key": "&Testuj klucz Gemini",
		"set_btn_test_openai_key": "Testuj klucz O&penAI",
		"set_btn_gestures": "&Skróty klawiszowe (Zdarzenia wejścia)...",
		"set_btn_guide": "&Instrukcja obsługi dodatku...",
		"set_btn_author_tools": "&Narzędzia twórcy dodatku...",
		"set_group_addon_info": "Informacje o dodatku",
		"set_info_version": "Wersja dodatku:",
		"set_info_build_date": "Data kompilacji:",
		"set_info_author": "Twórca dodatku:",
		"set_info_compatibility": "Zgodność:",
		"msg_outlook_only": "Skrót działa wyłącznie w programie Microsoft Outlook.",
		"msg_reading_pane_off": "Aby zidentyfikować grafiki i obrazy w tej wiadomości, otwórz ją klawiszem Enter.",
		"msg_reading_pane_off_subj": "Aby zidentyfikować grafiki i obrazy w wiadomości '{subj}', otwórz ją klawiszem Enter.",
		"msg_reading_pane_off_att": "Wiadomość zawiera załączniki ({count}). Aby zidentyfikować wklejone grafiki, otwórz wiadomość klawiszem Enter.",
		"msg_reading_pane_off_att_subj": "Wiadomość '{subj}' zawiera załączniki ({count}). Aby zidentyfikować wklejone grafiki, otwórz wiadomość klawiszem Enter.",
		"msg_reading_pane_off_ocr": "Rozpoznawanie tekstu (OCR) wklejonych grafik wymaga otwarcia wiadomości klawiszem Enter.",
		"msg_reading_pane_off_ocr_subj": "Rozpoznawanie tekstu (OCR) wklejonych grafik w wiadomości '{subj}' wymaga jej otwarcia klawiszem Enter.",
		"msg_reading_pane_off_ai": "Opisywanie wklejonych grafik przez AI wymaga otwarcia wiadomości klawiszem Enter.",
		"msg_reading_pane_off_ai_subj": "Opisywanie wklejonych grafik przez AI w wiadomości '{subj}' wymaga jej otwarcia klawiszem Enter.",
		"msg_reading_pane_off_invert": "Podgląd w negatywie wklejonych grafik wymaga otwarcia wiadomości klawiszem Enter.",
		"msg_reading_pane_off_invert_subj": "Podgląd w negatywie wklejonych grafik w wiadomości '{subj}' wymaga jej otwarcia klawiszem Enter.",
		"msg_multiple_selected": "Zaznaczono {count} wiadomości. Okno listy grafik i załączników dotyczy jednej wiadomości. Zaznacz pojedynczą wiadomość lub otwórz ją klawiszem Enter.",
		"msg_multiple_selected_att": "Zaznaczono {count} wiadomości (w tym {attCount} z załącznikami). Aby przejrzeć grafiki i załączniki, zaznacz jedną wiadomość lub otwórz ją klawiszem Enter.",
		"msg_multiple_selected_action": "Zaznaczono {count} wiadomości. Ta czynność dotyczy pojedynczej wiadomości. Zaznacz jedną wiadomość lub otwórz ją klawiszem Enter.",
		"msg_no_elements": "W bieżącej wiadomości nie odnaleziono grafik ani załączników.",
		"msg_no_elements_subj": "Brak załączników ani grafik w wiadomości '{subj}'.",
		"msg_no_ocr_subj": "Brak grafik ani załączników do OCR w wiadomości '{subj}'.",
		"msg_no_ai_subj": "Brak grafik ani załączników do opisu AI w wiadomości '{subj}'.",
		"msg_no_preview_subj": "Brak grafik do podglądu w wiadomości '{subj}'.",
		"msg_one_screenshot": "Wiadomość zawiera wklejony zrzut ekranu. Naciśnij NVDA+Windows+G lub Alt+Shift+O, aby rozpoznać tekst (OCR).",
		"msg_items_found": "Wiadomość zawiera {total} grafik (w tym {unlabeled} bez opisu). Naciśnij Ctrl+NVDA+A dla listy grafik.",
		"msg_items_all_labeled": "Wiadomość zawiera {total} wklejonych grafik. Ctrl+NVDA+A dla listy.",
		"script_list_doc": "Otwiera okno dialogowe z listą wszystkich grafik oraz załączników w bieżącej wiadomości programu Outlook",
		"script_ocr_doc": "Szybkie rozpoznanie tekstu (OCR) bieżącej grafiki lub zrzutu ekranu w wiadomości",
		"script_ai_doc": "Szybki opis bieżącej grafiki lub załącznika w wiadomości przez sztuczną inteligencję (AI)",
		"ocr_no_text_msg": "Brak rozpoznanego tekstu ({target}).",
		"ocr_no_text_title": "Brak tekstu — {target}",
		"ocr_no_text_body": "Brak rozpoznanego tekstu.\n\nGrafika {index} ({width}×{height} px).\n\nWskazówka: Jeśli grafika przedstawia zdjęcie, schemat, wykres, tabelę lub rysunek, skorzystaj z opcji 'Opisz przez AI' (Alt+O) lub 'Rozmowa z AI', która dokładnie zanalizuje jej treść wizualną.",
		"ocr_no_text_att_body": "Brak rozpoznanego tekstu.\n\nZałącznik: {filename} ({size}).\nSilnik OCR nie wykrył tekstu w tym pliku.",
		"set_btn_get_api_key": "&Jak uzyskać bezpłatny klucz API (Kreator i instrukcja)...",
		"set_btn_test_api_key": "&Testuj klucz API",
		"wizard_title": "Kreator pozyskiwania bezpłatnego klucza Google Gemini API",
		"wizard_instructions_lbl": "&Instrukcja pozyskania bezpłatnego klucza (czytaj kursorami góra/dół):",
		"wizard_info": "Sztuczna inteligencja Google Gemini jest całkowicie bezpłatna dla użytkowników indywidualnych (konta Google / Gmail) z darmowym limitem do 15 zapytań na minutę i 1500 na dobę.\n\nInstrukcja uzyskania bezpłatnego klucza krok po kroku:\n1. Kliknij poniższy przycisk '1. Otwórz stronę Google AI Studio w przeglądarce'.\n2. Zaloguj się na swoje konto Google (Gmail).\n3. Na otwartej stronie kliknij przycisk 'Create API key' (Utwórz klucz API).\n4. Wybierz opcję 'Create API key in new project' i zatwierdź.\n5. W okienku z gotowym kluczem kliknij przycisk 'Copy' (Kopiuj).\n6. Wróć do tego okna, kliknij '2. Wklej klucz ze schowka', a następnie '3. Testuj połączenie z Google Gemini'.\n7. Na koniec kliknij 'Zapisz klucz i zakończ'.",
		"wizard_btn_open": "&1. Otwórz stronę Google AI Studio w przeglądarce...",
		"wizard_lbl_key": "&Klucz API Google Gemini:",
		"wizard_btn_paste": "&2. Wklej klucz ze schowka",
		"wizard_btn_test": "&3. Testuj połączenie z Google Gemini",
		"wizard_btn_save": "&Zapisz klucz i zakończ",
		"wizard_key_empty": "Pole klucza API jest puste. Skopiuj klucz z Google AI Studio lub wpisz go ręcznie.",
		"wizard_testing": "Trwa testowanie połączenia z Google Gemini...",
		"wizard_test_success": "Sukces! Połączenie z Google Gemini działa prawidłowo (model: {model}). Klucz jest aktywny.",
		"wizard_test_err_invalid_key": "Nieprawidłowy klucz API (kod 400). Upewnij się, że skopiowano cały klucz (zwykle zaczyna się od 'AIzaSy').",
		"wizard_test_err_forbidden": "Odmowa dostępu (kod 403). Upewnij się, że usługa Gemini API jest dostępna dla Twojego konta Google.",
		"wizard_test_warn_quota": "Klucz jest poprawny, ale chwilowo wyczerpano limit zapytań (kod 429).",
		"wizard_test_err_http": "Błąd serwera Google Gemini: HTTP {code}.",
		"wizard_test_err_network": "Błąd połączenia z siecią: {err}.",
		"wizard_test_err_failed": "Nie udało się nawiązać połączenia: {err}.",
		"wizard_clipboard_empty": "Schowek nie zawiera tekstu.",
		"wizard_pasted_msg": "Wklejono klucz ze schowka ({chars} znaków).",
		"wizard_saved_msg": "Klucz API Google Gemini został pomyślnie zapisany.",
		"set_chk_show_advanced": "&Pokaż ustawienia zaawansowane (Sztuczna inteligencja AI i diagnostyka)",
		"set_btn_translate_ai": "&Tłumacz / Wygeneruj nowy język przez AI...",
		"set_btn_export_lang": "&Eksportuj szablon języka (JSON)...",
		"set_btn_import_lang": "&Importuj plik języka (JSON)...",
		"set_btn_open_lang_dir": "&Otwórz folder języków...",
		"dlg_trans_title": "Generator wersji językowej przez Google Gemini",
		"dlg_trans_target_name": "&Nazwa języka (np. Deutsch, Français, Español):",
		"dlg_trans_target_code": "&Kod języka ISO (np. de, fr, es, uk, cs):",
		"dlg_trans_include_doc": "&Przetłumacz również pełną dokumentację HTML (doc.html)",
		"dlg_trans_scope_box": "Zakres i opcje tłumaczenia:",
		"dlg_trans_chk_incremental": "Tłumaczenie przyrostowe (różnicowe — tłumacz tylko nowe lub brakujące elementy)",
		"dlg_trans_chk_manifest": "1. Opis i metadane dodatku (manifest.ini)",
		"dlg_trans_chk_strings": "2. Napisy interfejsu (menu, przyciski, komunikaty — 176 fraz)",
		"dlg_trans_chk_help": "3. Pomoc kontekstowa F1 (okno listy, czat AI, podgląd, ustawienia)",
		"dlg_trans_chk_doc": "4. Podręcznik użytkownika HTML (podzielony na sekcje, ok. 30-40 sekund)",
		"dlg_trans_chk_step_by_step": "Tryb sekwencyjny (zatrzymuj po każdym kroku i czekaj na naciśnięcie 'Dalej')",
		"dlg_trans_btn_next": "Dalej: Rozpocznij {step}...",
		"dlg_trans_log_lbl": "Dziennik postępu i komunikaty (możesz czytać strzałkami lub skopiować):",
		"dlg_trans_btn_copy_log": "Kopiuj dziennik do schowka",
		"dlg_trans_btn_start": "&Rozpocznij tłumaczenie AI",
		"dlg_trans_in_progress": "Trwa tłumaczenie interfejsu i dokumentacji na język {name} przez Google Gemini. Proszę czekać...",
		"dlg_trans_success": "Pomyślnie wygenerowano i dodano język '{name}' ({code})! Język jest teraz aktywny na liście.",
		"dlg_trans_err": "Błąd podczas generowania wersji językowej: {err}",
		"dlg_export_title": "Zapisz szablon językowy JSON",
		"dlg_export_success": "Pomyślnie wyeksportowano szablon języka do pliku: {file}",
		"dlg_import_title": "Wybierz plik języka JSON do zaimportowania",
		"dlg_import_success": "Pomyślnie zaimportowano język '{name}' ({code})! Jest już dostępny na liście.",
		"dlg_import_err_invalid": "Wybrany plik nie jest poprawnym plikiem językowym dodatku: {err}",
	},
	"en": {
		"addon_title": "Outlook - Image & Attachment Helper",
		"tag_text": "Text",
		"tag_text_empty": "Text: no text",
		"tag_ocr_ready": "OCR: ready",
		"tag_ai": "AI",
		"tag_ai_ready": "AI: ready",
		"tag_inverted": "[Inverted]",
		"detail_colors_lbl": "Colors: {status}",
		"detail_colors_inverted": "Inverted colors (negative) — remembered for this email",
		"detail_colors_normal": "Normal colors (original)",
		"label_divider": "Divider / separator line",
		"label_social_icon": "{svc} icon",
		"label_phone_icon": "Phone icon ({svc})",
		"label_email_icon": "Email icon ({svc})",
		"label_unlabeled_img": "Unlabeled image",
		"dlg_elements_title": "Images and Attachments in Message",
		"cat_real_graphics": "&1. Inline images and screenshots ({count})",
		"cat_attachments": "&2. Message attachments ({count})",
		"cat_divider_graphics": "&3. Dividers, separators and ornaments ({count})",
		"lbl_elements_category": "&Message element category:",
		"lbl_real_list": "&Detected images (Enter - OCR, P - preview, I - invert, Alt+C - actions):",
		"lbl_divider_list": "&Dividers (Enter - OCR, P - preview, I - invert, Alt+C - actions):",
		"lbl_att_list": "&Message attachments (Enter - OCR/read, Alt+C - actions):",
		"lbl_details": "&Details of selected item:",
		"btn_actions": "&Actions (menu)...",
		"btn_preview": "&Preview (P)...",
		"btn_invert": "&Invert (I)",
		"btn_ocr": "&Recognize text (OCR)",
		"btn_ocr_pdf": "&Transcribe PDF text (OCR / AI)",
		"btn_read_doc": "&Read document text",
		"btn_ai_describe": "&Describe with AI...",
		"btn_ai_chat": "&Chat with AI...",
		"btn_rename": "&Name (F2)...",
		"btn_open_link": "&Open link...",
		"btn_open_file": "&Open attachment file...",
		"btn_save": "&Save...",
		"btn_save_one_img": "&Save selected image...",
		"btn_save_one_att": "&Save this attachment...",
		"btn_help": "&Help (F1)...",
		"btn_close": "&Close",
		"menu_ocr": "&Recognize text (OCR)\tAlt+R",
		"menu_preview": "&Show image preview (magnified)...\tP",
		"menu_invert_toggle": "&Toggle color inversion (negative)\tI",
		"menu_invert_email": "Replace image in &email body with negative version...",
		"menu_ai_quick": "&Describe with AI (quick summary)\tAlt+O",
		"menu_ai_prompt": "Describe with AI using &prompt preset...",
		"menu_ai_chat": "&Chat with AI (ask follow-up questions)...",
		"menu_rename": "&Rename / set custom description...\tF2",
		"menu_clear_name": "&Remove custom name (restore default)",
		"menu_open_link": "&Open link in browser ({domain})...\tCtrl+Enter",
		"menu_open_link_simple": "&Open link in browser...\tCtrl+Enter",
		"menu_select_doc": "&Go to image in message",
		"menu_copy_img": "&Copy image to clipboard",
		"menu_save_file": "&Save selected image to file...\tAlt+Z",
		"menu_save_all": "Save &all images to folder...\tAlt+W",
		"menu_gemini_key": "Google Gemini API &Key...",
		"menu_help": "Contextual &help (keyboard shortcuts)...\tF1",
		"preview_title": "Image Preview: {name}",
		"preview_help_lbl": "Zoom: {zoom}% ({w}×{h} px){inv} | +/- zoom, 0 fit, 1 100%, I invert, O in Windows, Esc close, F1 help",
		"preview_btn_in": "Zoom In (+)",
		"preview_btn_out": "Zoom Out (-)",
		"preview_btn_fit": "Fit to Window (0)",
		"preview_btn_100": "100% Scale (1)",
		"preview_btn_invert": "Invert colors (I)",
		"preview_btn_invert_on": "Inversion: ON (I)",
		"preview_btn_invert_off": "Inversion: off (I)",
		"preview_btn_open_os": "Open in Windows Viewer (O)...",
		"preview_btn_help": "Help (F1)...",
		"preview_btn_close": "Close (Esc)",
		"preview_zoom_announced": "Zoom {zoom}%",
		"preview_invert_on": "Color inversion enabled (negative mode).",
		"preview_invert_off": "Color inversion disabled (normal colors).",
		"preview_opened_os": "Opened image in default Windows viewer.",
		"preview_only_images": "Image preview is only available for image files.",
		"err_extract_att": "Failed to extract attachment for preview.",
		"err_extract_img": "Failed to extract image for preview.",
		"msg_extracting_preview": "Extracting image for magnified preview...",
		"list_invert_toggled_on": "Graphic {idx}: color inversion enabled (negative).",
		"list_invert_toggled_off": "Graphic {idx}: color inversion disabled (normal colors).",
		"invert_msg_not_graphic": "Color inversion is only available for graphics.",
		"invert_msg_no_raw": "Unable to access the graphic object in message.",
		"invert_msg_extract_failed": "Failed to extract image from message.",
		"invert_msg_invert_failed": "Failed to generate image with inverted colors.",
		"invert_msg_replaced_success": "Successfully replaced graphic in email body with inverted colors version.",
		"invert_msg_shape_floating": "Inserted negative graphic at the shape anchor.",
		"invert_msg_locked_warning": "Message is write-protected (read-only mode). To replace the image in the email itself, enable in Outlook: Actions -> Edit Message. You can also view this graphic in negative in the preview window (P or I), which is 100% safe for mail servers.",
		"script_quick_invert_doc": "Quick preview of current graphic in message with instant color inversion (negative)",
		"chat_title": "Chat with AI — {target}",
		"chat_history_lbl": "Chat history with AI (navigation: H headings, K links, 2 questions H2, 3 answers H3, 4 bullets H4, B buttons, E question):",
		"chat_question_lbl": "Ask a follow-up question (type text and press Enter or Ctrl+Enter, Shift+Enter for new line):",
		"chat_btn_send": "Send question (Enter)",
		"chat_btn_copy_last": "Copy last answer",
		"chat_btn_copy_all": "Copy whole conversation",
		"chat_btn_export": "Export to file...",
		"chat_chk_include_img": "Include source image in clipboard / export",
		"chat_btn_help": "Help (F1)",
		"chat_btn_close": "Close (Esc)",
		"chat_btn_copy_ans": "Copy answer {num} to clipboard",
		"chat_btn_save_ans": "Save answer {num} to file...",
		"chat_user_hdr": "Your Question {num}{time}",
		"chat_model_hdr": "Gemini Answer #{num}{time}",
		"set_lang_lbl": "Add-on interface language / Język interfejsu:",
		"set_ocr_lbl": "&Default OCR engine language (Windows Media OCR):",
		"set_custom_ocr_lbl": "&Custom OCR language code (when 'Other / custom code...' is selected):",
		"set_auto_announce": "&Announce embedded images when opening a message",
		"set_play_sound": "&Play sound when message contains unlabeled images / screenshots",
		"set_context_menu": "&Open actions menu with context menu key (Apps / Shift+F10) only when cursor is directly on a graphic",
		"set_nav_sounds": "&Play subtle recognition sounds during navigation (links, icons, dividers)",
		"set_sound_theme": "&Alert sound theme:",
		"set_test_sound": "&Play sound",
		"set_sound_trigger_mode": "&Trigger location for image alert sounds:",
		"sound_mode_open_only": "Only when opening a message (default)",
		"sound_mode_list_only": "Only on the message list",
		"sound_mode_both": "Both on the list and when opening",
		"sound_mode_off": "Disabled",
		"set_sound_volume": "&Alert and navigation sounds volume (%):",
		"set_test_volume": "&Test sound and volume",
		"set_sound_coord_detected": "Sound coordination: Outlook Speech Filter detected — list image sounds are coordinated to prevent duplicates.",
		"set_sound_coord_standalone": "Sound coordination: Outlook Speech Filter is not active — sounds are handled independently.",
		"set_btn_diagnostics": "&Contact author & feedback (GitHub)...",
		"set_ai_provider": "&AI Provider:",
		"ai_prov_gemini": "Google Gemini (recommended, free tier)",
		"ai_prov_openai": "OpenAI / ChatGPT (GPT-4o, GPT-4o-mini)",
		"set_ai_naming": "&Use AI to generate filenames (requires API key, sends images to cloud)",
		"set_gemini_key": "Google Gemini &API Key:",
		"set_gemini_model": "Gemini &Model (default: gemini-3.8-flash):",
		"set_openai_key": "OpenAI (ChatGPT) &API Key:",
		"set_openai_model": "OpenAI &Model (default: gpt-4o-mini):",
		"set_gemini_prompt": "&Custom default prompt for AI (leave empty for standard):",
		"set_btn_get_gemini_key": "&Guide: Google Gemini Key...",
		"set_btn_get_openai_key": "G&uide: OpenAI (ChatGPT) Key...",
		"set_btn_test_gemini_key": "&Test Gemini Key",
		"set_btn_test_openai_key": "Test &OpenAI Key",
		"set_btn_gestures": "&Input Gestures (keyboard shortcuts)...",
		"set_btn_guide": "&User Guide...",
		"set_btn_author_tools": "&Creator Tools...",
		"set_group_addon_info": "Add-on Information",
		"set_info_version": "Add-on version:",
		"set_info_build_date": "Build date:",
		"set_info_author": "Add-on author:",
		"set_info_compatibility": "Compatibility:",
		"msg_outlook_only": "Shortcut works only in Microsoft Outlook.",
		"msg_reading_pane_off": "To identify images and graphics in this message, open it with Enter.",
		"msg_reading_pane_off_subj": "To identify images and graphics in message '{subj}', open it with Enter.",
		"msg_reading_pane_off_att": "Message contains attachments ({count}). To identify inline graphics, open the message with Enter.",
		"msg_reading_pane_off_att_subj": "Message '{subj}' contains attachments ({count}). To identify inline graphics, open it with Enter.",
		"msg_reading_pane_off_ocr": "Text recognition (OCR) of inline graphics requires opening the message with Enter.",
		"msg_reading_pane_off_ocr_subj": "Text recognition (OCR) of inline graphics in message '{subj}' requires opening it with Enter.",
		"msg_reading_pane_off_ai": "Describing inline graphics with AI requires opening the message with Enter.",
		"msg_reading_pane_off_ai_subj": "Describing inline graphics with AI in message '{subj}' requires opening it with Enter.",
		"msg_reading_pane_off_invert": "Inverted preview of inline graphics requires opening the message with Enter.",
		"msg_reading_pane_off_invert_subj": "Inverted preview of inline graphics in message '{subj}' requires opening it with Enter.",
		"msg_multiple_selected": "{count} messages are selected. The images and attachments dialog applies to a single message. Please select a single message or open it with Enter.",
		"msg_multiple_selected_att": "{count} messages are selected (including {attCount} with attachments). To browse images and attachments, select a single message or open it with Enter.",
		"msg_multiple_selected_action": "{count} messages are selected. This action applies to a single message. Please select a single message or open it with Enter.",
		"msg_no_elements": "No images or attachments found in current message.",
		"msg_no_elements_subj": "No attachments or images in message '{subj}'.",
		"msg_no_ocr_subj": "No images or attachments for OCR in message '{subj}'.",
		"msg_no_ai_subj": "No images or attachments for AI description in message '{subj}'.",
		"msg_no_preview_subj": "No images to preview in message '{subj}'.",
		"msg_one_screenshot": "Message contains a pasted screenshot. Press NVDA+Windows+G or Alt+Shift+O to recognize text (OCR).",
		"msg_items_found": "Message contains {total} images (including {unlabeled} without description). Press Ctrl+NVDA+A for images list.",
		"msg_items_all_labeled": "Message contains {total} pasted images. Press Ctrl+NVDA+A for list.",
		"script_list_doc": "Opens dialog with list of all images and attachments in current Outlook message",
		"script_ocr_doc": "Quick OCR recognition of image or screenshot under cursor or in message",
		"script_ai_doc": "Quick AI description of current image or attachment in message",
		"ocr_no_text_msg": "No text recognized ({target}).",
		"ocr_no_text_title": "No text — {target}",
		"ocr_no_text_body": "No text recognized.\n\nImage {index} ({width}×{height} px).\n\nTip: If the image is a photo, diagram, chart, table, or drawing, use 'Describe with AI' (Alt+O) or 'Chat with AI' to analyze its visual content.",
		"ocr_no_text_att_body": "No text recognized.\n\nAttachment: {filename} ({size}).\nThe OCR engine did not detect any text in this file.",
		"set_btn_get_api_key": "&How to get a free API Key (Wizard & Guide)...",
		"set_btn_test_api_key": "&Test API Key",
		"wizard_title": "Google Gemini Free API Key Setup Wizard",
		"wizard_instructions_lbl": "&Instructions for getting a free API key (read with Up/Down arrows):",
		"wizard_info": "Google Gemini AI is completely free for personal Google / Gmail accounts with generous limits (up to 15 requests per minute, 1500 per day).\n\nStep-by-step instructions:\n1. Click '1. Open Google AI Studio in browser'.\n2. Sign in with your Google account (Gmail).\n3. On the page, find and click 'Create API key'.\n4. Select 'Create API key in new project' and confirm.\n5. When the key dialog appears, click 'Copy'.\n6. Switch back here, click '2. Paste key from clipboard', then '3. Test connection with Google Gemini'.\n7. Finally, click 'Save key and finish'.",
		"wizard_btn_open": "&1. Open Google AI Studio in browser...",
		"wizard_lbl_key": "&Google Gemini API Key:",
		"wizard_btn_paste": "&2. Paste key from clipboard",
		"wizard_btn_test": "&3. Test connection with Google Gemini",
		"wizard_btn_save": "&Save key and finish",
		"wizard_key_empty": "API Key field is empty. Please copy the key from Google AI Studio or type it.",
		"wizard_testing": "Testing connection with Google Gemini...",
		"wizard_test_success": "Success! Google Gemini connection is working properly (model: {model}). Key is active.",
		"wizard_test_err_invalid_key": "Invalid API key (HTTP 400). Please ensure the entire key was copied (starts with 'AIzaSy').",
		"wizard_test_err_forbidden": "Access denied (HTTP 403). Ensure Gemini API is enabled for your Google account.",
		"wizard_test_warn_quota": "The key is valid, but rate limits were temporarily exceeded (HTTP 429).",
		"wizard_test_err_http": "Google Gemini server error: HTTP {code}.",
		"wizard_test_err_network": "Network connection error: {err}.",
		"wizard_test_err_failed": "Failed to connect: {err}.",
		"wizard_clipboard_empty": "Clipboard does not contain text.",
		"wizard_pasted_msg": "Pasted key from clipboard ({chars} characters).",
		"wizard_saved_msg": "Google Gemini API key saved successfully.",
		"set_chk_show_advanced": "&Show advanced settings (Artificial Intelligence AI & Diagnostics)",
		"set_btn_translate_ai": "&Translate / Generate new language via AI...",
		"set_btn_export_lang": "&Export language template (JSON)...",
		"set_btn_import_lang": "&Import language file (JSON)...",
		"set_btn_open_lang_dir": "&Open languages folder...",
		"dlg_trans_title": "Language Version Generator via Google Gemini",
		"dlg_trans_target_name": "&Language name (e.g. Deutsch, Français, Español):",
		"dlg_trans_target_code": "ISO Language &code (e.g. de, fr, es, uk, cs):",
		"dlg_trans_include_doc": "&Also translate full HTML documentation (doc.html)",
		"dlg_trans_scope_box": "Translation scope and options:",
		"dlg_trans_chk_incremental": "Incremental translation (differential — translate only new or missing items)",
		"dlg_trans_chk_manifest": "1. Add-on description & metadata (manifest.ini)",
		"dlg_trans_chk_strings": "2. User interface strings (menus, buttons, messages — 176 phrases)",
		"dlg_trans_chk_help": "3. Contextual help F1 (list, AI chat, preview, settings)",
		"dlg_trans_chk_doc": "4. User manual HTML (divided into sections, ~30-40 sec)",
		"dlg_trans_chk_step_by_step": "Sequential mode (pause after each step and wait for 'Next' button)",
		"dlg_trans_btn_next": "Next: Start {step}...",
		"dlg_trans_log_lbl": "Progress log & messages (read with arrow keys or copy):",
		"dlg_trans_btn_copy_log": "Copy log to clipboard",
		"dlg_trans_btn_start": "&Start AI Translation",
		"dlg_trans_in_progress": "Translating interface and documentation into {name} via Google Gemini. Please wait...",
		"dlg_trans_success": "Successfully generated and added language '{name}' ({code})! The language is now available in the list.",
		"dlg_trans_err": "Error generating language version: {err}",
		"dlg_export_title": "Save JSON Language Template",
		"dlg_export_success": "Successfully exported language template to: {file}",
		"dlg_import_title": "Select JSON Language File to Import",
		"dlg_import_success": "Successfully imported language '{name}' ({code})! It is now available in the list.",
		"dlg_import_err_invalid": "The selected file is not a valid add-on language file: {err}",
	},
}


CUSTOM_LANGUAGES_DIR = os.path.join(getConfigDir(), "outlookImageHelper_languages")
CUSTOM_LANGUAGES: dict[str, dict] = {}


def loadCustomLanguages():
	"""Ładuje wszystkie pliki języków JSON z katalogu CUSTOM_LANGUAGES_DIR."""
	global CUSTOM_LANGUAGES
	CUSTOM_LANGUAGES = {}
	try:
		if not os.path.isdir(CUSTOM_LANGUAGES_DIR):
			os.makedirs(CUSTOM_LANGUAGES_DIR, exist_ok=True)
			return
		for fname in os.listdir(CUSTOM_LANGUAGES_DIR):
			if fname.lower().endswith(".json"):
				fpath = os.path.join(CUSTOM_LANGUAGES_DIR, fname)
				try:
					with open(fpath, "r", encoding="utf-8") as f:
						data = json.load(f)
					code = str(data.get("langCode") or os.path.splitext(fname)[0]).strip().lower()
					if code and isinstance(data, dict):
						CUSTOM_LANGUAGES[code] = data
				except Exception as ex:
					log.warning(f"outlookImageHelper: Błąd ładowania języka {fname}: {ex}")
	except Exception as ex:
		log.warning(f"outlookImageHelper: Błąd skanowania CUSTOM_LANGUAGES_DIR: {ex}")


loadCustomLanguages()


def getAvailableLanguages() -> list[tuple[str, str, str]]:
	"""
	Zwraca listę dostępnych języków jako krotki: (kod, nazwa_pl, nazwa_en).
	Zawiera języki wbudowane oraz wszystkie niestandardowe języki z CUSTOM_LANGUAGES.
	"""
	langs = [
		("auto", "Automatyczny (zgodny z systemem i NVDA) / Auto", "Automatic (matches system and NVDA) / Auto"),
		("pl", "Polski (Polish)", "Polski (Polish)"),
		("en", "English", "English"),
	]
	seen = {"auto", "pl", "en"}
	for code, data in sorted(CUSTOM_LANGUAGES.items()):
		if code not in seen:
			name = data.get("langName", code)
			langs.append((code, name, name))
			seen.add(code)
	return langs


_activeSettingsPanelRef = None


def getEffectiveLanguage() -> str:
	"""
	Zwraca efektywny kod języka ('pl', 'en' lub kod języka niestandardowego):
	1. Jeśli panel ustawień jest otwarty w GUI, respektuje bieżący wybór użytkownika w locie.
	2. Jeśli zapisano język w konfiguracji -> zwraca ten język.
	3. Jeśli wybrano 'auto' (domyślny) -> precyzyjnie sprawdza język systemu Windows (GetUserDefaultUILanguage),
	   język główny NVDA (config.conf["general"]["language"]) oraz languageHandler.
	   Na polskim systemie/NVDA zawsze bezwzględnie zwraca 'pl'.
	"""
	# 1. Panel ustawień otwarty w GUI
	try:
		if _activeSettingsPanelRef is not None:
			panel = _activeSettingsPanelRef()
			if panel and hasattr(panel, "languageChoice") and hasattr(panel, "langKeys"):
				idx = panel.languageChoice.GetSelection()
				if 0 <= idx < len(panel.langKeys):
					chosen = panel.langKeys[idx]
					if chosen in ("pl", "en") or chosen in CUSTOM_LANGUAGES:
						return chosen
	except Exception:
		pass

	# 2. Zapisana konfiguracja dodatku
	try:
		sec = config.conf[CONFIG_SECTION]
		cfgLang = str(sec.get("language", "auto")).strip().lower()
		if cfgLang in ("pl", "en") or cfgLang in CUSTOM_LANGUAGES:
			return cfgLang
	except Exception:
		pass

	# 3. Tryb 'auto' / domyślny - wykryj język systemowy Windows oraz NVDA:
	# Sprawdzamy interfejs użytkownika Windows (GetUserDefaultUILanguage)
	try:
		# Maska 0x3FF wydobywa identyfikator języka podstawowego (PRIMARYLANGID)
		# 0x15 = 21 = Polish (Polski)
		# 0x09 = 9 = English (Angielski)
		langId = ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF
		if langId == 0x15:
			return "pl"
		elif langId == 0x09:
			return "en"
	except Exception:
		pass

	# Sprawdzamy język główny NVDA w konfiguracji ogólnej
	try:
		nvdaGenLang = str(config.conf.get("general", {}).get("language", "")).strip().lower()
		if nvdaGenLang.startswith("pl"):
			return "pl"
		elif nvdaGenLang.startswith("en"):
			return "en"
	except Exception:
		pass

	# Sprawdzamy moduł languageHandler NVDA
	try:
		import languageHandler
		l = str(languageHandler.getLanguage() or "").strip().lower()
		if l.startswith("pl"):
			return "pl"
		elif l.startswith("en"):
			return "en"
		if hasattr(languageHandler, "getWindowsLanguage"):
			wl = str(languageHandler.getWindowsLanguage() or "").strip().lower()
			if wl.startswith("pl"):
				return "pl"
			elif wl.startswith("en"):
				return "en"
	except Exception:
		pass

	# Sprawdzamy locale systemowe Pythona
	try:
		import locale
		loc = str(locale.getdefaultlocale()[0] or "").strip().lower()
		if loc.startswith("pl"):
			return "pl"
		elif loc.startswith("en"):
			return "en"
	except Exception:
		pass

	return "pl"


def _T(key: str, **kwargs) -> str:
	lang = getEffectiveLanguage()
	s = None
	if lang in CUSTOM_LANGUAGES:
		s = CUSTOM_LANGUAGES[lang].get("strings", {}).get(key)
	if not s:
		s = STRINGS.get(lang, {}).get(key) or STRINGS.get("pl", {}).get(key) or key
	if kwargs:
		try:
			return s.format(**kwargs)
		except Exception:
			return s
	return s


def _cleanSnippet(text: str, maxLen: int = 38) -> str:
	if not text:
		return ""
	cleaned = " ".join(text.strip().split())
	if len(cleaned) > maxLen:
		return cleaned[:maxLen - 3] + "..."
	return cleaned


def openUserGuide():
	"""Otwiera instrukcję obsługi dodatku w domyślnej przeglądarce internetowej."""
	lang = getEffectiveLanguage()
	candidates = []
	if lang not in ("pl", "auto"):
		candidates.append(os.path.join(CUSTOM_LANGUAGES_DIR, f"doc_{lang}.html"))
		candidates.append(os.path.join(ADDON_DIR, f"doc_{lang}.html"))
	if lang == "en":
		candidates.append(os.path.join(ADDON_DIR, "doc_en.html"))
	else:
		candidates.append(os.path.join(ADDON_DIR, "doc_pl.html"))
	candidates.append(os.path.join(ADDON_DIR, "doc.html"))

	for docPath in candidates:
		if os.path.isfile(docPath):
			try:
				os.startfile(docPath)
				return
			except Exception:
				try:
					webbrowser.open("file:///" + docPath.replace("\\", "/"))
					return
				except Exception:
					pass
	ui.message("Nie znaleziono pliku instrukcji obsługi.")


# ==============================================================================
# SYGNALIZATOR DŹWIĘKOWY OCZEKIWANIA / POSTĘPU (BUSY SOUND TICKER)
# ==============================================================================


class BusySoundTicker:
	"""
	Uniwersalny wątek dźwiękowy informujący o trwającym przetwarzaniu w tle
	(np. skanowanie wielu grafik w wiadomości, pobieranie obrazów z sieci, Windows OCR).
	Generuje subtelne, ciche i krótkie dźwięki postępu (pipnięcia) w osobnym wątku roboczym,
	dzięki czemu działa niezawodnie nawet wtedy, gdy wątek główny jest zajęty operacjami COM.
	"""

	def __init__(self, startTone: bool = True, interval: float = 0.28, baseFreq: int = 560, endTone: bool = False):
		self.interval = interval
		self.baseFreq = baseFreq
		self.endTone = endTone
		self._stopEvent = threading.Event()
		self._thread = None
		if startTone:
			try:
				tones.beep(baseFreq, 28)
			except Exception:
				pass

	def start(self):
		if self._thread and self._thread.is_alive():
			return self
		self._stopEvent.clear()
		self._thread = threading.Thread(target=self._run, daemon=True, name="BusySoundTickerThread")
		self._thread.start()
		return self

	def _run(self):
		step = 0
		while not self._stopEvent.wait(self.interval):
			step += 1
			try:
				# Lekko zmieniająca się częstotliwość (np. 560 Hz, 600 Hz, 640 Hz), bardzo krótki impuls (20 ms)
				freq = self.baseFreq + ((step % 4) * 40)
				tones.beep(freq, 20)
			except Exception:
				pass

	def stop(self, success: bool = True):
		self._stopEvent.set()
		if self._thread and self._thread.is_alive():
			try:
				self._thread.join(timeout=0.25)
			except Exception:
				pass
		if self.endTone:
			try:
				if success:
					tones.beep(880, 32)
				else:
					tones.beep(330, 80)
			except Exception:
				pass

	def __enter__(self):
		return self.start()

	def __exit__(self, exc_type, exc_val, exc_tb):
		self.stop(success=(exc_type is None))


def showContextHelp(context: str, parent=None):
	"""
	Wyświetla kontekstową pomoc i spis skrótów klawiszowych dla konkretnego okna
	w natywnym, czytelnym i w pełni dostępnym buforze przeglądania NVDA (ui.browseableMessage).
	"""
	lang = getEffectiveLanguage()
	if lang in CUSTOM_LANGUAGES and "help" in CUSTOM_LANGUAGES[lang]:
		h = CUSTOM_LANGUAGES[lang]["help"]
		title = h.get(f"{context}_title")
		text = h.get(f"{context}_text")
		if title and text:
			ui.browseableMessage(text, title=title)
			return

	if lang == "en":
		if context == "list":
			title = "Help: Images and Attachments List"
			text = (
				"OUTLOOK IMAGE & ATTACHMENT HELPER — LIST WINDOW HELP\n"
				+ "=" * 60 + "\n\n"
				"ELEMENT CATEGORIES (Switch with RadioBox or 1-3 keys):\n"
				"  1. Inline images and screenshots — embedded visuals and graphics in message body\n"
				"  2. Dividers and separator lines — horizontal lines and decorative separators\n"
				"  3. Message attachments — files attached to email (PDF, DOCX, XLSX, images)\n\n"
				"KEYBOARD SHORTCUTS IN LIST:\n"
				"  Enter         - Recognize text (OCR) or read document content\n"
				"  P             - Open magnified image preview (for low-vision users)\n"
				"  I             - Toggle color inversion (negative mode) for selected graphic\n"
				"  Alt+C / Apps  - Open context actions menu (OCR, AI analysis, rename, save, links)\n"
				"  F2            - Rename / assign custom descriptive label to graphic (saved per message)\n"
				"  Ctrl+Enter    - Open associated hyperlink in browser or open attachment file\n"
				"  Tab           - Move to technical details box (dimensions, CID, full OCR/AI text) and action buttons\n"
				"  Alt+O         - Quick AI description / analysis\n"
				"  Alt+R         - Recognize text via OCR\n"
				"  Alt+Z         - Save selected item to file\n"
				"  Alt+W         - Save all items to a folder\n"
				"  F1            - Show this contextual help\n"
				"  Escape        - Close list window\n"
			)
		elif context == "chat":
			title = "Help: AI Conversation Window"
			text = (
				"OUTLOOK IMAGE & ATTACHMENT HELPER — AI CHAT WINDOW HELP\n"
				+ "=" * 60 + "\n\n"
				"NAVIGATION IN CHAT HISTORY:\n"
				"  H / Shift+H   - Next / Previous heading (any level)\n"
				"  2 / Shift+2   - Next / Previous question (Heading Level 2)\n"
				"  3 / Shift+3   - Next / Previous AI answer (Heading Level 3)\n"
				"  4 / Shift+4   - Next / Previous bullet point or section (Heading Level 4)\n"
				"  5 / Shift+5   - Next / Previous sub-bullet (Heading Level 5)\n"
				"  B / Shift+B   - Next / Previous button (Copy and Save buttons under each answer)\n"
				"  E             - Jump directly to the question input edit box\n"
				"  C             - Copy current or last AI answer to clipboard\n"
				"  S             - Save current or last AI answer to file (HTML, Markdown, TXT)\n"
				"  Up / Down     - Read turn text line by line\n\n"
				"ASKING QUESTIONS (In question input box):\n"
				"  Enter / Ctrl+Enter - Send question to AI\n"
				"  Shift+Enter        - Insert a new line in question\n"
				"  Polish diacritics (AltGr+C, AltGr+A, etc.) are fully protected.\n\n"
				"ACTION BUTTONS AT BOTTOM:\n"
				"  Copy Last Answer, Copy Whole Conversation, Export to File,\n"
				"  Include source image checkbox, Close (Esc).\n\n"
				"  F1            - Show this contextual help\n"
				"  Escape        - Close conversation window\n"
			)
		elif context == "preview":
			title = "Help: Magnified Image Preview"
			text = (
				"OUTLOOK IMAGE & ATTACHMENT HELPER — IMAGE PREVIEW HELP\n"
				+ "=" * 60 + "\n\n"
				"MAGNIFIER & HIGH-CONTRAST CONTROLS (For low-vision users):\n"
				"  + / = / Num+  - Zoom in (increase magnification)\n"
				"  - / Num-      - Zoom out (decrease magnification)\n"
				"  0 / D         - Fit image to window\n"
				"  1             - 100% scale (actual pixel resolution)\n"
				"  I             - Toggle color inversion (negative / high contrast mode)\n"
				"  O             - Open image in Windows default photo viewer\n"
				"  Arrows / PgUp / PgDn - Scroll around magnified image\n"
				"  F1            - Show this contextual help\n"
				"  Escape        - Close preview\n"
			)
		else:  # settings
			title = "Help: Add-on Settings"
			text = (
				"OUTLOOK IMAGE & ATTACHMENT HELPER — SETTINGS HELP\n"
				+ "=" * 60 + "\n\n"
				"AVAILABLE OPTIONS:\n"
				"- Interface Language: Auto (NVDA), Polish, English\n"
				"- Default OCR Language: Windows Media OCR recognition language\n"
				"- Sound Theme: Associative (camera shutter), Minimalist, Instrumental\n"
				"- Google Gemini API Key: Required for multimodal AI image descriptions and Q&A\n"
				"- Gemini Model: gemini-3.8-flash or custom model\n"
				"- Input Gestures button: Configure global shortcuts in NVDA Input Gestures\n"
				"- User Guide button: Open comprehensive HTML documentation in browser\n"
			)
	else:
		if context == "list":
			title = "Pomoc: Lista grafik i załączników w wiadomości"
			text = (
				"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC OKNA LISTY\n"
				+ "=" * 60 + "\n\n"
				"KATEGORIE ELEMENTÓW (Przełączaj grupą opcji lub strzałkami):\n"
				"  1. Grafiki i zrzuty ekranu w treści — rzeczywiste obrazy, wykresy i schematy\n"
				"  2. Linie podziału i separatory — poziome kreski i linie oddzielające w stopce\n"
				"  3. Załączniki do wiadomości — pliki załączone do maila (PDF, DOCX, XLSX, obrazy)\n\n"
				"SKRÓTY KLAWISZOWE W WIDOKU LISTY:\n"
				"  Enter         - Szybkie rozpoznanie tekstu (OCR) lub odczytanie treści pliku\n"
				"  P             - Powiększony podgląd grafiki (okno podglądu dla słabowidzących)\n"
				"  I             - Przełącz inwersję kolorów (negatyw) dla wybranej grafiki\n"
				"  Alt+C / Menu  - Menu czynności (OCR, opisywanie AI, zmiana nazwy, zapis, linki)\n"
				"  F2            - Nadaj własną, czytelną nazwę grafice (zapamiętywana w tej wiadomości)\n"
				"  Ctrl+Enter    - Otwórz link internetowy w przeglądarce lub otwórz plik załącznika\n"
				"  Tab           - Przejdź do pola szczegółów technicznych (wymiary, link, OCR, AI) i przycisków\n"
				"  Alt+O         - Szybki opis przez sztuczną inteligencję (AI)\n"
				"  Alt+R         - Rozpoznanie tekstu przez OCR\n"
				"  Alt+Z         - Zapisz wybrany element jako plik\n"
				"  Alt+W         - Zapisz wszystkie elementy do folderu\n"
				"  F1            - Wyświetl tę pomoc kontekstową\n"
				"  Escape        - Zamknij okno listy\n"
			)
		elif context == "chat":
			title = "Pomoc: Rozmowa z AI o grafice"
			text = (
				"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC ROZMOWY Z AI\n"
				+ "=" * 60 + "\n\n"
				"NAWIGACJA PO HISTORII ROZMOWY:\n"
				"  H / Shift+H   - Następny / Poprzedni nagłówek (dowolnego poziomu)\n"
				"  2 / Shift+2   - Następne / Poprzednie pytanie (nagłówek H2)\n"
				"  3 / Shift+3   - Następna / Poprzednia odpowiedź AI (nagłówek H3)\n"
				"  4 / Shift+4   - Następne / Poprzednie wypunktowanie lub sekcja (nagłówek H4)\n"
				"  5 / Shift+5   - Następny / Poprzedni podpunkt (nagłówek H5)\n"
				"  B / Shift+B   - Następny / Poprzedni przycisk (kopiowanie i zapis pod odpowiedziami)\n"
				"  E             - Natychmiastowy skok do pola wprowadzania pytania\n"
				"  C             - Kopiuj wybraną lub ostatnią odpowiedź do schowka\n"
				"  S             - Zapisz wybraną lub ostatnią odpowiedź do pliku (HTML, MD, TXT)\n"
				"  Strzałki      - Czytaj treść odpowiedzi wiersz po wierszu\n\n"
				"ZADAWANIE PYTAŃ (w polu edycji pytania):\n"
				"  Enter / Ctrl+Enter - Wyślij pytanie do AI\n"
				"  Shift+Enter        - Nowa linia w treści pytania\n"
				"  Polskie znaki diakrytyczne (Prawy Alt+C, Prawy Alt+A itd.) są w pełni chronione.\n\n"
				"PRZYCISKI NA DOLE OKNA:\n"
				"  Kopiuj ostatnią odpowiedź, Kopiuj całą rozmowę, Eksportuj do pliku,\n"
				"  Dołącz grafikę źródłową (pole wyboru), Zamknij (Esc).\n\n"
				"  F1            - Wyświetl tę pomoc kontekstową\n"
				"  Escape        - Zamknij okno rozmowy\n"
			)
		elif context == "preview":
			title = "Pomoc: Powiększony podgląd grafiki"
			text = (
				"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC PODGLĄDU OBRAZU\n"
				+ "=" * 60 + "\n\n"
				"STEROWANIE POWIĘKSZENIEM I KONTRASTEM DLA OSÓB SŁABOWIDZĄCYCH:\n"
				"  + / = / Num+  - Zwiększ powiększenie (przybliż)\n"
				"  - / Num-      - Zmniejsz powiększenie (oddal)\n"
				"  0 / D         - Dopasuj cały obraz do rozmiaru okna\n"
				"  1             - Skala 100% (oryginalna rozdzielczość pikselowa)\n"
				"  I             - Przełącz inwersję kolorów (tryb negatywu / wysoki kontrast)\n"
				"  O             - Otwórz obraz w domyślnej przeglądarce graficznej Windows\n"
				"  Strzałki / PgUp / PgDn - Przewijaj powiększony obraz\n"
				"  F1            - Wyświetl tę pomoc kontekstową\n"
				"  Escape        - Zamknij okno podglądu\n"
			)
		else:  # settings
			title = "Pomoc: Ustawienia dodatku"
			text = (
				"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC USTAWIEŃ\n"
				+ "=" * 60 + "\n\n"
				"DOSTĘPNE OPCJE:\n"
				"- Język interfejsu dodatku: Automatyczny (zgodny z NVDA), Polski, English\n"
				"- Domyślny język silnika OCR: Wybór języka rozpoznawania Windows Media OCR\n"
				"- Motyw dźwiękowy powiadomień: Skojarzeniowy, Minimalistyczny, Instrumentalny\n"
				"- Klucz API Google Gemini: Wymagany do multimodalnych opisów AI i rozmowy Q&A\n"
				"- Model Gemini: Wybór modelu (domyślnie: gemini-3.8-flash)\n"
				"- Przycisk 'Skróty klawiszowe (Zdarzenia wejścia)': Konfiguracja globalnych gestów NVDA\n"
				"- Przycisk 'Instrukcja obsługi': Otwiera pełną dokumentację HTML w przeglądarce\n"
			)
	ui.browseableMessage(text, title=title)


class AccessibleMessageDialog(wx.Dialog):
	"""
	W pełni dostępne okno komunikatów i instrukcji dla czytnika ekranu NVDA.
	Fokus jest automatycznie ustawiany w wielowierszowym polu tekstowym tylko do odczytu (wx.TextCtrl).
	Dzięki temu użytkownik programu NVDA od razu słyszy treść komunikatu,
	może swobodnie czytać go linijka po linijce strzałkami w górę i w dół,
	skopiować tekst (Ctrl+C), a także zamknąć okno klawiszem Enter lub Escape.
	"""

	def __init__(self, parent, message: str, title: str = "Informacja", buttonLabel: str = "&Zamknij", actionBtnLabel: str | None = None, actionCallback = None):
		super().__init__(
			parent or gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.SetMinSize((560, 320))
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		self.textCtrl = wx.TextCtrl(
			self,
			value=message,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.BORDER_THEME,
			size=(640, 340),
		)
		self.textCtrl.SetName(title)
		mainSizer.Add(self.textCtrl, 1, wx.ALL | wx.EXPAND, 12)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		if actionBtnLabel and actionCallback:
			self.btnAction = wx.Button(self, label=actionBtnLabel)
			self.btnAction.Bind(wx.EVT_BUTTON, actionCallback)
			btnSizer.Add(self.btnAction, 0, wx.RIGHT, 10)

		self.btnClose = wx.Button(self, wx.ID_OK, label=buttonLabel)
		btnSizer.Add(self.btnClose, 0, wx.ALL, 0)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 12)


		self.SetSizerAndFit(mainSizer)
		self.CenterOnScreen()

		self.SetAffirmativeId(wx.ID_OK)
		self.SetEscapeId(wx.ID_OK)

		def _onTextKeyDown(evt):
			key = evt.GetKeyCode()
			if key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER) and not evt.ShiftDown() and not evt.ControlDown():
				self.EndModal(wx.ID_OK)
				return
			elif key == wx.WXK_ESCAPE:
				self.EndModal(wx.ID_OK)
				return
			evt.Skip()

		self.textCtrl.Bind(wx.EVT_KEY_DOWN, _onTextKeyDown)

		def _focusText():
			try:
				if self.textCtrl:
					self.textCtrl.SetFocus()
					self.textCtrl.SetInsertionPoint(0)
			except Exception:
				pass

		def _onShow(evt):
			if evt.IsShown():
				wx.CallAfter(_focusText)
				wx.CallLater(50, _focusText)
			evt.Skip()

		self.Bind(wx.EVT_SHOW, _onShow)
		wx.CallAfter(_focusText)


class AccessibleConfirmDialog(wx.Dialog):
	"""
	W pełni dostępne okno potwierdzenia (Tak / Nie) dla czytnika ekranu NVDA.
	Fokus jest automatycznie ustawiany w polu tekstowym z pytaniem/ostrzeżeniem,
	co pozwala użytkownikowi NVDA od razu zapoznać się z treścią, a następnie
	przejść klawiszem Tab do przycisków 'Tak' lub 'Nie' (lub użyć Alt+T / Alt+N bądź Esc).
	"""

	def __init__(self, parent, message: str, title: str = "Potwierdzenie", defaultYes: bool = False):
		super().__init__(
			parent or gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.SetMinSize((540, 280))
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		self.textCtrl = wx.TextCtrl(
			self,
			value=message,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.BORDER_THEME,
			size=(620, 220),
		)
		self.textCtrl.SetName(title)
		mainSizer.Add(self.textCtrl, 1, wx.ALL | wx.EXPAND, 12)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		isPl = (getEffectiveLanguage() == "pl")
		yesLabel = "&Tak" if isPl else "&Yes"
		noLabel = "&Nie" if isPl else "&No"

		self.btnYes = wx.Button(self, wx.ID_YES, label=yesLabel)
		self.btnNo = wx.Button(self, wx.ID_NO, label=noLabel)

		if defaultYes:
			self.btnYes.SetDefault()
		else:
			self.btnNo.SetDefault()

		btnSizer.Add(self.btnYes, 0, wx.RIGHT, 8)
		btnSizer.Add(self.btnNo, 0, 0)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 12)

		self.SetSizerAndFit(mainSizer)
		self.CenterOnScreen()

		self.SetAffirmativeId(wx.ID_YES)
		self.SetEscapeId(wx.ID_NO)

		def _onTextKeyDown(evt):
			key = evt.GetKeyCode()
			if key == wx.WXK_ESCAPE:
				self.EndModal(wx.ID_NO)
				return
			evt.Skip()

		self.textCtrl.Bind(wx.EVT_KEY_DOWN, _onTextKeyDown)

		def _focusText():
			try:
				if self.textCtrl:
					self.textCtrl.SetFocus()
					self.textCtrl.SetInsertionPoint(0)
			except Exception:
				pass

		def _onShow(evt):
			if evt.IsShown():
				wx.CallAfter(_focusText)
				wx.CallLater(50, _focusText)
			evt.Skip()

		self.Bind(wx.EVT_SHOW, _onShow)
		wx.CallAfter(_focusText)


def showAccessibleMessageBox(message: str, title: str = "Informacja", parent=None, buttonLabel: str = "&Zamknij", actionBtnLabel: str | None = None, actionCallback = None):
	"""Wyświetla w pełni dostępne okno komunikatu z fokusem na tekście."""
	if not wx.IsMainThread():
		wx.CallAfter(showAccessibleMessageBox, message, title, parent, buttonLabel, actionBtnLabel, actionCallback)
		return
	gui.mainFrame.prePopup()
	try:
		dlg = AccessibleMessageDialog(parent, message, title=title, buttonLabel=buttonLabel, actionBtnLabel=actionBtnLabel, actionCallback=actionCallback)
		dlg.ShowModal()
		dlg.Destroy()
	finally:
		gui.mainFrame.postPopup()



def showAccessibleConfirmBox(message: str, title: str = "Potwierdzenie", parent=None, defaultYes: bool = False) -> bool:
	"""Wyświetla w pełni dostępne okno pytania Tak/Nie z fokusem na tekście pytania."""
	gui.mainFrame.prePopup()
	try:
		dlg = AccessibleConfirmDialog(parent, message, title=title, defaultYes=defaultYes)
		res = dlg.ShowModal()
		dlg.Destroy()
		return (res == wx.ID_YES)
	finally:
		gui.mainFrame.postPopup()


DEFAULT_PROMPT_PRESETS = [
	{
		"id": "detailed",
		"title": "Szczegółowy opis dla osoby niewidomej (domyślny)",
		"prompt": (
			"Jesteś precyzyjnym asystentem osoby niewidomej. "
			"Dokładnie i wyczerpująco opisz załączony obraz w języku polskim.\n"
			"- Jeśli to zrzut ekranu programu, systemu lub formularza: połącz widoczne dane z kontekstem wiadomości, "
			"podaj nagłówki kolumn, wartości w wierszach oraz istotne podsumowania.\n"
			"- Jeśli to wykres lub diagram: opisz osie, wartości, etykiety i kluczowe wnioski.\n"
			"- Jeśli to zdjęcie lub ilustracja: opisz scenę, obiekty, ludzi i kolory.\n"
			"- Jeśli na obrazie widoczne są jakiekolwiek linki, adresy WWW, adresy e-mail, numery telefonów lub lokalizacje/pozycje geograficzne, "
			"zawsze podawaj je w pełnej, dokładnej postaci (np. w formacie Markdown [Tekst](https://...) lub https://...), "
			"aby użytkownik mógł je bezpośrednio otworzyć lub skopiować."
		),
	},
	{
		"id": "table_hr",
		"title": "Tabela kadrowa / zestawienie urlopów",
		"prompt": (
			"Przeanalizuj tabelę widoczną na zrzucie ekranu pod kątem danych pracowniczych i limitów urlopowych. "
			"Wskaż: nazwisko pracownika (jeśli wynika z kontekstu lub nagłówka), definicję/rodzaj urlopu, okres ważności, "
			"przyznany limit (dni i godziny) oraz liczbę pozostałych dni i godzin do wykorzystania. "
			"Przedstaw dane w czytelnej liście punktowanej."
		),
	},
	{
		"id": "ui_form",
		"title": "Okno aplikacji, formularz lub komunikat błędu",
		"prompt": (
			"Opisz interfejs użytkownika widoczny na zrzucie ekranu. "
			"Wskaż: tytuł okna/programu, aktywne zakładki, widoczne pola edycyjne i ich wartości, "
			"stan przełączników/checkboxów, treść komunikatów o błędach lub ostrzeżeniach oraz dostępne przyciski."
		),
	},
	{
		"id": "chart_diagram",
		"title": "Wykres, diagram lub schemat danych",
		"prompt": (
			"Dokładnie zanalizuj załączony wykres lub schemat. "
			"Podaj: typ wykresu, tytuł, opis osi X i Y, jednostki miary, wartości poszczególnych serii danych, "
			"elementy legendy oraz główne trendy i wnioski wynikające z wykresu."
		),
	},
	{
		"id": "short_summary",
		"title": "Krótkie podsumowanie (2-3 zdania)",
		"prompt": "Podaj krótkie, zwięzłe podsumowanie w 2-3 zdaniach tego, co przedstawia ten obraz, skupiając się na najważniejszych informacjach.",
	},
	{
		"id": "ocr_raw",
		"title": "Dokładne przepisanie tekstu i liczb z obrazu",
		"prompt": (
			"Przepisz cały widoczny tekst, nagłówki, etykiety i liczby znajdujące się na tym obrazie, "
			"zachowując logiczny układ i hierarchię, bez zbędnych wstępów. "
			"Wszelkie adresy URL, e-maile, numery telefonów oraz pozycje geograficzne przedstaw w pełnej, dokładnej formie, "
			"z zachowaniem linków Markdown dla hiperłączy."
		),
	},
]

DEFAULT_PROMPT_PRESETS_EN = [
	{
		"id": "detailed",
		"title": "Detailed description for a blind user (default)",
		"prompt": (
			"You are a precise assistant for a blind person. "
			"Thoroughly and comprehensively describe the attached image in English.\n"
			"- If it is a software screenshot, system UI, or form: connect the visible data with email context, "
			"provide column headers, row values, status indicators, and relevant summaries.\n"
			"- If it is a chart or diagram: describe the axes, series, labels, values, and key takeaways.\n"
			"- If it is a photograph or illustration: describe the scene, objects, people, and visual style.\n"
			"- If any links, website URLs, email addresses, phone numbers, or geographic coordinates/addresses appear on the image, "
			"always provide them in complete, exact form (e.g. Markdown format [Label](https://...) or plain URL), "
			"so the user can directly open or copy them."
		),
	},
	{
		"id": "table_hr",
		"title": "HR table / leave and vacation balance",
		"prompt": (
			"Analyze the table visible in the screenshot regarding employee data and vacation/leave limits. "
			"Specify: employee name (if apparent from context/header), leave definition/type, validity period, "
			"granted limit (days and hours), and remaining balance. Present as a clean bulleted list in English."
		),
	},
	{
		"id": "ui_form",
		"title": "Application window, form, or error dialog",
		"prompt": (
			"Describe the user interface visible in the screenshot in English. "
			"Specify: window/app title, active tabs, text fields and their values, "
			"checkboxes/toggles state, error or warning messages, and available action buttons."
		),
	},
	{
		"id": "chart_diagram",
		"title": "Chart, diagram, or data schema",
		"prompt": (
			"Thoroughly analyze the attached chart or diagram in English. "
			"State: chart type, title, X and Y axes, units of measurement, data points/trends, legend entries, and key conclusions."
		),
	},
	{
		"id": "short_summary",
		"title": "Short summary (2-3 sentences)",
		"prompt": "Provide a concise summary in 2-3 sentences of what this image shows in English, focusing on the most critical information.",
	},
	{
		"id": "ocr_raw",
		"title": "Exact transcription of text and numbers",
		"prompt": "Transcribe all visible text, headings, labels, and numbers in this image exactly as they appear, preserving structure and hierarchy, without introductions.",
	},
]


_cached_prompt_presets: list[dict] | None = None


def loadPromptPresets() -> list[dict]:
	"""Ładuje zapisane szablony promptów użytkownika z pamięci podręcznej lub pliku JSON."""
	global _cached_prompt_presets
	if _cached_prompt_presets is not None:
		return list(_cached_prompt_presets)
	if os.path.isfile(PROMPTS_FILE):
		try:
			with open(PROMPTS_FILE, "r", encoding="utf-8") as f:
				data = json.load(f)
				if isinstance(data, list) and len(data) > 0:
					_cached_prompt_presets = data
					return list(data)
		except Exception as e:
			log.error(f"outlookImageHelper: Błąd odczytu {PROMPTS_FILE}: {e}")
	lang = getEffectiveLanguage()
	presets = list(DEFAULT_PROMPT_PRESETS_EN if lang == "en" else DEFAULT_PROMPT_PRESETS)
	_cached_prompt_presets = presets
	return list(presets)


def savePromptPresets(presets: list[dict]):
	"""Zapisuje listę szablonów promptów do pliku JSON oraz aktualizuje cache."""
	global _cached_prompt_presets
	_cached_prompt_presets = list(presets)
	try:
		with open(PROMPTS_FILE, "w", encoding="utf-8") as f:
			json.dump(presets, f, ensure_ascii=False, indent=2)
	except Exception as e:
		log.error(f"outlookImageHelper: Błąd zapisu {PROMPTS_FILE}: {e}")


CUSTOM_NAMES_FILE = os.path.join(getConfigDir(), "outlookImageHelper_custom_names.json")
_cached_custom_names: dict[str, dict[str, str]] | None = None


def loadAllCustomNames() -> dict[str, dict[str, str]]:
	"""Ładuje słownik własnych nazw grafik nadanych przez użytkownika (z cache w pamięci)."""
	global _cached_custom_names
	if _cached_custom_names is not None:
		return _cached_custom_names
	if os.path.isfile(CUSTOM_NAMES_FILE):
		try:
			with open(CUSTOM_NAMES_FILE, "r", encoding="utf-8") as f:
				data = json.load(f)
				if isinstance(data, dict):
					_cached_custom_names = data
					return data
		except Exception as e:
			log.error(f"outlookImageHelper: Błąd odczytu {CUSTOM_NAMES_FILE}: {e}")
	_cached_custom_names = {}
	return _cached_custom_names


def saveAllCustomNames(data: dict):
	"""Zapisuje słownik własnych nazw grafik do pliku JSON oraz aktualizuje cache w pamięci."""
	global _cached_custom_names
	_cached_custom_names = data
	try:
		with open(CUSTOM_NAMES_FILE, "w", encoding="utf-8") as f:
			json.dump(data, f, ensure_ascii=False, indent=2)
	except Exception as e:
		log.error(f"outlookImageHelper: Błąd zapisu {CUSTOM_NAMES_FILE}: {e}")


def getCustomNameForGraphic(msgKey: str, itemIndex: int) -> str:
	if not msgKey:
		return ""
	allNames = loadAllCustomNames()
	return allNames.get(msgKey, {}).get(str(itemIndex), "")


def setCustomNameForGraphic(msgKey: str, itemIndex: int, name: str):
	if not msgKey:
		return
	allNames = loadAllCustomNames()
	if msgKey not in allNames:
		allNames[msgKey] = {}
	cleanName = name.strip()
	if cleanName:
		allNames[msgKey][str(itemIndex)] = cleanName
	else:
		allNames[msgKey].pop(str(itemIndex), None)
		if not allNames[msgKey]:
			allNames.pop(msgKey, None)
	saveAllCustomNames(allNames)


IMAGE_SETTINGS_FILE = os.path.join(getConfigDir(), "outlookImageHelper_image_settings.json")
_cached_image_settings: dict[str, dict[str, dict]] | None = None


def loadAllImageSettings() -> dict[str, dict[str, dict]]:
	"""Ładuje słownik ustawień obrazów (np. inwersja) dla poszczególnych wiadomości (z cache w pamięci)."""
	global _cached_image_settings
	if _cached_image_settings is not None:
		return _cached_image_settings
	if os.path.isfile(IMAGE_SETTINGS_FILE):
		try:
			with open(IMAGE_SETTINGS_FILE, "r", encoding="utf-8") as f:
				data = json.load(f)
				if isinstance(data, dict):
					_cached_image_settings = data
					return data
		except Exception as e:
			log.error(f"outlookImageHelper: Błąd odczytu {IMAGE_SETTINGS_FILE}: {e}")
	_cached_image_settings = {}
	return _cached_image_settings


def saveAllImageSettings(data: dict):
	"""Zapisuje słownik ustawień obrazów do pliku JSON oraz aktualizuje cache w pamięci."""
	global _cached_image_settings
	_cached_image_settings = data
	try:
		with open(IMAGE_SETTINGS_FILE, "w", encoding="utf-8") as f:
			json.dump(data, f, ensure_ascii=False, indent=2)
	except Exception as e:
		log.error(f"outlookImageHelper: Błąd zapisu {IMAGE_SETTINGS_FILE}: {e}")


def getImageSetting(msgKey: str, itemIndex: int, key: str, default=None):
	if not msgKey:
		return default
	allSettings = loadAllImageSettings()
	return allSettings.get(msgKey, {}).get(str(itemIndex), {}).get(key, default)


def setImageSetting(msgKey: str, itemIndex: int, key: str, value):
	if not msgKey:
		return
	allSettings = loadAllImageSettings()
	if msgKey not in allSettings:
		allSettings[msgKey] = {}
	itemKey = str(itemIndex)
	if itemKey not in allSettings[msgKey]:
		allSettings[msgKey][itemKey] = {}
	if value is None:
		allSettings[msgKey][itemKey].pop(key, None)
		if not allSettings[msgKey][itemKey]:
			allSettings[msgKey].pop(itemKey, None)
			if not allSettings[msgKey]:
				allSettings.pop(msgKey, None)
	else:
		allSettings[msgKey][itemKey][key] = value
	saveAllImageSettings(allSettings)


def isImageInverted(msgKey: str, itemIndex: int) -> bool:
	"""Sprawdza, czy dla danej grafiki w wiadomości zapamiętano tryb inwersji kolorów (negatyw)."""
	return bool(getImageSetting(msgKey, itemIndex, "inverted", False))


def setImageInverted(msgKey: str, itemIndex: int, inverted: bool):
	"""Zapisuje lub usuwa ustawienie inwersji kolorów dla wskazanej grafiki w wiadomości."""
	setImageSetting(msgKey, itemIndex, "inverted", True if inverted else None)


# Wstępnie obliczona tablica translacji bajtów do natychmiastowej inwersji barw RGB (0->255, 255->0)
_INVERT_TABLE = bytes(range(255, -1, -1))


def invertWxImage(img: wx.Image) -> wx.Image:
	"""
	Dokonuje błyskawicznej inwersji barw (negatyw) obiektu wx.Image w kodzie C Pythona.
	Zachowuje nienaruszony kanał przezroczystości (Alpha) oraz ewentualną maskę koloru.
	"""
	if not img or not img.IsOk():
		return img
	try:
		rgbData = img.GetData()
		if not rgbData:
			return img
		invertedRgb = rgbData.translate(_INVERT_TABLE)
		inv = img.Copy()
		inv.SetData(invertedRgb)
		if img.HasAlpha():
			inv.SetAlpha(img.GetAlpha())
		if img.HasMask():
			inv.SetMaskColour(255 - img.GetMaskRed(), 255 - img.GetMaskGreen(), 255 - img.GetMaskBlue())
		return inv
	except Exception as e:
		log.error(f"outlookImageHelper: invertWxImage error: {e}")
		return img


def createInvertedPngFile(srcPath: str) -> str | None:
	"""Wczytuje plik obrazu z dysku, odwraca kolory i zapisuje tymczasowy plik PNG z negatywem."""
	if not srcPath or not os.path.isfile(srcPath):
		return None
	try:
		img = wx.Image(srcPath, wx.BITMAP_TYPE_ANY)
		if not img.IsOk():
			return None
		invImg = invertWxImage(img)
		tempDir = tempfile.gettempdir()
		tempInverted = os.path.join(tempDir, f"nvda_oih_inverted_{int(time.time() * 1000)}.png")
		invImg.SaveFile(tempInverted, wx.BITMAP_TYPE_PNG)
		return tempInverted
	except Exception as e:
		log.error(f"outlookImageHelper: createInvertedPngFile error: {e}")
		return None


def replaceGraphicInEmailWithInverted(item, parentDialog=None):
	"""
	Próbuje zastąpić grafikę w edytowalnej wiadomości Outlook wersją w negatywie.
	Jeśli wiadomość jest w trybie tylko do odczytu (chroniona przez WordMail),
	informuje użytkownika o konieczności włączenia trybu edycji (Akcje -> Edytuj wiadomość)
	oraz o możliwości natychmiastowego bezpiecznego oglądania grafiki w negatywie w podglądzie (P lub I).
	"""
	if not item or not hasattr(item, "rawObject") or getattr(item, "isDivider", False):
		ui.message(_T("invert_msg_not_graphic"))
		return

	if not getattr(item, "rawObject", None):
		ui.message(_T("invert_msg_no_raw"))
		return

	# Wyodrębnij obraz źródłowy
	tempPng = extractGraphicToPngFile(item)
	if not tempPng or not os.path.isfile(tempPng):
		ui.message(_T("invert_msg_extract_failed"))
		return

	invPng = createInvertedPngFile(tempPng)
	try:
		if tempPng and os.path.isfile(tempPng):
			os.remove(tempPng)
	except Exception:
		pass

	if not invPng or not os.path.isfile(invPng):
		ui.message(_T("invert_msg_invert_failed"))
		return

	try:
		raw = item.rawObject
		doc = None
		try:
			app = getattr(raw, "Application", None)
			doc = getattr(app, "ActiveDocument", None)
		except Exception:
			pass
		if not doc:
			doc = getActiveWordDocument()

		if not doc:
			raise RuntimeError("Brak dostępu do obiektu Document")

		if item.isInline:
			rng = raw.Range
			oldW = raw.Width
			oldH = raw.Height
			# Wstaw nowy obraz w miejscu dotychczasowej grafiki inline
			newShape = doc.InlineShapes.AddPicture(FileName=invPng, LinkToFile=False, SaveWithDocument=True, Range=rng)
			try:
				newShape.Width = oldW
				newShape.Height = oldH
			except Exception:
				pass
			raw.Delete()
			item.rawObject = newShape
		else:
			anchor = raw.Anchor
			oldW = raw.Width
			oldH = raw.Height
			newShape = doc.Shapes.AddPicture(FileName=invPng, LinkToFile=False, SaveWithDocument=True, Anchor=anchor)
			try:
				newShape.Width = oldW
				newShape.Height = oldH
			except Exception:
				pass
			raw.Delete()
			item.rawObject = newShape

		# Zapisz stan inwersji
		item.isInverted = True
		msgKey = getattr(item, "msgKey", "") or getMessageKey()
		idx = getattr(item, "index", 0)
		if msgKey and idx:
			setImageInverted(msgKey, idx, True)

		if parentDialog and hasattr(parentDialog, "refreshListLabels"):
			parentDialog.refreshListLabels()

		ui.message(_T("invert_msg_replaced_success"))
	except Exception as ex:
		log.warning(f"outlookImageHelper: replaceGraphicInEmailWithInverted lock/error: {ex}")
		# Zapamiętaj inwersję w bazie lokalnej mimo braku uprawnień zapisu w dokumencie
		item.isInverted = True
		msgKey = getattr(item, "msgKey", "") or getMessageKey()
		idx = getattr(item, "index", 0)
		if msgKey and idx:
			setImageInverted(msgKey, idx, True)
		if parentDialog and hasattr(parentDialog, "refreshListLabels"):
			parentDialog.refreshListLabels()

		msg = _T("invert_msg_locked_warning")
		def showLockNotice():
			showAccessibleMessageBox(msg, _T("dlg_elements_title"), parent=parentDialog)
		wx.CallAfter(showLockNotice)
	finally:
		try:
			if invPng and os.path.isfile(invPng):
				os.remove(invPng)
		except Exception:
			pass


def decodeSafeLinks(url: str) -> str:
	"""Odwija prawdziwy adres docelowy z linków ochronnych Microsoft SafeLinks oraz Proofpoint."""
	if not url:
		return ""
	urlLower = url.lower()
	if "safelinks.protection.outlook.com" in urlLower:
		try:
			parsed = urllib.parse.urlparse(url)
			query = urllib.parse.parse_qs(parsed.query)
			if "url" in query and query["url"]:
				return urllib.parse.unquote(query["url"][0])
		except Exception:
			pass
	elif "urldefense.proofpoint.com" in urlLower or "urldefense.com" in urlLower:
		try:
			parsed = urllib.parse.urlparse(url)
			query = urllib.parse.parse_qs(parsed.query)
			if "u" in query and query["u"]:
				uVal = query["u"][0].replace("-3A", ":").replace("-2F", "/").replace("-5F", "_")
				return urllib.parse.unquote(uVal)
			if "/v3/__" in url:
				parts = url.split("/v3/__", 1)[1].split("__;", 1)
				if parts:
					return parts[0]
		except Exception:
			pass
	return url


def getCleanDomain(url: str) -> str:
	"""Zwraca skróconą, czytelną domenę z adresu URL (np. youtube.com, linkedin.com)."""
	if not url:
		return ""
	try:
		realUrl = decodeSafeLinks(url)
		parsed = urllib.parse.urlparse(realUrl)
		netloc = parsed.netloc or parsed.path
		if netloc.startswith("www."):
			netloc = netloc[4:]
		return netloc.split(":")[0]
	except Exception:
		return ""


def getLinkCategoryAndName(url: str) -> tuple[str, str]:
	"""
	Rozpoznaje kategorię i czytelną nazwę usługi dla adresu URL:
	Zwraca: (kategoria, nazwa_usługi)
	Kategorie: 'social', 'phone', 'email', 'web', ''
	"""
	if not url:
		return "", ""
	realUrl = decodeSafeLinks(url).strip()
	lower = realUrl.lower()

	if lower.startswith("mailto:"):
		addr = realUrl.split(":", 1)[1].split("?")[0].strip()
		return "email", f"E-mail ({addr})" if addr else "Napisz e-mail"

	if lower.startswith("tel:"):
		num = realUrl.split(":", 1)[1].split("?")[0].strip()
		return "phone", f"Telefon ({num})" if num else "Zadzwoń"

	socialMap = [
		("linkedin.com", "LinkedIn"),
		("facebook.com", "Facebook"),
		("fb.me", "Facebook"),
		("youtube.com", "YouTube"),
		("youtu.be", "YouTube"),
		("twitter.com", "Twitter (X)"),
		("x.com", "Twitter (X)"),
		("instagram.com", "Instagram"),
		("tiktok.com", "TikTok"),
		("github.com", "GitHub"),
		("t.me", "Telegram"),
		("whatsapp.com", "WhatsApp"),
		("wa.me", "WhatsApp"),
		("pinterest.com", "Pinterest"),
		("spotify.com", "Spotify"),
	]
	for domain, name in socialMap:
		if domain in lower:
			return "social", name

	dom = getCleanDomain(realUrl)
	return "web", dom or "Strona WWW"


def getBlockedImageInstructionsText(item=None, actionName: str = "") -> str:
	isPl = (getEffectiveLanguage() == "pl")
	if isPl:
		actionPart = f"wykonania czynności: {actionName}" if actionName else "przetworzenia tego obrazu (OCR, podgląd, opis AI)"
		lines = [
			f"Obraz nie może zostać wyświetlony w celu {actionPart}.",
			"",
			"POWÓD: Program Outlook zablokował automatyczne pobieranie zewnętrznych grafik w tej wiadomości ze względów ochrony prywatności (na ekranie widoczna jest pusta ramka lub ikona czerwonego krzyżyka).",
			"",
			"Aby odblokować i pobrać obrazy, wykonaj następujące proste kroki w programie Outlook:",
			"",
			"KROK 1: Zamknij to okno (naciśnij klawisz Enter lub Escape) i upewnij się, że jesteś w oknie wiadomości w programie Outlook.",
			"",
			"KROK 2: W oknie wiadomości naciśnij skrót klawiszowy:",
			"        Ctrl + Shift + W",
			"        (Skrót ten natychmiast przenosi kursor na pasek informacyjny 'InfoBar' u góry wiadomości i rozwija jego menu).",
			"",
			"KROK 3: Naciśnij klawisz Enter.",
			"        (Pierwsza opcja w menu paska to 'Pobierz obrazy').",
			"",
			"KROK 4: Program Outlook natychmiast pobierze wszystkie zdjęcia w tej wiadomości.",
			"",
			"KROK 5: Po pobraniu obrazów naciśnij ponownie skrót Ctrl + NVDA + A (lub klawisz aplikacji / Shift+F10 na grafice).",
			"        Wszystkie obrazy będą od tej pory w pełni załadowane i gotowe do podglądu powiększonego, rozpoznawania tekstu (OCR) oraz analizy przez sztuczną inteligencję (AI)!",
			"",
			"--- Wskazówka alternatywna ---",
			"W otwartej wiadomości możesz również kliknąć myszą na pasek informacyjny u góry ('Kliknij tutaj, aby pobrać obrazy. Aby pomóc chronić prywatność...') lub nacisnąć Shift+F10 w treści wiadomości i wybrać opcję 'Pobierz obrazy'.",
			"",
			"--- Jak na stałe włączyć automatyczne pobieranie obrazów w programie Outlook ---",
			"Jeśli nie chcesz każdorazowo odblokowywać obrazów, możesz włączyć ich automatyczne pobieranie na stałe:",
			"1. W programie Outlook wejdź w menu: Plik -> Opcje (skrót klawiszowy: Alt+P, a następnie O).",
			"2. W oknie opcji przejdź strzałką w dół na zakładkę 'Centrum zaufania'.",
			"3. Klawiszem Tab przejdź na przycisk 'Ustawienia Centrum zaufania...' i naciśnij Enter.",
			"4. W kolejnym oknie przejdź strzałką na pozycję 'Automatyczne pobieranie'.",
			"5. Odznacz pole wyboru: 'Nie pobieraj automatycznie obrazów w standardowych wiadomościach e-mail w formacie HTML lub elementach RSS' (lub zaznacz pobieranie od zaufanych nadawców).",
			"6. Naciśnij Enter lub wybierz przycisk 'OK', a w oknie głównym Opcji ponownie 'OK'.",
			"Od tego momentu program Outlook będzie zawsze automatycznie pobierał i wyświetlał wszystkie grafiki w Twoich mailach.",
		]
		if item:
			desc = getattr(item, "altText", "") or getattr(item, "title", "") or getattr(item, "customName", "")
			url = getattr(item, "sourceUrl", "") or getattr(item, "hyperlinkUrl", "") or getattr(item, "cid", "")
			realUrl = decodeSafeLinks(url)
			dom = getCleanDomain(realUrl) or getattr(item, "linkServiceName", "") or ""
			extra = []
			if desc:
				extra.append(f"Opis elementu: \"{desc}\"")
			if dom:
				extra.append(f"Serwis / domena: {dom}")
			if getattr(item, "widthPx", 0) and getattr(item, "heightPx", 0):
				extra.append(f"Wymiary w wiadomości: {item.widthPx} × {item.heightPx} px")
			if realUrl and realUrl.startswith(("http://", "https://")):
				extra.append(f"Adres URL grafiki: {realUrl}")
			if extra:
				lines.extend(["", "Informacje o wybranym elemencie:"] + extra)
		return "\n".join(lines)
	else:
		actionPart = f"perform {actionName}" if actionName else "process this image (OCR, preview, AI)"
		lines = [
			f"The image cannot be loaded to {actionPart}.",
			"",
			"REASON: Outlook has blocked automatic downloading of external pictures in this message to protect privacy (a blank frame or red 'X' placeholder is displayed).",
			"",
			"To unblock and download the pictures, follow these simple steps in Outlook:",
			"",
			"STEP 1: Close this dialog (press Enter or Escape) and return to the Outlook message window.",
			"",
			"STEP 2: In the message window, press the shortcut:",
			"        Ctrl + Shift + W",
			"        (This moves focus to the Information Bar at the top of the message and opens its menu).",
			"",
			"STEP 3: Press Enter.",
			"        (The first option in the menu is 'Download Pictures').",
			"",
			"STEP 4: Outlook will download all pictures in the message.",
			"",
			"STEP 5: Press Ctrl + NVDA + A again (or Shift + F10 on the graphic).",
			"        All pictures will now be fully loaded and available for Zoom Preview, OCR text recognition, and AI description!",
			"",
			"--- Alternative Tip ---",
			"In the open message, you can also click the InfoBar at the top ('Click here to download pictures') or press Shift+F10 in the message body and select 'Download pictures'.",
			"",
			"--- How to permanently enable automatic picture downloading in Outlook ---",
			"If you do not want to unblock images manually each time, you can enable automatic downloading permanently:",
			"1. In Outlook, go to: File -> Options (shortcut: Alt+F, then T or O).",
			"2. In the Outlook Options dialog, select 'Trust Center'.",
			"3. Press Tab to find 'Trust Center Settings...' and press Enter.",
			"4. In the Trust Center window, select 'Automatic Download'.",
			"5. Uncheck the checkbox: 'Don\'t download pictures automatically in standard HTML email messages or RSS items' (or customize trusted senders).",
			"6. Click 'OK', and in the main Options dialog click 'OK' again to save.",
			"From now on, Outlook will automatically download and display all images in your emails.",
		]
		if item:
			desc = getattr(item, "altText", "") or getattr(item, "title", "") or getattr(item, "customName", "")
			url = getattr(item, "sourceUrl", "") or getattr(item, "hyperlinkUrl", "") or getattr(item, "cid", "")
			realUrl = decodeSafeLinks(url)
			dom = getCleanDomain(realUrl) or getattr(item, "linkServiceName", "") or ""
			extra = []
			if desc:
				extra.append(f"Description: \"{desc}\"")
			if dom:
				extra.append(f"Domain: {dom}")
			if getattr(item, "widthPx", 0) and getattr(item, "heightPx", 0):
				extra.append(f"Dimensions: {item.widthPx} × {item.heightPx} px")
			if realUrl and realUrl.startswith(("http://", "https://")):
				extra.append(f"URL: {realUrl}")
			if extra:
				lines.extend(["", "Selected element info:"] + extra)
		return "\n".join(lines)


class BlockedImageInstructionsDialog(AccessibleMessageDialog):
	"""
	W pełni dostępne okno z instrukcją odblokowania obrazów w programie Outlook.
	Wyjaśnia krok po kroku procedurę odblokowania przez Ctrl+Shift+W -> Enter.
	"""

	def __init__(self, parent, item=None, actionName: str = ""):
		isPl = (getEffectiveLanguage() == "pl")
		title = (
			"Instrukcja: Jak odblokować i pobrać obrazy w programie Outlook"
			if isPl else
			"Instructions: How to Unblock and Download Pictures in Outlook"
		)
		btnLabel = "&Zamknij (Esc)" if isPl else "&Close (Esc)"
		text = getBlockedImageInstructionsText(item=item, actionName=actionName)
		super().__init__(parent, message=text, title=title, buttonLabel=btnLabel)


def showBlockedImageInstructionsDialog(parent, item=None, actionName: str = ""):
	"""Wyświetla okno z instrukcją odblokowania obrazów w programie Outlook."""
	if not wx.IsMainThread():
		wx.CallAfter(showBlockedImageInstructionsDialog, parent, item, actionName)
		return
	gui.mainFrame.prePopup()
	try:
		dlg = BlockedImageInstructionsDialog(parent, item=item, actionName=actionName)
		dlg.ShowModal()
		dlg.Destroy()
	finally:
		gui.mainFrame.postPopup()


DownloadWebImageDialog = BlockedImageInstructionsDialog


class GraphicItem:
	"""Reprezentacja pojedynczego obiektu graficznego w dokumencie WordMail."""

	def __init__(
		self,
		index: int,
		isInline: bool,
		rawObject,
		widthPt: float,
		heightPt: float,
		altText: str,
		title: str,
		rangeStart: int = 0,
		rangeEnd: int = 0,
		cid: str = "",
		contextText: str = "",
		hyperlinkUrl: str = "",
		hyperlinkTip: str = "",
		customName: str = "",
		msgKey: str = "",
		isInverted: bool = False,
		sourceUrl: str = "",
	):
		self.index = index
		self.isInline = isInline
		self.rawObject = rawObject
		self.widthPt = widthPt
		self.heightPt = heightPt
		# Przybliżona konwersja punktów na piksele ekranowe (1 pt ≈ 1.33 px)
		self.widthPx = int(round(widthPt * 1.333))
		self.heightPx = int(round(heightPt * 1.333))
		self.altText = (altText or "").strip()
		self.title = (title or "").strip()
		self.rangeStart = rangeStart
		self.rangeEnd = rangeEnd
		self.cid = cid or ""
		self.contextText = (contextText or "").strip()
		self.hyperlinkUrl = (hyperlinkUrl or "").strip()
		self.hyperlinkTip = (hyperlinkTip or "").strip()
		self.customName = (customName or "").strip()
		self.msgKey = msgKey or ""
		self.isInverted = bool(isInverted)
		self.sourceUrl = (sourceUrl or "").strip()
		self.isDownloaded: bool = False
		self._downloadedTempPath: str = ""
		self._extractedTempPath: str = ""
		self._hasLocalRenderedPixels: bool = False

		# Pamięć podręczna (cache) wyników oraz rozmowy
		self.ocrText: str = ""
		self.aiDescription: str = ""
		self.chatHistory: list[dict] = []
		self.base64Image: str = ""
		self.isBlankOrSolid: bool = False
		self.blankReason: str = ""

	@property
	def isExternalWebImage(self) -> bool:
		"""
		Czy grafika jest zewnętrznym obrazem WWW, który NIE został pobrany przez program Outlook
		(zablokowany przez zasady prywatności Outlooka lub pusty placeholder w dokumencie).
		Jeśli obraz został już wyrenderowany lokalnie w WordMail lub pobrany, zwraca False.
		"""
		if getattr(self, "isDownloaded", False):
			return False
		if getattr(self, "_hasLocalRenderedPixels", False):
			return False
		# Obraz jest zewnętrzny i niepobrany tylko wtedy, gdy w WordMail jest potwierdzonym pustym placeholderem
		# i posiada źródłowy adres URL WWW:
		if getattr(self, "isBlankOrSolid", False):
			src = getattr(self, "sourceUrl", "") or ""
			if src.startswith("http://") or src.startswith("https://"):
				return True
			c = getattr(self, "cid", "") or ""
			if c.startswith("http://") or c.startswith("https://"):
				return True
			if getattr(self, "hyperlinkUrl", ""):
				return True
		return False

	@property
	def isDivider(self) -> bool:
		"""Czy grafika jest linią, separatorem, pustym ozdobnikiem, niewidocznym polem lub kształtem wektorowym/AutoShape."""
		# 1. Kształty wektorowe programu Outlook/Word (AutoShape, TextBox, Group) bez rastrowych pikseli
		# trafiają do sekcji elementów ozdobnych i separatorów:
		if self.isVectorShape:
			return True

		# 2. Jeśli to obraz zewnętrzny ze strony internetowej wymagający pobrania,
		# NIE klasyfikuj go jako separator (chyba że to mikroskopijny piksel <= 3px lub jawny spacer)
		if self.isExternalWebImage and not getattr(self, "isDownloaded", False):
			if (self.widthPx <= 3 and self.heightPx <= 3) or (self.widthPt <= 2.5 and self.heightPt <= 2.5):
				return True
			checkText = f"{self.altText} {self.title} {self.cid} {getattr(self, 'sourceUrl', '')}".lower()
			if any(w in checkText for w in ("tracking", "1x1", "pixel", "clear.gif")):
				return True
			return False

		if getattr(self, "isBlankOrSolid", False):
			return True
		# Linie i mikro-punkty
		if (self.widthPx <= 3 or self.heightPx <= 3) or (self.widthPt <= 2.5 or self.heightPt <= 2.5):
			return True
		# Paski poziome lub pionowe
		if (self.widthPx > 100 and self.heightPx <= 10) or (self.heightPx > 100 and self.widthPx <= 10):
			return True
		# Kwadratowe lub małe spacery / ozdobniki formatujące (do 36x36 px) pozbawione opisu, linku i CID
		if (self.widthPx <= 36 and self.heightPx <= 36) and not self.altText and not self.title and not self.hyperlinkUrl and not self.cid:
			return True
		checkText = f"{self.altText} {self.title} {self.cid}".lower()
		if any(w in checkText for w in ("spacer", "blank", "tracking", "pixel", "divider", "separator", "clear.gif")):
			return True
		return False

	@property
	def linkCategory(self) -> str:
		cat, _ = getLinkCategoryAndName(self.hyperlinkUrl)
		return cat

	@property
	def linkServiceName(self) -> str:
		_, name = getLinkCategoryAndName(self.hyperlinkUrl)
		return name

	@property
	def shapeType(self) -> int:
		"""Zwraca identyfikator typu kształtu Worda (np. 1=msoAutoShape, 11=msoLinkedPicture, 13=msoPicture, 17=msoTextBox)."""
		if getattr(self, "rawObject", None):
			try:
				return int(getattr(self.rawObject, "Type", 0) or 0)
			except Exception:
				pass
		return 0

	@property
	def isVectorShape(self) -> bool:
		"""Czy obiekt jest kształtem wektorowym / autokształtem / polem tekstowym Worda, a nie rastrowym obrazem."""
		if not self.isInline:
			st = self.shapeType
			return st not in (11, 13) and st > 0
		return False

	@property
	def isScreenshotOrUnlabeled(self) -> bool:
		"""Czy grafika prawdopodobnie jest zrzutem ekranu lub nie posiada opisu."""
		if self.isDivider:
			return False
		if self.linkCategory in ("social", "phone", "email"):
			return False
		return not self.customName and not self.altText and not self.title

	def getDisplayLabel(self, totalDocLength: int = 0) -> str:
		"""Tworzy zwięzłą, czytelną i dostępną etykietę dla listy w oknie dialogowym bez zbędnego szumu i bez powtarzania 'Grafika'."""
		rep = getattr(self, "_repeatCount", 1)
		repSuffix = f" (powtórzona {rep}× w wiadomości)" if rep > 1 else ""

		# 1. Własna nazwa nadana przez użytkownika
		if self.customName:
			parts = [f"{self.index}: [{self.customName}]"]
			if self.isExternalWebImage and not self.isDownloaded:
				webUrl = self.sourceUrl or self.hyperlinkUrl or self.cid
				dom = getCleanDomain(decodeSafeLinks(webUrl)) or self.linkServiceName or "internet"
				parts.append(f"[Obraz w internecie: {dom}] — [Zablokowany przez Outlook — odblokuj: Ctrl+Shift+W]")
			elif self.hyperlinkUrl:
				svc = self.linkServiceName or getCleanDomain(self.hyperlinkUrl)
				parts.append(f"[Link: {svc}]")
			if self.ocrText:
				if self.ocrText == "(Brak rozpoznanego tekstu)":
					parts.append(f"[{_T('tag_text_empty')}]")
				else:
					snip = _cleanSnippet(self.ocrText)
					parts.append(f"[{_T('tag_text')}: \"{snip}\"]" if snip else f"[{_T('tag_ocr_ready')}]")
			if self.aiDescription:
				snip = _cleanSnippet(self.aiDescription)
				parts.append(f"[{_T('tag_ai')}: \"{snip}\"]" if snip else f"[{_T('tag_ai_ready')}]")
			if self.isInverted:
				parts.append(_T("tag_inverted"))
			if repSuffix:
				parts.append(repSuffix.strip())
			return " — ".join(parts)

		# 1b. Obraz zewnętrzny ze strony WWW niepobrany przez Outlook
		if self.isExternalWebImage and not self.isDownloaded:
			webUrl = self.sourceUrl or self.hyperlinkUrl or self.cid
			dom = getCleanDomain(decodeSafeLinks(webUrl)) or self.linkServiceName or "internet"
			parts = []
			desc = self.altText or self.title
			if desc:
				parts.append(f"{self.index}: \"{desc}\"")
			elif self.contextText:
				ctx = _cleanSnippet(self.contextText, maxLen=40)
				if ctx:
					parts.append(f"{self.index}: \"{ctx}\"")
				else:
					parts.append(f"{self.index}: [Obraz w internecie: {dom}] ({self.widthPx}×{self.heightPx} px)")
			else:
				parts.append(f"{self.index}: [Obraz w internecie: {dom}] ({self.widthPx}×{self.heightPx} px)")

			if not any("[Obraz w internecie" in p for p in parts):
				parts.append(f"[Obraz w internecie: {dom}]")
			parts.append("[Zablokowany przez Outlook — odblokuj: Ctrl+Shift+W]")

			if self.ocrText:
				if self.ocrText == "(Brak rozpoznanego tekstu)":
					parts.append(f"[{_T('tag_text_empty')}]")
				else:
					snip = _cleanSnippet(self.ocrText)
					parts.append(f"[{_T('tag_text')}: \"{snip}\"]" if snip else f"[{_T('tag_ocr_ready')}]")
			if self.aiDescription:
				snip = _cleanSnippet(self.aiDescription)
				parts.append(f"[{_T('tag_ai')}: \"{snip}\"]" if snip else f"[{_T('tag_ai_ready')}]")
			if self.isInverted:
				parts.append(_T("tag_inverted"))
			if repSuffix:
				parts.append(repSuffix.strip())
			return " — ".join(parts)

		# 2. Linia podziału / separator / pusty ozdobnik / kształt wektorowy
		if self.isDivider:
			if self.isVectorShape:
				mainDesc = self.altText or self.title
				if mainDesc:
					divLbl = f"{self.index}: [Kształt/Przycisk: \"{mainDesc}\"] ({self.widthPx}×{self.heightPx} px)"
				else:
					divLbl = f"{self.index}: [Kształt wektorowy / AutoShape] ({self.widthPx}×{self.heightPx} px)"
				if self.ocrText:
					if self.ocrText == "(Brak rozpoznanego tekstu)":
						divLbl += f" — [{_T('tag_text_empty')}]"
					else:
						snip = _cleanSnippet(self.ocrText)
						divLbl += f" — [{_T('tag_text')}: \"{snip}\"]" if snip else f" — [{_T('tag_ocr_ready')}]"
			elif getattr(self, "isBlankOrSolid", False):
				reason = getattr(self, "blankReason", "") or "Pusty ozdobnik"
				divLbl = f"{self.index}: [{reason}] ({self.widthPx}×{self.heightPx} px)"
			elif (self.widthPx <= 36 and self.heightPx <= 36) and not self.altText and not self.title and not self.hyperlinkUrl and not self.cid:
				divLbl = f"{self.index}: [Pusty ozdobnik/spacer] ({self.widthPx}×{self.heightPx} px)"
			elif self.widthPx <= 3 or self.heightPx <= 3:
				divLbl = f"{self.index}: [{_T('tag_divider_point')}] ({self.widthPx}×{self.heightPx} px)"
			else:
				divLbl = f"{self.index}: {_T('label_divider')}"
			if self.hyperlinkUrl:
				svc = self.linkServiceName or getCleanDomain(self.hyperlinkUrl)
				divLbl += f" [Link: {svc}]" if svc else f" [{_T('tag_link')}]"
			if self.isInverted:
				divLbl += f" — {_T('tag_inverted')}"
			if repSuffix:
				divLbl += repSuffix
			return divLbl

		# 3. Ikony mediów społecznościowych, telefonu, e-maila
		cat = self.linkCategory
		svc = self.linkServiceName
		if cat == "social":
			lbl = f"{self.index}: {_T('label_social_icon', svc=svc)}"
		elif cat == "phone":
			lbl = f"{self.index}: {_T('label_phone_icon', svc=svc)}"
		elif cat == "email":
			lbl = f"{self.index}: {_T('label_email_icon', svc=svc)}"
		else:
			lbl = ""

		parts = []
		if lbl:
			parts.append(lbl)
		else:
			mainDesc = self.altText or self.title
			if mainDesc:
				parts.append(f"{self.index}: \"{mainDesc}\"")
			elif self.hyperlinkUrl:
				dom = svc or getCleanDomain(self.hyperlinkUrl)
				parts.append(f"{self.index}: [Link: {dom}]")
			else:
				parts.append(f"{self.index}: {_T('label_unlabeled_img')}")

		# Jeśli nie było linku w głównej części, a link istnieje (dla zwykłych grafik z linkiem)
		if self.hyperlinkUrl and not any("[Link:" in p for p in parts) and cat not in ("social", "phone", "email"):
			dom = svc or getCleanDomain(self.hyperlinkUrl)
			parts.append(f"[Link: {dom}]")

		# OCR fragment tekstu
		if self.ocrText:
			if self.ocrText == "(Brak rozpoznanego tekstu)":
				parts.append(f"[{_T('tag_text_empty')}]")
			else:
				snip = _cleanSnippet(self.ocrText)
				if snip:
					parts.append(f"[{_T('tag_text')}: \"{snip}\"]")
				else:
					parts.append(f"[{_T('tag_ocr_ready')}]")

		# AI fragment opisu
		if self.aiDescription:
			snip = _cleanSnippet(self.aiDescription)
			if snip:
				parts.append(f"[{_T('tag_ai')}: \"{snip}\"]")
			else:
				parts.append(f"[{_T('tag_ai_ready')}]")

		if self.isInverted:
			parts.append(_T("tag_inverted"))

		if repSuffix:
			parts.append(repSuffix.strip())

		return " — ".join(parts)

	def getDetailedInfo(self) -> str:
		"""Zwraca pełny raport tekstowy o wybranej grafice."""
		colorStatus = _T("detail_colors_inverted") if self.isInverted else _T("detail_colors_normal")
		if self.isInline:
			typeStr = "Osadzona w linii tekstu (InlineShape)"
		elif self.isVectorShape:
			typeStr = f"Swobodna z opływaniem (Kształt wektorowy / AutoShape, typ {self.shapeType})"
		else:
			typeStr = "Swobodna z opływaniem (Shape)"
		lines = [
			f"Numer obiektu: {self.index}",
			f"Typ: {typeStr}",
			f"Wymiary: {self.widthPx} × {self.heightPx} px ({self.widthPt:.1f} × {self.heightPt:.1f} pt)",
			_T("detail_colors_lbl", status=colorStatus),
		]
		if self.isVectorShape:
			lines.append("Uwaga: Obiekt jest kształtem wektorowym / przyciskiem Worda, a nie plikiem graficznym.")
		if self.customName:
			lines.append(f"Własna nazwa użytkownika: {self.customName}")
		if self.altText:
			lines.append(f"Tekst alternatywny: {self.altText}")
		if self.title:
			lines.append(f"Tytuł: {self.title}")
		if self.hyperlinkUrl:
			realUrl = decodeSafeLinks(self.hyperlinkUrl)
			lines.append(f"Powiązany link (URL): {realUrl}")
			if self.linkServiceName:
				lines.append(f"Rozpoznana usługa / cel: {self.linkServiceName}")
			if self.hyperlinkTip:
				lines.append(f"Etykieta linku: {self.hyperlinkTip}")
		if self.cid:
			lines.append(f"Źródło (Content-ID): {self.cid}")
		if self.rangeEnd > 0:
			lines.append(f"Pozycja w dokumencie (znaki): {self.rangeStart} - {self.rangeEnd}")
		if self.contextText:
			lines.append(f"Tekst kontekstowy nad grafiką w mailu: {self.contextText}")
		if self.ocrText:
			firstLine = self.ocrText.strip().splitlines()[0] if self.ocrText.strip() else ""
			if len(firstLine) > 100:
				firstLine = firstLine[:97] + "..."
			lines.append(f"Rozpoznany tekst OCR: {firstLine}")
		if self.aiDescription:
			firstLine = self.aiDescription.strip().splitlines()[0] if self.aiDescription.strip() else ""
			if len(firstLine) > 100:
				firstLine = firstLine[:97] + "..."
		if self.isExternalWebImage:
			if self.isDownloaded:
				lines.append("Status: Obraz pobrany ze strony internetowej")
			else:
				webUrl = self.sourceUrl or self.hyperlinkUrl or self.cid
				dom = getCleanDomain(decodeSafeLinks(webUrl)) or self.linkServiceName or "internet"
				lines.append(f"Status: Obraz na serwerze WWW ({dom}) — zablokowany przez program Outlook")
				if self.sourceUrl:
					lines.append(f"Adres URL grafiki: {self.sourceUrl}")
				lines.append("Wskazówka: Aby odblokować i pobrać obrazy w tej wiadomości, przejdź do okna programu Outlook i naciśnij Ctrl+Shift+W, a następnie Enter.")
		elif self.sourceUrl:
			lines.append(f"Adres URL źródła grafiki: {self.sourceUrl}")
		if len(self.chatHistory) > 2:
			lines.append(f"Liczba zapytań w rozmowie z AI: {len(self.chatHistory) // 2}")
		return "\r\n".join(lines)


class AttachmentItem:
	"""Reprezentacja załącznika plikowego (PDF, DOCX, obrazy itp.) w wiadomości programu Outlook."""

	def __init__(
		self,
		index: int,
		fileName: str,
		sizeBytes: int,
		rawAttachment,
	):
		self.index = index
		self.fileName = (fileName or "").strip()
		self.sizeBytes = sizeBytes or 0
		self.rawAttachment = rawAttachment
		self.ext = os.path.splitext(self.fileName)[1].lower()

		# Rozpoznawanie typów plików
		self.isImage = self.ext in (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff")
		self.isPdf = self.ext == ".pdf"
		self.isDocx = self.ext == ".docx"
		self.isDoc = self.ext == ".doc"
		self.isText = self.ext in (".txt", ".csv", ".log", ".json", ".xml", ".ini")
		self.isSpreadsheet = self.ext in (".xlsx", ".xls")

		# Wymiary obrazu jeśli to grafika (uzupełniane przy pobieraniu)
		self.widthPx = 0
		self.heightPx = 0
		self.contextText = self.fileName
		self.hyperlinkUrl = ""
		self.hyperlinkTip = ""
		self.altText = ""
		self.title = self.fileName
		self.customName = ""
		self.isDivider = False
		self.linkCategory = ""
		self.linkServiceName = ""

		# Pamięć podręczna (cache) wyników oraz rozmowy
		self.ocrText: str = ""
		self.aiDescription: str = ""
		self.chatHistory: list[dict] = []
		self.base64Data: str = ""
		self.tempSavedPath: str = ""
		self.unlockedPdfPath: str = ""
		self.pdfPassword: str = ""

	@property
	def isScreenshotOrUnlabeled(self) -> bool:
		return False

	@property
	def typeName(self) -> str:
		if self.isPdf:
			return "Dokument PDF"
		if self.isDocx:
			return "Dokument programu Word (DOCX)"
		if self.isImage:
			return f"Plik graficzny ({self.ext.upper().lstrip('.')})"
		if self.isSpreadsheet:
			return f"Arkusz kalkulacyjny ({self.ext.upper().lstrip('.')})"
		if self.isText:
			return f"Plik tekstowy ({self.ext.upper().lstrip('.')})"
		return f"Plik ({self.ext.upper().lstrip('.') if self.ext else 'bez rozszerzenia'})"

	@property
	def formattedSize(self) -> str:
		if self.sizeBytes < 1024:
			return f"{self.sizeBytes} B"
		if self.sizeBytes < 1024 * 1024:
			return f"{self.sizeBytes / 1024:.1f} KB"
		return f"{self.sizeBytes / (1024 * 1024):.1f} MB"

	def getDisplayLabel(self, totalDocLength: int = 0) -> str:
		tag = f"[{self.ext.upper().lstrip('.')}]" if self.ext else "[PLIK]"
		parts = [f"{self.index}: {tag} {self.fileName} ({self.formattedSize})"]
		if self.ocrText:
			if self.ocrText == "(Brak rozpoznanego tekstu)":
				parts.append(f"[{_T('tag_text_empty')}]")
			else:
				snip = _cleanSnippet(self.ocrText)
				parts.append(f"[{_T('tag_text')}: \"{snip}\"]" if snip else f"[{_T('tag_ocr_ready')}]")
		if self.aiDescription:
			snip = _cleanSnippet(self.aiDescription)
			parts.append(f"[{_T('tag_ai')}: \"{snip}\"]" if snip else f"[{_T('tag_ai_ready')}]")
		return " — ".join(parts)

	def getDetailedInfo(self) -> str:
		lines = [
			f"Załącznik nr: {self.index}",
			f"Nazwa pliku: {self.fileName}",
			f"Typ elementu: {self.typeName}",
			f"Rozmiar pliku: {self.formattedSize} ({self.sizeBytes:,} bajtów)".replace(",", " "),
		]
		if self.ocrText:
			firstLine = self.ocrText.strip().splitlines()[0] if self.ocrText.strip() else ""
			if len(firstLine) > 100:
				firstLine = firstLine[:97] + "..."
			lines.append(f"Rozpoznany tekst OCR: {firstLine}")
		if self.aiDescription:
			firstLine = self.aiDescription.strip().splitlines()[0] if self.aiDescription.strip() else ""
			if len(firstLine) > 100:
				firstLine = firstLine[:97] + "..."
			lines.append(f"Opis / analiza AI: {firstLine}")
		if len(self.chatHistory) > 2:
			lines.append(f"Liczba zapytań w rozmowie z AI: {len(self.chatHistory) // 2}")
		return "\r\n".join(lines)


def safeOpenUrl(url: str) -> bool:
	"""Bezpiecznie otwiera adres URL w domyślnej przeglądarce, weryfikując protokół."""
	if not url:
		return False
	realUrl = decodeSafeLinks(url).strip()
	try:
		parsed = urllib.parse.urlparse(realUrl)
		if parsed.scheme.lower() not in ALLOWED_URL_SCHEMES:
			log.warning(f"outlookImageHelper: Zablokowano niebezpieczny protokół URL: {parsed.scheme} ({realUrl})")
			ui.message(f"Zablokowano nieobsługiwany lub potencjalnie niebezpieczny protokół: {parsed.scheme}")
			return False
		if parsed.scheme.lower() in ("mailto", "tel", "geo"):
			try:
				os.startfile(realUrl)
				return True
			except Exception:
				pass
		webbrowser.open(realUrl)
		return True
	except Exception as e:
		log.error(f"outlookImageHelper: Błąd otwierania linku {url}: {e}")
		ui.message(f"Nie udało się otworzyć linku: {e}")
		return False


def openInteractiveTarget(target: str, kind: str = "") -> bool:
	"""
	Otwiera interaktywny link (URL, e-mail, telefon, GEO) za pomocą domyślnej aplikacji w Windows:
	- URL -> Domyślna przeglądarka internetowa
	- e-mail -> Domyślny klient pocztowy (np. Outlook)
	- telefon -> Domyślna aplikacja połączeń / Telefon (Windows Phone Link)
	- GEO / współrzędne -> Domyślna aplikacja map lub Mapy Google w przeglądarce
	"""
	if not target:
		return False
	cleanTarget = target.strip()

	if kind == "email":
		if not cleanTarget.lower().startswith("mailto:"):
			cleanTarget = f"mailto:{cleanTarget}"
	elif kind == "phone":
		if not cleanTarget.lower().startswith("tel:"):
			telDigits = re.sub(r"[^\d+]", "", cleanTarget)
			cleanTarget = f"tel:{telDigits}"
	elif kind == "geo":
		if not (cleanTarget.lower().startswith("http://") or cleanTarget.lower().startswith("https://") or cleanTarget.lower().startswith("geo:")):
			cleanTarget = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(cleanTarget)}"
	elif kind == "url":
		if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", cleanTarget):
			cleanTarget = f"https://{cleanTarget}"

	# Zawsze walidujemy protokół PRZED wywołaniem jakiegokolwiek adresu w powłoce systemowej
	return safeOpenUrl(cleanTarget)


def safeOpenAttachment(item: AttachmentItem, parent=None) -> bool:
	"""
	Bezpiecznie zapisuje i uruchamia plik załącznika.
	Weryfikuje rozszerzenie zarówno oryginalnej nazwy, jak i rzeczywistego pliku uruchamianego.
	W przypadku potencjalnie niebezpiecznych rozszerzeń wykonywalnych (np. .exe, .bat, .vbs)
	wyświetla okno dialogowe z ostrzeżeniem przed uruchomieniem kodu z maila.
	"""
	if not item:
		return False

	tempPath = saveAttachmentToTempFile(item)
	if not tempPath or not os.path.isfile(tempPath):
		showAccessibleMessageBox(f"Nie udało się pobrać pliku załącznika '{item.fileName}' z serwera pocztowego.", title="Błąd załącznika", parent=parent)
		return False

	# Sprawdź rozszerzenie faktycznie uruchamianego pliku oraz oryginalną nazwę
	actualExt = os.path.splitext(tempPath)[1].lower()
	origExt = os.path.splitext((item.fileName or "").strip().rstrip(". "))[1].lower()
	isDangerous = (actualExt in DANGEROUS_ATTACHMENT_EXTENSIONS) or (origExt in DANGEROUS_ATTACHMENT_EXTENSIONS)

	if isDangerous:
		isPl = (getEffectiveLanguage() == "pl")
		warnTitle = "Ostrzeżenie o bezpieczeństwie" if isPl else "Security Warning"
		warnExt = actualExt if actualExt in DANGEROUS_ATTACHMENT_EXTENSIONS else origExt
		warnMsg = (
			f"Załącznik '{item.fileName}' jest plikiem wykonywalnym lub skryptem ({warnExt}).\n\n"
			f"Uruchomienie takiego pliku niesie wysokie ryzyko uruchomienia szkodliwego oprogramowania.\n\n"
			f"Czy na pewno chcesz go otworzyć?"
			if isPl else
			f"Attachment '{item.fileName}' is an executable file or script ({warnExt}).\n\n"
			f"Opening this file carries a high risk of executing malicious software.\n\n"
			f"Are you sure you want to run it?"
		)
		if not showAccessibleConfirmBox(warnMsg, warnTitle, parent=parent, defaultYes=False):
			try:
				os.remove(tempPath)
			except Exception:
				pass
			ui.message("Anulowano otwieranie potencjalnie niebezpiecznego pliku." if isPl else "Canceled opening potentially dangerous file.")
			return False

	try:
		os.startfile(tempPath)
		ui.message(f"Otwieram plik: {item.fileName}")
		return True
	except Exception as e:
		showAccessibleMessageBox(f"Nie udało się otworzyć załącznika '{item.fileName}':\n{e}", title="Błąd załącznika", parent=parent)
		return False


def safeOpenItem(item: GraphicItem | AttachmentItem, parent=None) -> bool:
	"""Wspólna, bezpieczna funkcja otwierająca plik załącznika lub hiperłącze grafiki."""
	if isinstance(item, AttachmentItem):
		return safeOpenAttachment(item, parent=parent)
	else:
		if not getattr(item, "hyperlinkUrl", ""):
			showAccessibleMessageBox("Wybrany element nie zawiera powiązanego hiperłącza (adresu WWW).", title="Brak hiperłącza", parent=parent)
			return False
		res = safeOpenUrl(item.hyperlinkUrl)
		if res:
			ui.message(f"Otwieram link w przeglądarce: {decodeSafeLinks(item.hyperlinkUrl)}")
		else:
			showAccessibleMessageBox(f"Nie udało się otworzyć adresu URL w domyślnej przeglądarce:\n{decodeSafeLinks(item.hyperlinkUrl)}", title="Błąd otwierania linku", parent=parent)
		return res


def sanitizeFilename(name: str, maxLen: int = 50) -> str:
	"""Oczyszcza ciąg tekstowy, tworząc bezpieczną nazwę pliku w systemie Windows."""
	if not name:
		return ""
	cleaned = _RE_CLEAN_FILENAME_INVALID.sub('', name)
	cleaned = _RE_CLEAN_FILENAME_WHITESPACE.sub(' ', cleaned)
	cleaned = _RE_CLEAN_FILENAME_DASHES.sub('_', cleaned)
	cleaned = cleaned.strip('._')
	if len(cleaned) > maxLen:
		cleaned = cleaned[:maxLen].rstrip('._')
	return cleaned


def generateDescriptiveFilename(item: GraphicItem, useAi: bool = False, apiKey: str = "", aiCandidate: str = "") -> str:
	"""Generuje czytelną, opisową nazwę pliku dla grafiki bez blokowania wątku głównego."""
	indexPrefix = f"{item.index:02d}"
	candidate = ""

	# 1. Własna nazwa nadana przez użytkownika
	if getattr(item, "customName", ""):
		candidate = item.customName

	# 2. Tekst alternatywny lub tytuł
	if not candidate:
		if item.altText:
			candidate = item.altText
		elif item.title:
			candidate = item.title

	# 3. Jeśli grafika zawiera rozpoznaną usługę lub link (np. LinkedIn, YouTube, Facebook)
	if not candidate and item.hyperlinkUrl:
		cat, svc = getLinkCategoryAndName(item.hyperlinkUrl)
		if svc:
			candidate = svc

	# 4. Jeśli OCR był wykonany, użyj słów kluczowych z pierwszej linii tekstu
	if not candidate and item.ocrText:
		firstLine = item.ocrText.strip().splitlines()[0] if item.ocrText.strip() else ""
		words = [w for w in re.split(r'[\s,.;:!?]+', firstLine) if len(w) >= 3]
		if words:
			candidate = "_".join(words[:4])

	# 5. Jeśli przekazano nazwę wygenerowaną asynchronicznie przez AI
	if not candidate and aiCandidate:
		candidate = aiCandidate.strip().strip('"').strip("'")

	# 6. Domyślny opis w oparciu o typ i rozmiar
	if not candidate:
		if getattr(item, "isDivider", False):
			candidate = f"Linia_podzialu_{item.widthPx}x{item.heightPx}"
		elif item.isScreenshotOrUnlabeled:
			candidate = f"Zrzut_{item.widthPx}x{item.heightPx}"
		else:
			candidate = f"Obraz_{item.widthPx}x{item.heightPx}"

	safeName = sanitizeFilename(candidate)
	if not safeName:
		safeName = f"Grafika_{item.index}"
	return f"{indexPrefix}_{safeName}.png"


def extractHyperlinkFromShape(rawObj, isInline: bool) -> tuple[str, str]:
	"""
	Wyodrębnia adres URL oraz opcjonalny ScreenTip z hiperłącza
	przypisanego do grafiki w edytorze WordMail.
	"""
	if not rawObj:
		return "", ""
	url = ""
	tip = ""

	# 1. Sprawdź właściwość .Hyperlink obiektu kształtu
	try:
		hl = getattr(rawObj, "Hyperlink", None)
		if hl:
			addr = getattr(hl, "Address", "") or ""
			subAddr = getattr(hl, "SubAddress", "") or ""
			tip = getattr(hl, "ScreenTip", "") or ""
			if addr and subAddr:
				url = f"{addr}#{subAddr}"
			elif addr:
				url = addr
			elif subAddr:
				url = subAddr
			if url:
				return url.strip(), tip.strip()
	except Exception as e:
		log.debug(f"outlookImageHelper: rawObj.Hyperlink check: {e}")

	# 2. Sprawdź kolekcję Hyperlinks w Range / Anchor
	try:
		rng = getattr(rawObj, "Range", None) if isInline else getattr(rawObj, "Anchor", None)
		if rng:
			hls = getattr(rng, "Hyperlinks", None)
			if hls and int(hls.Count) > 0:
				hl = hls.Item(1)
				addr = getattr(hl, "Address", "") or ""
				subAddr = getattr(hl, "SubAddress", "") or ""
				tip = getattr(hl, "ScreenTip", "") or ""
				if addr and subAddr:
					url = f"{addr}#{subAddr}"
				elif addr:
					url = addr
				elif subAddr:
					url = subAddr
				if url:
					return url.strip(), tip.strip()
	except Exception as e:
		log.debug(f"outlookImageHelper: rng.Hyperlinks check: {e}")

	return url.strip(), tip.strip()


def extractContextTextForShape(doc, rawObj, isInline: bool) -> str:
	"""
	Wyodrębnia tekst kontekstowy znajdujący się bezpośrednio nad grafiką
	(np. imię i nazwisko pracownika, nagłówek, etykietę).
	"""
	if not doc or not rawObj:
		return ""
	try:
		rng = getattr(rawObj, "Range", None) if isInline else getattr(rawObj, "Anchor", None)
		if not rng:
			return ""
		rngStart = int(rng.Start)
		if rngStart <= 0:
			return ""

		candidates = []

		# 1. Sprawdź akapit zawierający grafikę
		p = None
		try:
			p = rng.Paragraphs.Item(1)
		except Exception:
			try:
				p = rng.Paragraphs(1)
			except Exception:
				p = None

		if p:
			# A. Tekst w tym samym akapicie PRZED grafiką
			try:
				pStart = int(p.Range.Start)
				if pStart < rngStart:
					sameText = doc.Range(pStart, rngStart).Text or ""
					cleanSame = sameText.replace("\r", " ").replace("\n", " ").replace("\x01", "").strip()
					if cleanSame and len(cleanSame) >= 2:
						candidates.append(cleanSame)
			except Exception as e:
				log.debug(f"outlookImageHelper: sameText error: {e}")

			# B. Poprzednie akapity (do 5 wstecz)
			currP = p
			steps = 0
			while currP and steps < 6 and len(candidates) < 4:
				steps += 1
				try:
					prevP = currP.Previous()
				except Exception:
					break
				if not prevP:
					break
				currP = prevP
				try:
					pTxt = prevP.Range.Text or ""
					pTxt = pTxt.replace("\r", "\n").replace("\x01", "").strip()
					lines = [l.strip() for l in pTxt.split("\n") if l.strip()]
					for line in reversed(lines):
						if line and line not in candidates:
							candidates.append(line)
				except Exception:
					break

		# 2. Rezerwa: bezpośredni zakres doc.Range do 400 znaków wstecz
		if not candidates:
			try:
				lookbackStart = max(0, rngStart - 400)
				rawPre = doc.Range(lookbackStart, rngStart).Text or ""
				rawPre = rawPre.replace("\r", "\n").replace("\x01", "").strip()
				lines = [l.strip() for l in rawPre.split("\n") if l.strip()]
				for line in reversed(lines):
					if line and line not in candidates:
						candidates.append(line)
			except Exception as e:
				log.debug(f"outlookImageHelper: rawPre error: {e}")

		# 3. Wybór najlepszego kandydata
		for cand in candidates:
			clean = cand.strip(" \t\r\n:=-_")
			if clean and len(clean) >= 2:
				if len(clean) > 80:
					clean = clean[:77] + "..."
				return clean

	except Exception as ex:
		log.debug(f"outlookImageHelper: extractContextTextForShape error: {ex}")
	return ""


def _getWordDocumentFromHwnd(hwnd: int):
	"""Pobiera natywny obiekt Document programu Word na podstawie uchwytu okna _WwG."""
	if not hwnd or not winUser.isWindow(hwnd):
		return None
	try:
		import comtypes.client
		from comtypes.automation import IDispatch
		from comtypes import POINTER
		from ctypes import byref, oledll

		pDispatch = POINTER(IDispatch)()
		hr = oledll.oleacc.AccessibleObjectFromWindow(
			hwnd,
			OBJID_NATIVEOM,
			byref(IDispatch._iid_),
			byref(pDispatch),
		)
		if hr == 0 and pDispatch:
			try:
				disp = comtypes.client.dynamic.Dispatch(pDispatch)
			except Exception:
				disp = comtypes.client.wrap(pDispatch)
			return getattr(disp, "Document", None)
	except Exception as e:
		log.debug(f"outlookImageHelper: _getWordDocumentFromHwnd error: {e}")
	return None


def _findWordGWindow(parentHwnd: int) -> int:
	"""Przeszukuje okno w poszukiwaniu potomka o klasie _WwG (edytor WordMail / Reading Pane)."""
	found = [0]

	def enumChild(hwnd, lParam):
		cls = winUser.getClassName(hwnd)
		if cls == "_WwG" and winUser.isWindowVisible(hwnd):
			found[0] = hwnd
			return False
		return True

	try:
		from ctypes import wintypes
		WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
		ctypes.windll.user32.EnumChildWindows(parentHwnd, WNDENUMPROC(enumChild), 0)
	except Exception as ex:
		log.debug(f"outlookImageHelper: _findWordGWindow EnumChildWindows error: {ex}")
	return found[0]


def isMessageListRow(obj) -> bool:
	"""Sprawdza, czy obiekt NVDA reprezentuje wiersz/element wiadomości na liście w Outlooku."""
	if not obj:
		return False
	try:
		cachedCls = getattr(getattr(obj, "UIAElement", None), "cachedClassName", "") or getattr(obj, "windowClassName", "") or ""
		if cachedCls in ("LeafRow", "ThreadItem", "ThreadHeader", "TableItem", "UIAGridRow"):
			return True
		role = getattr(obj, "role", None)
		if role in (controlTypes.Role.LISTITEM, controlTypes.Role.TABLECELL, controlTypes.Role.TABLEROW):
			return True
	except Exception:
		pass
	return False


def isGroupHeader(obj) -> bool:
	"""Sprawdza, czy obiekt NVDA reprezentuje nagłówek grupy (np. 'Dzisiaj') na liście wiadomości w Outlooku."""
	if not obj:
		return False
	try:
		cachedCls = getattr(getattr(obj, "UIAElement", None), "cachedClassName", "") or getattr(obj, "windowClassName", "") or ""
		if cachedCls in ("LeafRow", "ThreadItem", "ThreadHeader", "TableItem", "UIAGridRow"):
			return False
		if cachedCls in ("GroupHeader", "OutlookGridGroupHeader"):
			return True
		role = getattr(obj, "role", None)
		if role == controlTypes.Role.LISTITEM:
			return False
		if role in (controlTypes.Role.GROUPING, controlTypes.Role.HEADER):
			rawName = getattr(getattr(obj, "UIAElement", None), "cachedName", "") or getattr(getattr(obj, "UIAElement", None), "currentName", "") or getattr(obj, "name", "") or ""
			if rawName and any(k in rawName.lower() for k in ("dzisiaj", "wczoraj", "today", "yesterday", "ten tydzień", "this week", "starsze", "older", "grupuj według", "group by", "kategorie:", "kategoria:")):
				return True
			if cachedCls in ("GroupHeader", "OutlookGridGroupHeader"):
				return True
			return False
	except Exception:
		pass
	return False


def isMessageBodyFocused(focus=None) -> bool:
	"""Zwraca True, gdy kursor/fokus NVDA znajduje się bezpośrednio w edytorze lub treści otwartej wiadomości."""
	if not focus:
		focus = api.getFocusObject()
	if not focus:
		return False
	try:
		if getattr(focus, "WinwordDocumentObject", None) is not None:
			return True
		ti = getattr(focus, "treeInterceptor", None)
		if ti and getattr(getattr(ti, "rootNVDAObject", None), "WinwordDocumentObject", None) is not None:
			return True
		hwnd = getattr(focus, "windowHandle", 0)
		if hwnd and winUser.getClassName(hwnd) == "_WwG":
			return True
	except Exception:
		pass
	return False


def isMessageListFocused(focus=None) -> bool:
	"""Zwraca True, gdy fokus NVDA znajduje się na liście wiadomości (Explorer) w Outlooku, a nie w treści."""
	if not focus:
		focus = api.getFocusObject()
	if not focus:
		return False
	if isMessageBodyFocused(focus):
		return False
	if isGroupHeader(focus):
		return True
	try:
		cls = (getattr(focus, "windowClassName", "") or "").lower()
		cachedCls = getattr(getattr(focus, "UIAElement", None), "cachedClassName", "") or ""
		if any(k in cls for k in ("supergrid", "outlookgrid", "netui", "tableview")) or cachedCls in ("LeafRow", "ThreadItem", "ThreadHeader", "TableItem", "UIAGridRow", "GroupHeader"):
			return True
		role = getattr(focus, "role", None)
		if role in (controlTypes.Role.LISTITEM, controlTypes.Role.TABLECELL, ROLE_TABLEROW, controlTypes.Role.LIST, controlTypes.Role.DATAITEM):
			return True
		app = getOutlookApp()
		if app:
			actWin = app.ActiveWindow()
			actExp = getattr(app, "ActiveExplorer", lambda: None)()
			if actWin and actExp and actWin == actExp:
				return True
	except Exception:
		pass
	return False


def getActiveWordDocument(mailItem=None):
	"""
	Zwraca aktywny obiekt Word.Document powiązany z bieżącą wiadomością Outlooka.
	Obsługuje aktywne okno wiadomości (Inspector) oraz okienko podglądu (Reading Pane).
	Zwraca None, jeśli wiadomość nie jest otwarta w osobnym oknie ani w aktywnym okienku podglądu.
	"""
	focus = api.getFocusObject()
	if not focus:
		return None

	# Sprawdź w pierwszej kolejności, czy aktywną aplikacją jest rzeczywiście Outlook
	# (aby nie przechwytywać dokumentów otwartych w programie Microsoft Word WINWORD.EXE!)
	appMod = getattr(focus, "appModule", None)
	isOutlookApp = bool(appMod and getattr(appMod, "appName", "").lower() == "outlook")
	if not isOutlookApp:
		fgHwnd = winUser.getForegroundWindow()
		if fgHwnd and "renwnd32" in (winUser.getClassName(fgHwnd) or "").lower():
			isOutlookApp = True
	if not isOutlookApp:
		return None

	# Jeśli fokus jest na nagłówku grupy -> brak dokumentu maila
	if isGroupHeader(focus):
		return None

	# 1. Bezpośredni obiekt WordDocument na fokusie lub treeInterceptorze
	doc = getattr(focus, "WinwordDocumentObject", None)
	if doc:
		return doc

	ti = getattr(focus, "treeInterceptor", None)
	if ti:
		root = getattr(ti, "rootNVDAObject", None)
		doc = getattr(root, "WinwordDocumentObject", None)
		if doc:
			return doc

	hwnd = getattr(focus, "windowHandle", 0)
	if hwnd and winUser.getClassName(hwnd) == "_WwG":
		doc = _getWordDocumentFromHwnd(hwnd)
		if doc:
			return doc

	nativeOm = getattr(appMod, "nativeOm", None) or getOutlookApp()
	if not nativeOm:
		return None

	try:
		activeWin = nativeOm.ActiveWindow()
		if not activeWin:
			activeWin = getattr(nativeOm, "ActiveExplorer", lambda: None)()
		if activeWin:
			# Jeśli aktywne okno to Inspector (otwarta w osobnym oknie wiadomość)
			if hasattr(activeWin, "CurrentItem"):
				doc = getattr(activeWin, "WordEditor", None)
				if doc:
					return doc

			# Jeśli aktywne okno to Explorer:
			# Pobieramy WordEditor z okienka podglądu (Reading Pane) TYLKO wtedy,
			# gdy okienko podglądu _WwG rzeczywiście istnieje w oknie głównym i jest widoczne.
			if hasattr(activeWin, "Selection") or hasattr(activeWin, "CurrentFolder"):
				fgHwnd = winUser.getForegroundWindow()
				if fgHwnd:
					wHwnd = _findWordGWindow(fgHwnd)
					if wHwnd and winUser.isWindowVisible(wHwnd):
						doc = _getWordDocumentFromHwnd(wHwnd)
						if doc:
							return doc
	except Exception as ex:
		log.debug(f"outlookImageHelper: getActiveWordDocument błąd: {ex}")

	return None


def getCurrentMailItem(doc=None):
	"""
	Pobiera obiekt MailItem ściśle powiązany z bieżącym fokusem użytkownika w Outlooku:
	1. Jeśli przekazano doc (Word.Document) -> sprawdza doc.MailEnvelope.Item lub doc.Parent.
	2. Jeśli fokus jest na nagłówku grupy -> zwraca None.
	3. Sprawdza obiekt doc na bieżącym fokusie / treeInterceptorze / oknie _WwG -> doc.MailEnvelope.Item.
	4. Jeśli aktywne okno to Inspector (otwarta wiadomość) -> zwraca CurrentItem.
	5. Jeśli aktywne okno to Explorer:
	   - sprawdza ActiveInlineResponse,
	   - sprawdza AttachmentSelection (gdy kursor/kliknięcie jest na kafelku załącznika w Reading Pane),
	   - sprawdza Selection (wiersz listy lub treść),
	   - fallback: ActiveExplorer.Selection / AttachmentSelection / ActiveInlineResponse,
	   - fallback: ActiveInspector.CurrentItem,
	   - fallback: wyszukanie widocznego okna _WwG w oknie pierwszoplanowym i pobranie MailEnvelope.Item.
	"""
	if doc:
		try:
			mailEnv = getattr(doc, "MailEnvelope", None)
			if mailEnv:
				it = getattr(mailEnv, "Item", None)
				if it and (hasattr(it, "Subject") or hasattr(it, "Attachments")):
					return it
		except Exception:
			pass
		try:
			p = getattr(doc, "Parent", None)
			if p and (hasattr(p, "Subject") or hasattr(p, "Attachments")):
				return p
			elif p and hasattr(p, "CurrentItem"):
				cur = getattr(p, "CurrentItem", None)
				if cur:
					return cur
		except Exception:
			pass

	focus = api.getFocusObject()
	if not focus:
		return None

	if isGroupHeader(focus):
		return None

	# 1. Bezpośredni obiekt WordDocument na fokusie lub treeInterceptorze (Reading Pane / Inspector)
	docObj = getattr(focus, "WinwordDocumentObject", None)
	if not docObj:
		ti = getattr(focus, "treeInterceptor", None)
		if ti:
			root = getattr(ti, "rootNVDAObject", None)
			docObj = getattr(root, "WinwordDocumentObject", None)
	if not docObj:
		hwnd = getattr(focus, "windowHandle", 0)
		if hwnd and winUser.getClassName(hwnd) == "_WwG":
			docObj = _getWordDocumentFromHwnd(hwnd)
	if docObj:
		try:
			mailEnv = getattr(docObj, "MailEnvelope", None)
			if mailEnv:
				it = getattr(mailEnv, "Item", None)
				if it and (hasattr(it, "Subject") or hasattr(it, "Attachments")):
					return it
		except Exception:
			pass
		try:
			p = getattr(docObj, "Parent", None)
			if p and (hasattr(p, "Subject") or hasattr(p, "Attachments")):
				return p
			elif p and hasattr(p, "CurrentItem"):
				cur = getattr(p, "CurrentItem", None)
				if cur:
					return cur
		except Exception:
			pass

	appMod = getattr(focus, "appModule", None)
	nativeOm = getattr(appMod, "nativeOm", None) or getOutlookApp()
	if not nativeOm:
		return None

	try:
		activeWin = None
		try:
			activeWin = nativeOm.ActiveWindow()
		except Exception:
			pass
		if not activeWin:
			try:
				activeWin = getattr(nativeOm, "ActiveExplorer", lambda: None)()
			except Exception:
				pass

		if activeWin:
			# Jeśli okno to Inspector (otwarta w osobnym oknie wiadomość)
			if hasattr(activeWin, "CurrentItem"):
				cur = getattr(activeWin, "CurrentItem", None)
				if cur:
					return cur

			# Sprawdź ActiveInlineResponse w Explorerze
			try:
				inlineResp = getattr(activeWin, "ActiveInlineResponse", None)
				if inlineResp:
					return inlineResp
			except Exception:
				pass

			# Sprawdź zaznaczony załącznik w AttachmentSelection (np. kliknięty lub zaznaczony kafelek załącznika w Reading Pane)
			try:
				attSel = getattr(activeWin, "AttachmentSelection", None)
				if attSel and int(getattr(attSel, "Count", 0)) > 0:
					att = attSel.Item(1)
					p = getattr(att, "Parent", None)
					if p and (hasattr(p, "Subject") or hasattr(p, "Attachments")):
						return p
			except Exception:
				pass

			# Sprawdź Selection w Explorerze
			try:
				sel = getattr(activeWin, "Selection", None)
				if sel and int(getattr(sel, "Count", 0)) > 0:
					it = sel.Item(1)
					if hasattr(it, "Subject") or hasattr(it, "Attachments"):
						return it
					elif hasattr(it, "Parent"):
						p = getattr(it, "Parent", None)
						if p and (hasattr(p, "Subject") or hasattr(p, "Attachments")):
							return p
			except Exception:
				pass

		# Fallback na ActiveExplorer jeśli activeWin był czymś innym
		try:
			actExp = getattr(nativeOm, "ActiveExplorer", lambda: None)()
			if actExp and actExp != activeWin:
				try:
					inlineResp = getattr(actExp, "ActiveInlineResponse", None)
					if inlineResp:
						return inlineResp
				except Exception:
					pass
				try:
					attSel = getattr(actExp, "AttachmentSelection", None)
					if attSel and int(getattr(attSel, "Count", 0)) > 0:
						att = attSel.Item(1)
						p = getattr(att, "Parent", None)
						if p and (hasattr(p, "Subject") or hasattr(p, "Attachments")):
							return p
				except Exception:
					pass
				try:
					sel = getattr(actExp, "Selection", None)
					if sel and int(getattr(sel, "Count", 0)) > 0:
						it = sel.Item(1)
						if hasattr(it, "Subject") or hasattr(it, "Attachments"):
							return it
				except Exception:
					pass
		except Exception:
			pass

		# Fallback na ActiveInspector
		try:
			actInsp = getattr(nativeOm, "ActiveInspector", lambda: None)()
			if actInsp:
				cur = getattr(actInsp, "CurrentItem", None)
				if cur:
					return cur
		except Exception:
			pass

		# Fallback: okienko podglądu _WwG w oknie pierwszoplanowym
		try:
			fgHwnd = winUser.getForegroundWindow()
			if fgHwnd:
				wHwnd = _findWordGWindow(fgHwnd)
				if wHwnd and winUser.isWindowVisible(wHwnd):
					docFromHwnd = _getWordDocumentFromHwnd(wHwnd)
					if docFromHwnd:
						mailEnv = getattr(docFromHwnd, "MailEnvelope", None)
						if mailEnv:
							it = getattr(mailEnv, "Item", None)
							if it and (hasattr(it, "Subject") or hasattr(it, "Attachments")):
								return it
		except Exception:
			pass

	except Exception as ex:
		log.debug(f"outlookImageHelper: getCurrentMailItem ogólny błąd: {ex}")

	return None


def getSelectedMessagesInfo() -> tuple[int, object | None, bool]:
	"""
	Zwraca informację o zaznaczonych wiadomościach w programie Outlook:
	(liczba zaznaczonych elementów, obiekt Selection, czy zaznaczono wiele wiadomości).
	Jeśli aktywna jest otwarta w osobnym oknie wiadomość (Inspector) -> zwraca (1, None, False).
	"""
	focus = api.getFocusObject()
	if not focus:
		return 0, None, False

	# Jeśli fokus jest w edytorze treści otwartej wiadomości (WordMail), dotyczy pojedynczej wiadomości
	if isMessageBodyFocused(focus):
		return 1, None, False

	appMod = getattr(focus, "appModule", None)
	nativeOm = getattr(appMod, "nativeOm", None) or getOutlookApp()

	sel = None
	cnt = 0

	if nativeOm:
		try:
			# Najpierw sprawdź, czy aktywne okno to Inspector (osobne okno wiadomości)
			activeWin = None
			for attr in ("ActiveInspector", "activeInspector", "ActiveWindow", "activeWindow"):
				fn = getattr(nativeOm, attr, None)
				if fn:
					try:
						res = fn() if callable(fn) else fn
						if res:
							if "inspector" in attr.lower():
								activeWin = res
								break
							clsVal = getattr(res, "Class", None)
							clsNum = clsVal() if callable(clsVal) else clsVal
							if clsNum == 35 or hasattr(res, "CurrentItem") or hasattr(res, "currentItem"):
								activeWin = res
								break
					except Exception:
						pass

			if activeWin is not None:
				curItem = getattr(activeWin, "CurrentItem", None) or getattr(activeWin, "currentItem", None)
				itemObj = curItem() if callable(curItem) else curItem
				return 1, itemObj, False

			exp = None
			for attr in ("activeExplorer", "ActiveExplorer"):
				fn = getattr(nativeOm, attr, None)
				if fn:
					try:
						exp = fn() if callable(fn) else fn
						if exp:
							break
					except Exception:
						pass

			if exp:
				for attr in ("selection", "Selection"):
					s = getattr(exp, attr, None)
					if s is not None:
						if hasattr(s, "Count") or hasattr(s, "count"):
							sel = s
						elif callable(s):
							try:
								sel = s()
							except Exception:
								sel = s
						else:
							sel = s
						if sel is not None:
							break

				if sel is not None:
					for attr in ("count", "Count"):
						c = getattr(sel, attr, None)
						if c is not None:
							try:
								cnt = int(c() if callable(c) and not isinstance(c, (int, float)) else c)
								break
							except Exception:
								pass
					if cnt == 0:
						try:
							cnt = len(sel)
						except Exception:
							pass
		except Exception as ex:
			log.debug(f"outlookImageHelper: getSelectedMessagesInfo explorer selection error: {ex}")

	# Fallback: UIA Selection Pattern na liście / tabeli wiadomości
	if cnt <= 1 and focus:
		try:
			parentList = focus
			for _ in range(3):
				if not parentList:
					break
				role = getattr(parentList, "role", None)
				if role in (controlTypes.Role.LIST, controlTypes.Role.TABLE, controlTypes.Role.TREEVIEW):
					break
				parentList = getattr(parentList, "parent", None)
			if parentList and hasattr(parentList, "UIASelectionPattern"):
				selItems = parentList.UIASelectionPattern.GetSelection()
				if selItems and len(selItems) > 1:
					cnt = len(selItems)
		except Exception:
			pass

	if cnt > 1:
		return cnt, sel, True
	elif cnt == 1:
		return 1, sel, False

	return 0, None, False


def handleMultipleMessagesSelected(count: int, sel, action: str = "list") -> bool:
	"""
	Obsługuje sytuację, gdy użytkownik wywołuje funkcję dodatku mając zaznaczonych
	kilka wiadomości naraz w programie Outlook.
	"""
	attMsgs = 0
	if sel:
		try:
			limit = min(count, 50)
			for i in range(1, limit + 1):
				it = None
				for itName in ("item", "Item"):
					itFn = getattr(sel, itName, None)
					if itFn:
						try:
							it = itFn(i)
							break
						except Exception:
							pass
				if not it and sel:
					try:
						it = sel[i - 1]
					except Exception:
						pass

				if it:
					atts = getattr(it, "attachments", getattr(it, "Attachments", None))
					attCount = 0
					if atts:
						for cName in ("count", "Count"):
							cVal = getattr(atts, cName, None)
							if cVal is not None:
								try:
									attCount = int(cVal() if callable(cVal) else cVal)
									break
								except Exception:
									pass
					if attCount > 0:
						attMsgs += 1
		except Exception as ex:
			log.debug(f"outlookImageHelper: handleMultipleMessagesSelected error: {ex}")

	playAlertSound("image_alert", volume=70, context="Powiadomienie o zaznaczeniu wielu wiadomości")
	if action == "list":
		if attMsgs > 0:
			ui.message(_T("msg_multiple_selected_att", count=count, attCount=attMsgs))
		else:
			ui.message(_T("msg_multiple_selected", count=count))
	else:
		ui.message(_T("msg_multiple_selected_action", count=count))
	return True


def handleReadingPaneOff(mailItem, action: str = "list") -> bool:
	"""
	Obsługuje sytuację, gdy użytkownik wywołuje funkcję dodatku na liście wiadomości,
	ale wiadomość nie jest otwarta w osobnym oknie (Inspector), a okienko podglądu (Reading Pane)
	jest wyłączone lub nie załadowało dokumentu WordEditor.
	Zwraca True informując, że zdarzenie zostało obsłużone i zakomunikowane użytkownikowi.
	"""
	subj = ""
	attCount = 0
	if mailItem:
		try:
			subj = str(getattr(mailItem, "Subject", "") or "").strip()
		except Exception:
			pass
		try:
			realAtts = extractAttachmentsFromCurrentMessage(mailItem=mailItem, includeInline=False)
			attCount = len(realAtts)
		except Exception:
			pass

	playAlertSound("image_alert", volume=70, context="Powiadomienie o wyłączonym podglądzie wiadomości")

	if action == "list":
		if attCount > 0:
			if subj:
				ui.message(_T("msg_reading_pane_off_att_subj", subj=subj, count=attCount))
			else:
				ui.message(_T("msg_reading_pane_off_att", count=attCount))
		else:
			if subj:
				ui.message(_T("msg_reading_pane_off_subj", subj=subj))
			else:
				ui.message(_T("msg_reading_pane_off"))
	elif action == "ocr":
		if subj:
			ui.message(_T("msg_reading_pane_off_ocr_subj", subj=subj))
		else:
			ui.message(_T("msg_reading_pane_off_ocr"))
	elif action == "ai":
		if subj:
			ui.message(_T("msg_reading_pane_off_ai_subj", subj=subj))
		else:
			ui.message(_T("msg_reading_pane_off_ai"))
	elif action == "invert":
		if subj:
			ui.message(_T("msg_reading_pane_off_invert_subj", subj=subj))
		else:
			ui.message(_T("msg_reading_pane_off_invert"))
	return True


def getMessageKey(mailItem=None) -> str:
	"""Zwraca unikalny klucz identyfikujący bieżącą wiadomość e-mail (EntryID lub Subject+Time)."""
	if not mailItem:
		mailItem = getCurrentMailItem()
	if mailItem:
		try:
			entryID = getattr(mailItem, "EntryID", "") or ""
			if entryID:
				return str(entryID).strip()
		except Exception:
			pass
		try:
			subj = str(getattr(mailItem, "Subject", "") or "").strip()
			rec = str(getattr(mailItem, "ReceivedTime", "") or "").strip()
			if subj or rec:
				return f"{subj}_{rec}"
		except Exception:
			pass
	return "current_message"


def _log_sound_event(sound: str, volume: int | str, context: str, status: str):
	"""Rejestruje zdarzenie audio w buforze kołowym (maks. 30 wpisów) dla celów diagnostyki."""
	global _sound_events_history
	entry = {
		"time": time.strftime("%Y-%m-%d %H:%M:%S"),
		"sound": sound,
		"volume": f"{volume}%" if isinstance(volume, int) else str(volume),
		"context": context,
		"status": status,
	}
	_sound_events_history.append(entry)
	if len(_sound_events_history) > 30:
		_sound_events_history.pop(0)


def _get_scaled_sound_path(sound_name_or_path: str, volume: int | None = None) -> str | None:
	"""
	Zwraca ścieżkę do pliku .wav przeskalowanego do podanej głośności (0-100%).
	Dla głośności 100 zwraca ścieżkę oryginalną.
	Dla głośności <= 0 zwraca None (wyciszenie).
	Przeskalowane pliki są buforowane w katalogu SOUNDS_CACHE_DIR dla zerowej latencji.
	"""
	if volume is None:
		try:
			volume = int(config.conf.get(CONFIG_SECTION, {}).get("soundVolume", 80))
		except Exception:
			volume = 80
	else:
		try:
			volume = int(volume)
		except (ValueError, TypeError):
			volume = 80

	if volume <= 0:
		return None

	theme = "associative"
	try:
		theme = str(config.conf.get(CONFIG_SECTION, {}).get("soundTheme", "associative") or "associative")
	except Exception:
		theme = "associative"

	orig_path = None
	if os.path.isabs(sound_name_or_path) and os.path.isfile(sound_name_or_path):
		orig_path = sound_name_or_path
		sound_base = os.path.splitext(os.path.basename(sound_name_or_path))[0]
	else:
		sound_base = sound_name_or_path.replace(".wav", "")
		theme_cand = os.path.join(SOUNDS_DIR, theme, f"{sound_base}.wav")
		if os.path.isfile(theme_cand):
			orig_path = theme_cand
		else:
			base_cand = os.path.join(SOUNDS_DIR, f"{sound_base}.wav")
			if os.path.isfile(base_cand):
				orig_path = base_cand
			elif os.path.isfile(ALERT_SOUND_PATH):
				orig_path = ALERT_SOUND_PATH

	if not orig_path or not os.path.isfile(orig_path):
		return None

	if volume >= 100:
		return orig_path

	cached_path = os.path.join(SOUNDS_CACHE_DIR, f"{theme}_{sound_base}_vol{volume}.wav")
	if os.path.isfile(cached_path) and os.path.getsize(cached_path) > 44:
		return cached_path

	try:
		os.makedirs(SOUNDS_CACHE_DIR, exist_ok=True)
		factor = volume / 100.0
		with wave.open(orig_path, "rb") as in_wf:
			nchannels = in_wf.getnchannels()
			sampwidth = in_wf.getsampwidth()
			framerate = in_wf.getframerate()
			nframes = in_wf.getnframes()
			raw_frames = in_wf.readframes(nframes)

		if sampwidth == 2:
			samples = array.array("h", raw_frames)
			scaled = array.array("h", (max(-32768, min(32767, int(s * factor))) for s in samples))
			scaled_bytes = scaled.tobytes()
		elif sampwidth == 1:
			samples = array.array("B", raw_frames)
			scaled = array.array("B", (max(0, min(255, int(128 + (b - 128) * factor))) for b in samples))
			scaled_bytes = scaled.tobytes()
		else:
			return orig_path

		tmp_path = cached_path + f".tmp_{os.getpid()}_{int(time.time() * 1000)}"
		with wave.open(tmp_path, "wb") as out_wf:
			out_wf.setnchannels(nchannels)
			out_wf.setsampwidth(sampwidth)
			out_wf.setframerate(framerate)
			out_wf.writeframes(scaled_bytes)

		os.replace(tmp_path, cached_path)
		return cached_path
	except Exception as e:
		log.debug(f"outlookImageHelper: Nie udało się przeskalować dźwięku {sound_name_or_path} do głośności {volume}: {e}")
		return orig_path


def playAlertSound(soundName: str = "image_alert", volume: int | None = None, context: str = "", force: bool = False):
	"""Odtwarza dźwięk alertu lub nawigacji z uwzględnieniem skalowania głośności oraz bufora zdarzeń."""
	sec = config.conf.get(CONFIG_SECTION, {})
	if not force and not _to_bool(sec.get("playSound", True), default=True):
		_log_sound_event(soundName, 0, context, "Zablokowano (opcja odtwarzania dźwięków wyłączona)")
		return

	if volume is None:
		try:
			volume = int(sec.get("soundVolume", 80))
		except Exception:
			volume = 80

	scaled_path = _get_scaled_sound_path(soundName, volume=volume)
	if scaled_path and os.path.isfile(scaled_path):
		try:
			nvwave.playWaveFile(scaled_path, asynchronous=True)
			_log_sound_event(soundName, volume, context, "Odtworzono pomyślnie")
		except Exception as e:
			_log_sound_event(soundName, volume, context, f"Błąd odtwarzania: {e}")
	else:
		_log_sound_event(soundName, volume, context, "Wyciszono (głośność 0% lub brak pliku)")


def isSpeechFilterInstalled() -> bool:
	"""Sprawdza, czy dodatek Outlook Speech Filter jest aktualnie załadowany i aktywny w NVDA."""
	try:
		import globalPluginHandler
		for p in globalPluginHandler.runningPlugins:
			if "outlookSpeechFilter" in getattr(p, "__module__", "") or "outlookSpeechFilter" in getattr(getattr(p, "__class__", None), "__module__", ""):
				return True
	except Exception:
		pass
	return False


def isMessageBodyFocused(focus=None) -> bool:
	"""Zwraca True, gdy kursor/fokus NVDA znajduje się bezpośrednio w edytorze lub treści otwartej wiadomości."""
	if not focus:
		focus = api.getFocusObject()
	if not focus:
		return False
	try:
		if getattr(focus, "WinwordDocumentObject", None) is not None:
			return True
		ti = getattr(focus, "treeInterceptor", None)
		if ti and getattr(getattr(ti, "rootNVDAObject", None), "WinwordDocumentObject", None) is not None:
			return True
		hwnd = getattr(focus, "windowHandle", 0)
		if hwnd and winUser.getClassName(hwnd) == "_WwG":
			return True
	except Exception:
		pass
	return False


def isMessageListFocused(focus=None) -> bool:
	"""Zwraca True, gdy fokus NVDA znajduje się na liście wiadomości (Explorer) w Outlooku, a nie w treści."""
	if not focus:
		focus = api.getFocusObject()
	if not focus:
		return False
	if isMessageBodyFocused(focus):
		return False
	if isGroupHeader(focus):
		return True
	try:
		cls = (getattr(focus, "windowClassName", "") or "").lower()
		cachedCls = getattr(getattr(focus, "UIAElement", None), "cachedClassName", "") or ""
		if any(k in cls for k in ("supergrid", "outlookgrid", "netui", "tableview")) or cachedCls in ("LeafRow", "ThreadItem", "ThreadHeader", "TableItem", "UIAGridRow", "GroupHeader"):
			return True
		role = getattr(focus, "role", None)
		if role in (controlTypes.Role.LISTITEM, controlTypes.Role.TABLECELL, ROLE_TABLEROW, controlTypes.Role.LIST, controlTypes.Role.DATAITEM):
			return True
		app = getOutlookApp()
		if app:
			actWin = app.ActiveWindow()
			actExp = getattr(app, "ActiveExplorer", lambda: None)()
			if actWin and actExp and actWin == actExp:
				return True
	except Exception:
		pass
	return False


def isSpeechFilterHandlingInlineImages() -> bool:
	"""
	Sprawdza, czy dodatek Outlook Speech Filter jest aktywny i obsługuje
	kategorię dźwięków obecności obrazów/grafik na liście wiadomości.
	"""
	try:
		if not isSpeechFilterInstalled():
			return False
		osf_conf = config.conf.get("outlookSpeechFilter", {})
		if not osf_conf:
			return False
		use_sounds = _to_bool(osf_conf.get("useSounds", True), default=True)
		sound_inline = _to_bool(osf_conf.get("soundInlineImages", True), default=True)
		return bool(use_sounds and sound_inline)
	except Exception as e:
		log.debug(f"outlookImageHelper: błąd sprawdzania isSpeechFilterHandlingInlineImages: {e}")
	return False


def mailItemHasInlineImages(mailItem) -> bool:
	"""
	Błyskawicznie sprawdza, czy dany obiekt MailItem zawiera wklejone obrazy lub zrzuty ekranu.
	Używa wyłącznie bezpiecznych, niestrzeżonych właściwości obiektu Attachment (Type, FileName),
	nigdy nie odwołując się do PropertyAccessor ani HTMLBody, co chroni przed alertami bezpieczeństwa Outlook.
	Wyniki są buforowane w pamięci podręcznej według EntryID.
	"""
	if not mailItem:
		return False
	entryID = None
	try:
		entryID = getattr(mailItem, "EntryID", None)
		if entryID and entryID in _mail_inline_img_cache:
			return _mail_inline_img_cache[entryID]
	except Exception:
		entryID = None

	has_img = False
	try:
		attColl = getattr(mailItem, "Attachments", None)
		count = int(getattr(attColl, "Count", 0)) if attColl else 0
		if count > 0:
			for i in range(1, count + 1):
				try:
					att = attColl.Item(i)
					attType = int(getattr(att, "Type", 1) or 1)
					fn = (getattr(att, "FileName", "") or "").lower()
					# Typ 5=olEmbeddeditem, 6=olOLE to grafiki osadzone w treści
					if attType in (5, 6) or fn.startswith("image") or fn.startswith("oledata") or not fn:
						has_img = True
						break
				except Exception:
					pass
	except Exception as e:
		log.debug(f"outlookImageHelper: mailItemHasInlineImages błąd: {e}")

	if entryID:
		if len(_mail_inline_img_cache) > 200:
			_mail_inline_img_cache.clear()
		_mail_inline_img_cache[entryID] = has_img
	return has_img


def _cleanup_old_reports(dir_path: str, addon_prefix: str = "OutlookImageHelper", max_count: int = 20):
	"""Usuwa najstarsze raporty diagnostyczne (FIFO) należące wyłącznie do danego dodatku."""
	try:
		if not os.path.isdir(dir_path):
			return
		target_prefix = f"Raport_{addon_prefix}_"
		files = [
			os.path.join(dir_path, f)
			for f in os.listdir(dir_path)
			if f.startswith(target_prefix) and f.endswith(".txt")
		]
		if len(files) > max_count:
			files.sort(key=lambda p: os.path.getmtime(p))
			to_delete = files[:len(files) - max_count]
			for old_file in to_delete:
				try:
					os.remove(old_file)
				except Exception:
					pass
	except Exception as ex:
		log.debug(f"outlookImageHelper: _cleanup_old_reports błąd: {ex}")


def save_report_to_diagnostics_folder(reportContent: str, addon_prefix: str = "OutlookImageHelper") -> str:
	"""
	Zapisuje raport diagnostyczny w bezpiecznym folderze w profilu konfiguracji użytkownika NVDA.
	Gwarantuje limit rozmiaru (maks. 100 KB) oraz automatycznie usuwa najstarsze własne raporty (FIFO).
	"""
	encoded = reportContent.encode("utf-8", errors="replace")
	if len(encoded) > MAX_REPORT_BYTES:
		truncated_text = encoded[:MAX_REPORT_BYTES - 256].decode("utf-8", errors="ignore")
		reportContent = truncated_text + "\n\n[...RAPORT OGRANICZONO DO 100 KB W CELU OCHRONY DYSKU...]\n"

	target_dir = os.path.join(getConfigDir(), "Raporty_Diagnostyczne")
	try:
		os.makedirs(target_dir, exist_ok=True)
	except Exception:
		target_dir = tempfile.gettempdir()

	filename = f"Raport_{addon_prefix}_{time.strftime('%Y%m%d_%H%M%S')}.txt"
	saved_path = os.path.join(target_dir, filename)

	try:
		with open(saved_path, "w", encoding="utf-8") as f:
			f.write(reportContent)
	except Exception as e:
		log.error(f"outlookImageHelper: Błąd zapisu raportu diagnostycznego: {e}")

	_cleanup_old_reports(target_dir, addon_prefix=addon_prefix, max_count=MAX_REPORTS_RETAINED)

	return saved_path


def generate_diagnostic_report(level: str = "basic") -> str:
	"""
	Generuje szczegółowy raport diagnostyczny środowiska i stanu dodatku Outlook Image Helper.
	Poziom 'basic': podstawowe informacje, konfiguracja, audio i koordynacja.
	Poziom 'advanced': pełna telemetria audio (30 zdarzeń), diagnostyka COM/UIA, wyciąg z logu NVDA.
	Klucze API są zawsze maskowane (np. AIzaSy...***) w celu ochrony prywatności.
	"""
	is_adv = (level == "advanced")
	lines = []
	lines.append("=" * 65)
	lines.append(f"OUTLOOK IMAGE HELPER - RAPORT DIAGNOSTYCZNY ({'ZAAWANSOWANY / ADVANCED' if is_adv else 'PODSTAWOWY / BASIC'})")
	lines.append("=" * 65)
	lines.append(f"Data wygenerowania / Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}")
	lines.append(f"Wersja dodatku / Add-on Version: {ADDON_VERSION} ({ADDON_NAME})")
	effLang = getEffectiveLanguage()
	confLang = config.conf.get(CONFIG_SECTION, {}).get("language", "auto")
	lines.append(f"Język interfejsu / UI Language: {confLang} (efektywny: {effLang})")
	lines.append("")

	lines.append("--- ŚRODOWISKO SYSTEMOWE / SYSTEM ENVIRONMENT ---")
	try:
		lines.append(f"System operacyjny / OS: {platform.platform()} ({platform.architecture()[0]})")
		lines.append(f"Wersja Windows / Windows build: {sys.getwindowsversion()}")
		lines.append(f"Wersja Python: {sys.version.split()[0]} ({sys.executable})")
	except Exception as e:
		lines.append(f"Błąd odczytu środowiska: {e}")
	try:
		import versionInfo
		lines.append(f"NVDA Version: {versionInfo.version} ({versionInfo.version_year}.{versionInfo.version_major})")
	except Exception as e:
		lines.append(f"NVDA Version: {e}")
	try:
		lines.append(f"Katalog konfiguracji NVDA / Config Dir: {getConfigDir()}")
	except Exception:
		pass
	lines.append("")

	lines.append("--- STAN PROGRAMU MICROSOFT OUTLOOK I SKŁADNIKÓW ---")
	try:
		import psutil
		procs = [p for p in psutil.process_iter(['name']) if p.info.get('name') and 'outlook' in p.info['name'].lower()]
		lines.append(f"Proces Outlook uruchomiony: {'TAK' if procs else 'NIE'} (liczba procesów: {len(procs)})")
	except Exception:
		lines.append("psutil niedostępny do weryfikacji procesów")

	try:
		focus = api.getFocusObject()
		appMod = getattr(focus, "appModule", None)
		lines.append(f"Fokus w aplikacji: {getattr(appMod, 'appName', None)}")
		if appMod and getattr(appMod, "appName", "").lower() == "outlook":
			nativeOm = getattr(appMod, "nativeOm", None)
			if nativeOm:
				try:
					lines.append(f"Wersja Outlook COM nativeOm: {nativeOm.Version}")
				except Exception as ex:
					lines.append(f"Wersja Outlook COM nativeOm: {ex}")
	except Exception as e:
		lines.append(f"Błąd sprawdzania fokusu Outlooka: {e}")

	lines.append(f"Skrypt OCR worker (PowerShell): {'ISTNIEJE' if os.path.isfile(OCR_SCRIPT_PATH) else 'BRAK'} ({OCR_SCRIPT_PATH})")
	lines.append("")

	lines.append("--- STAN PODSYSTEMU AUDIO I PLIKÓW DŹWIĘKOWYCH ---")
	try:
		vol = ctypes.c_ulong()
		if ctypes.windll.winmm.waveOutGetVolume(0, ctypes.byref(vol)) == 0:
			left = (vol.value & 0xFFFF) * 100 // 65535
			right = ((vol.value >> 16) & 0xFFFF) * 100 // 65535
			lines.append(f"Głośność systemowa Windows (waveOut): Lewy {left}%, Prawy {right}%")
	except Exception:
		lines.append("Podsystem audio Windows: Aktywny")

	sec = config.conf.get(CONFIG_SECTION, {})
	theme = sec.get("soundTheme", "associative")
	volSetting = sec.get("soundVolume", 80)
	trigMode = sec.get("soundTriggerMode", "open_only")
	lines.append(f"Wybrany motyw dźwiękowy: {theme}")
	lines.append(f"Głośność dźwięków dodatku: {volSetting}%")
	lines.append(f"Miejsce odtwarzania dźwięków: {trigMode}")

	theme_dir = os.path.join(SOUNDS_DIR, theme)
	alert_cand = os.path.join(theme_dir, "image_alert.wav")
	alert_exists = os.path.isfile(alert_cand) or os.path.isfile(ALERT_SOUND_PATH)
	lines.append(f"Główny dźwięk alertu (image_alert.wav): {'ISTNIEJE (' + str(os.path.getsize(alert_cand if os.path.isfile(alert_cand) else ALERT_SOUND_PATH)) + ' B)' if alert_exists else 'BRAK'}")
	lines.append("")

	lines.append("--- KOORDYNACJA I DEDUPLIKACJA Z OUTLOOK SPEECH FILTER ---")
	sf_installed = isSpeechFilterInstalled()
	sf_handling = isSpeechFilterHandlingInlineImages()
	lines.append(f"Dodatek Outlook Speech Filter zainstalowany / aktywny: {'TAK' if sf_installed else 'NIE'}")
	lines.append(f"Speech Filter obsługuje dźwięki grafik na liście: {'TAK (Pomocnik obrazów wycisza duplikaty na liście)' if sf_handling else 'NIE (Pomocnik obrazów obsługuje listę samodzielnie)'}")
	lines.append("")

	lines.append("--- AKTUALNA KONFIGURACJA DODATKU ---")
	SAFE_CONFIG_KEYS = {
		"autoAnnounce", "playSound", "navSounds", "aiNaming", "showAdvancedSettings",
		"minDimension", "geminiModel", "openaiModel", "aiProvider", "defaultOcrLang",
		"soundTheme", "soundVolume", "soundTriggerMode", "updateSourceUrl",
		"autoCheckUpdates", "contextMenuEnabled", "uiLanguage"
	}
	for k, v in sorted(sec.items()):
		k_lower = k.lower()
		if any(w in k_lower for w in ("key", "password", "token", "secret", "cred", "hash")):
			key_str = str(v).strip()
			if key_str:
				masked = key_str[:4] + "..." + ("*" * 6) + f" [długość: {len(key_str)} znaków]"
			else:
				masked = "(nie skonfigurowano)"
			lines.append(f"  {k} = {masked}")
		elif "prompt" in k_lower:
			p_val = str(v).strip()
			if p_val:
				lines.append(f"  {k} = [UKRYTO_DLA_OCHRONY_PRYWATNOSCI_RODO: prompt użytkownika, długość: {len(p_val)} znaków]")
			else:
				lines.append(f"  {k} = (pusty)")
		elif k in SAFE_CONFIG_KEYS or not isinstance(v, str) or len(str(v)) < 50:
			lines.append(f"  {k} = {v}")
		else:
			lines.append(f"  {k} = [wartość tekstowa: {len(str(v))} znaków]")
	lines.append("")

	lines.append(f"--- DZIENNIK ZDARZEŃ AUDIO ({'OSTATNIE 30 ZDARZEŃ' if is_adv else 'OSTATNIE 5 ZDARZEŃ'}) ---")
	if _sound_events_history:
		events_slice = _sound_events_history if is_adv else _sound_events_history[-5:]
		for ev in events_slice:
			lines.append(f"  [{ev.get('time', '')}] Dźwięk: {ev.get('sound', '')} | Głośność: {ev.get('volume', '')} | Kontekst: {ev.get('context', '')} | Wynik: {ev.get('status', '')}")
	else:
		lines.append("  (Brak zarejestrowanych zdarzeń audio w bieżącej sesji NVDA)")
	lines.append("")

	if is_adv:
		lines.append("--- SZCZEGÓŁOWA DIAGNOSTYKA COM / UIA (ZAAWANSOWANA) ---")
		try:
			focus = api.getFocusObject()
			lines.append(f"  Klasa obiektu fokusu NVDA: {type(focus).__name__}")
			focus_name = getattr(focus, "name", "") or ""
			lines.append(f"  Nazwa (Name): {'obecna (' + str(len(focus_name)) + ' znaków, ukryto dla ochrony prywatności RODO)' if focus_name else 'brak'}")
			lines.append(f"  Rola (Role): {getattr(focus, 'role', '')}")
			lines.append(f"  Klasa okna (windowClassName): {getattr(focus, 'windowClassName', '')}")
			lines.append(f"  Uchwyt okna (windowHandle): {getattr(focus, 'windowHandle', 0)}")
			parent = getattr(focus, "parent", None)
			if parent:
				lines.append(f"  Obiekt nadrzędny (Parent): {type(parent).__name__} | Rola: {getattr(parent, 'role', '')}")
		except Exception as ex:
			lines.append(f"  Błąd diagnostyki fokusu: {ex}")

		lines.append("")
		lines.append("--- WYCIĄG Z DZIENNIKA ZDARZEŃ NVDA (OSTATNIE WPISY) ---")
		try:
			logFile = getattr(getattr(globalVars, "appArgs", None), "logFileName", None)
			if logFile and os.path.isfile(logFile):
				with open(logFile, "r", encoding="utf-8", errors="replace") as lf:
					all_lines = lf.readlines()
				filtered = [
					l.strip() for l in all_lines
					if any(k in l.lower() for k in ("outlook", "imagehelper", "speechfilter", "error", "exception", "traceback"))
					and not any(s in l.lower() for s in ("speaking [", "io - speech", "speaktext"))
				]
				for fl in filtered[-60:]:
					redacted_fl = fl
					for sec_k, sec_v in sec.items():
						if any(w in sec_k.lower() for w in ("key", "password", "token")):
							sv = str(sec_v).strip()
							if len(sv) >= 6 and sv in redacted_fl:
								redacted_fl = redacted_fl.replace(sv, "[ZREDAGOWANO_SEKRET]")
					redacted_fl = re.sub(r'sk-[a-zA-Z0-9_\-]{20,}', '[ZREDAGOWANO_KLUCZ_OPENAI]', redacted_fl)
					redacted_fl = re.sub(r'AIza[a-zA-Z0-9_\-]{30,}', '[ZREDAGOWANO_KLUCZ_GEMINI]', redacted_fl)
					redacted_fl = re.sub(r'[A-Za-z]:\\[Uu]sers\\[^\\]+', r'C:\\Users\\[UZYTKOWNIK]', redacted_fl)
					redacted_fl = re.sub(r'att_[a-f0-9]{12}', 'att_[ZREDAGOWANO_ID]', redacted_fl)
					redacted_fl = re.sub(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', '[ZREDAGOWANO_EMAIL]', redacted_fl)
					lines.append(f"  {redacted_fl}")
			else:
				lines.append("  Brak dostępu do pliku dziennika NVDA")
		except Exception as ex:
			lines.append(f"  Błąd odczytu dziennika NVDA: {ex}")
		lines.append("")

	lines.append("=" * 65)
	lines.append("KONIEC RAPORTU DIAGNOSTYCZNEGO")
	lines.append("=" * 65)
	return "\n".join(lines)


class DiagnosticReportDialog(wx.Dialog):
	"""Okno dialogowe kontaktu z twórcą, zgłaszania problemów (GitHub Issues) oraz raportu diagnostycznego."""

	def __init__(self, parent=None):
		super().__init__(
			parent,
			title="Kontakt z twórcą i zgłaszanie uwag (GitHub) - Outlook Image Helper",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
			size=(680, 520)
		)
		self.reportContent = ""
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		introLabel = wx.StaticText(
			self,
			label="Ten moduł umożliwia bezpośredni kontakt z twórcą dodatku, zgłaszanie problemów technicznych, "
			      "propozycji nowych funkcji oraz przesyłanie opinii w oficjalnym serwisie GitHub (GitHub Issues).\n\n"
			      "Możesz otworzyć formularz zgłoszeń w przeglądarce internetowej oraz jednym kliknięciem skopiować "
			      "bezpieczny raport diagnostyczny (konfiguracja i wersje bez treści wiadomości czy haseł), aby wkleić go do zgłoszenia."
		)
		introLabel.Wrap(640)
		mainSizer.Add(introLabel, flag=wx.ALL, border=10)

		# Sekcja przycisków głównych kontaktu
		actionBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnGitHub = wx.Button(self, label="&Otwórz stronę zgłoszeń na GitHubie (GitHub Issues)...")
		self.btnGitHub.Bind(wx.EVT_BUTTON, self.onOpenGitHubIssues)
		actionBox.Add(self.btnGitHub, flag=wx.RIGHT, border=10)

		self.copyBtn = wx.Button(self, label="&Kopiuj raport diagnostyczny do schowka")
		self.copyBtn.Bind(wx.EVT_BUTTON, self.onCopyReport)
		actionBox.Add(self.copyBtn)
		mainSizer.Add(actionBox, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM, border=10)

		# Sekcja szczegółowości raportu
		lvlBox = wx.BoxSizer(wx.HORIZONTAL)
		lvlLabel = wx.StaticText(self, label="Szczegółowość raportu diagnostycznego:")
		lvlBox.Add(lvlLabel, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.rbBasic = wx.RadioButton(self, label="&Podstawowy (konfiguracja i audio)", style=wx.RB_GROUP)
		self.rbBasic.SetValue(True)
		self.rbBasic.Bind(wx.EVT_RADIOBUTTON, self.onLevelChanged)
		lvlBox.Add(self.rbBasic, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.rbAdvanced = wx.RadioButton(self, label="&Zaawansowany (pełna telemetria i dziennik NVDA)")
		self.rbAdvanced.Bind(wx.EVT_RADIOBUTTON, self.onLevelChanged)
		lvlBox.Add(self.rbAdvanced, flag=wx.ALIGN_CENTER_VERTICAL)

		mainSizer.Add(lvlBox, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM, border=10)

		# Sekcja podglądu i zapisu
		btnBox = wx.BoxSizer(wx.HORIZONTAL)
		self.viewBtn = wx.Button(self, label="&Podgląd raportu...")
		self.viewBtn.Bind(wx.EVT_BUTTON, self.onViewReport)
		btnBox.Add(self.viewBtn, flag=wx.RIGHT, border=10)

		self.saveBtn = wx.Button(self, label="&Zapisz raport na Pulpicie...")
		self.saveBtn.Bind(wx.EVT_BUTTON, self.onSaveReport)
		btnBox.Add(self.saveBtn)

		mainSizer.Add(btnBox, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM, border=10)

		closeSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.closeBtn = wx.Button(self, wx.ID_CANCEL, label="Zam&knij")
		closeSizer.Add(self.closeBtn)
		mainSizer.Add(closeSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=10)

		self.SetSizer(mainSizer)
		self.Layout()
		self.CentreOnParent()

	def onOpenGitHubIssues(self, evt):
		safeOpenUrl("https://github.com/strugo10/outlookImageHelper/issues")

	def getReport(self) -> str:
		level = "advanced" if self.rbAdvanced.GetValue() else "basic"
		return generate_diagnostic_report(level=level)

	def onLevelChanged(self, evt):
		self.reportContent = ""

	def onCopyReport(self, evt):
		rep = self.getReport()
		ok = False
		try:
			import api
			if hasattr(api, "copyToClip"):
				api.copyToClip(rep)
				ok = True
		except Exception:
			pass
		if not ok:
			try:
				if wx.TheClipboard.Open():
					wx.TheClipboard.SetData(wx.TextDataObject(rep))
					wx.TheClipboard.Close()
					ok = True
			except Exception:
				pass
		if ok:
			ui.message("Raport techniczny został skopiowany do schowka. Możesz wkleić go do swojego zgłoszenia na GitHubie.")
		else:
			ui.message("Nie udało się skopiować raportu do schowka.")

	def onViewReport(self, evt):
		rep = self.getReport()
		dlg = AccessibleMessageDialog(
			self,
			message=rep,
			title="Podgląd raportu diagnostycznego"
		)
		dlg.ShowModal()
		dlg.Destroy()

	def onSaveReport(self, evt):
		rep = self.getReport()
		saved_panel = save_report_to_diagnostics_folder(rep, addon_prefix="OutlookImageHelper")

		desktop = os.path.join(os.path.expanduser("~"), "Desktop")
		fn = f"Raport_OutlookImageHelper_{time.strftime('%Y%m%d_%H%M%S')}.txt"
		desk_path = os.path.join(desktop, fn)
		try:
			with open(desk_path, "w", encoding="utf-8") as f:
				f.write(rep)
		except Exception:
			desk_path = saved_panel

		msg = f"Raport diagnostyczny został pomyślnie zapisany w lokalizacjach:\n1. Pulpit: {desk_path}\n2. Panel Administratora: {saved_panel}"
		dlg = AccessibleMessageDialog(self, message=msg, title="Raport zapisany")
		dlg.ShowModal()
		dlg.Destroy()


def extractImgInfoFromHtml(html: str) -> list[dict]:
	"""
	Wyodrębnia informacje o obrazach z kodu HTMLBody wiadomości:
	adres źródłowy (src), tekst alternatywny (alt), tytuł (title) oraz powiązane hiperłącze (href).
	Wykorzystuje prekompilowane wyrażenia regularne dla maksymalnej szybkości.
	"""
	if not html:
		return []
	results = []
	try:
		tokens = _RE_HTML_TOKENS.finditer(html)
		currentHref = ""
		for m in tokens:
			if m.group('href'):
				currentHref = m.group('href').strip()
			else:
				imgTag = m.group(0)
				sm = _RE_IMG_SRC.search(imgTag)
				am = _RE_IMG_ALT.search(imgTag)
				tm = _RE_IMG_TITLE.search(imgTag)
				if sm:
					s = sm.group(1).strip()
					a = am.group(1).strip() if am else ""
					t = tm.group(1).strip() if tm else ""
					results.append({
						"src": s,
						"alt": a,
						"title": t,
						"href": currentHref,
					})
	except Exception as ex:
		log.debug(f"outlookImageHelper: extractImgInfoFromHtml error: {ex}")
	return results


def collectGraphics(doc, minDimension: int = 15, fast: bool = False) -> list[GraphicItem]:
	"""Przeszukuje dokument Worda i zwraca listę wszystkich wykrytych grafik wraz z kontekstem i linkami."""
	items: list[GraphicItem] = []
	if not doc:
		return items

	msgKey = getMessageKey()
	customNames = loadAllCustomNames().get(msgKey, {}) if msgKey else {}
	allImageSettings = loadAllImageSettings().get(msgKey, {}) if msgKey else {}

	# Metadane obrazów pobierane są bezpośrednio z modelu WordEditor (InlineShapes, Shapes, Hyperlinks),
	# co zapobiega wywoływaniu ostrzeżeń bezpieczeństwa Outlook (Object Model Guard) przy odczycie HTMLBody.
	html_imgs: list[dict] = []

	# 1. Przeszukaj InlineShapes (osadzone w tekście)
	try:
		inlineCount = int(doc.InlineShapes.Count)
	except Exception:
		inlineCount = 0

	for i in range(1, inlineCount + 1):
		try:
			ishape = doc.InlineShapes.Item(i)
			w = float(ishape.Width)
			h = float(ishape.Height)
			if w < minDimension and h < minDimension:
				continue

			alt = getattr(ishape, "AlternativeText", "") or ""
			title = ""
			try:
				title = getattr(ishape, "Title", "") or ""
			except Exception:
				pass

			rngStart = 0
			rngEnd = 0
			try:
				rng = ishape.Range
				rngStart = int(rng.Start)
				rngEnd = int(rng.End)
			except Exception:
				pass

			cid = ""
			sourceUrl = ""
			try:
				lf = getattr(ishape, "LinkFormat", None)
				if lf and getattr(lf, "SourceFullName", None):
					sfn = str(lf.SourceFullName).strip()
					if sfn.startswith("http://") or sfn.startswith("https://"):
						sourceUrl = sfn
					else:
						cid = sfn
			except Exception:
				pass

			hlUrl, hlTip = extractHyperlinkFromShape(ishape, isInline=True)

			if not sourceUrl and hlUrl:
				cleanHl = hlUrl.split("?")[0].lower()
				if any(cleanHl.endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp")):
					sourceUrl = hlUrl

			ctx = "" if fast else extractContextTextForShape(doc, ishape, isInline=True)

			itemIdx = len(items) + 1

			gItem = GraphicItem(
				index=itemIdx,
				isInline=True,
				rawObject=ishape,
				widthPt=w,
				heightPt=h,
				altText=alt,
				title=title,
				rangeStart=rngStart,
				rangeEnd=rngEnd,
				cid=cid,
				contextText=ctx,
				hyperlinkUrl=hlUrl,
				hyperlinkTip=hlTip,
				customName="",
				msgKey=msgKey,
				isInverted=False,
				sourceUrl=sourceUrl,
			)
			checkText = f"{alt} {title} {cid} {sourceUrl}".lower()
			if any(k in checkText for k in ("spacer", "blank", "tracking", "pixel", "clear.gif")):
				gItem.isBlankOrSolid = True
				gItem.blankReason = "Pusty ozdobnik (spacer)"

			items.append(gItem)
		except Exception as e:
			log.debug(f"outlookImageHelper: Błąd podczas odczytu InlineShape {i}: {e}")

	# 2. Przeszukaj Shapes (grafiki swobodne z opływaniem tekstu)
	try:
		shapeCount = int(doc.Shapes.Count)
	except Exception:
		shapeCount = 0

	for i in range(1, shapeCount + 1):
		try:
			shape = doc.Shapes.Item(i)
			w = float(shape.Width)
			h = float(shape.Height)
			if w < minDimension and h < minDimension:
				continue

			alt = getattr(shape, "AlternativeText", "") or ""
			title = ""
			try:
				title = getattr(shape, "Title", "") or ""
			except Exception:
				pass

			rngStart = 0
			rngEnd = 0
			try:
				anchor = shape.Anchor
				rngStart = int(anchor.Start)
				rngEnd = int(anchor.End)
			except Exception:
				pass

			cid = ""
			sourceUrl = ""
			try:
				lf = getattr(shape, "LinkFormat", None)
				if lf and getattr(lf, "SourceFullName", None):
					sfn = str(lf.SourceFullName).strip()
					if sfn.startswith("http://") or sfn.startswith("https://"):
						sourceUrl = sfn
					else:
						cid = sfn
			except Exception:
				pass

			hlUrl, hlTip = extractHyperlinkFromShape(shape, isInline=False)

			if not sourceUrl and hlUrl:
				cleanHl = hlUrl.split("?")[0].lower()
				if any(cleanHl.endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp")):
					sourceUrl = hlUrl

			ctx = "" if fast else extractContextTextForShape(doc, shape, isInline=False)

			itemIdx = len(items) + 1

			gItem = GraphicItem(
				index=itemIdx,
				isInline=False,
				rawObject=shape,
				widthPt=w,
				heightPt=h,
				altText=alt,
				title=title,
				rangeStart=rngStart,
				rangeEnd=rngEnd,
				cid=cid,
				contextText=ctx,
				hyperlinkUrl=hlUrl,
				hyperlinkTip=hlTip,
				customName="",
				msgKey=msgKey,
				isInverted=False,
				sourceUrl=sourceUrl,
			)
			checkText = f"{alt} {title} {sourceUrl}".lower()
			if any(k in checkText for k in ("spacer", "blank", "tracking", "pixel", "clear.gif")):
				gItem.isBlankOrSolid = True
				gItem.blankReason = "Pusty ozdobnik (spacer)"

			items.append(gItem)
		except Exception as e:
			log.debug(f"outlookImageHelper: Błąd podczas odczytu Shape {i}: {e}")

	# Posortuj elementy według pozycji początkowej w tekście (od góry do dołu wiadomości)
	items.sort(key=lambda it: it.rangeStart)
	for idx, it in enumerate(items, 1):
		it.index = idx
		it.customName = customNames.get(str(idx), "")
		it.isInverted = bool(allImageSettings.get(str(idx), {}).get("inverted", False))

	return items


def getOrCollectGraphics(
	doc,
	minDimension: int = 15,
	fast: bool = True,
	forceRefresh: bool = False,
) -> list[GraphicItem]:
	"""
	Zwraca listę grafik z dokumentu WordMail.
	Wykorzystuje pamięć podręczną (cache z czasem życia TTL 15s), co sprawia,
	że kolejne naciśnięcia klawiszy G, Shift+G oraz ruch strzałkami wykonują się natychmiast (< 1 ms).
	"""
	if not doc:
		return []

	now = time.time()
	docKey = None
	try:
		docKey = (
			getattr(doc, "Name", ""),
			getattr(getattr(doc, "ActiveWindow", None), "Hwnd", 0),
		)
	except Exception:
		try:
			docKey = getattr(doc, "Name", "")
		except Exception:
			docKey = id(doc)

	if GlobalPlugin.instance and not forceRefresh:
		cachedKey = getattr(GlobalPlugin.instance, "_cachedGraphicsDocKey", None)
		cachedItems = getattr(GlobalPlugin.instance, "_cachedGraphicsItems", None)
		cachedTime = getattr(GlobalPlugin.instance, "_cachedGraphicsTime", 0.0)
		cachedFast = getattr(GlobalPlugin.instance, "_cachedGraphicsFast", True)
		if cachedKey == docKey and cachedItems and (now - cachedTime < 15.0):
			if not (not fast and cachedFast):
				return cachedItems

	items = collectGraphics(doc, minDimension=minDimension, fast=fast)
	if GlobalPlugin.instance and items:
		GlobalPlugin.instance._cachedGraphicsDocKey = docKey
		GlobalPlugin.instance._cachedGraphicsItems = items
		GlobalPlugin.instance._cachedGraphicsTime = now
		GlobalPlugin.instance._cachedGraphicsFast = fast

	return items


def getStrictGraphicUnderCaret(
	doc=None,
	treeInterceptor=None,
	targetPos: int | None = None,
) -> GraphicItem | None:
	"""
	Zwraca obiekt GraphicItem TYLKO I WYŁĄCZNIE wtedy, gdy kursor/fokus NVDA
	znajduje się bezpośrednio i ściśle na grafice (a NIE na tekście).
	Zwraca None, jeśli użytkownik czyta tekst lub kursor stoi na tekście wiadomości.
	"""
	if not doc:
		doc = _getDocFromTreeInterceptorOrFocus(treeInterceptor)
	if not doc:
		return None

	items = getOrCollectGraphics(
		doc,
		minDimension=config.conf[CONFIG_SECTION].get("minDimension", 15),
		fast=True,
	)
	if not items:
		return None

	# Pobierz bieżącą pozycję kursora w dokumencie
	curPos = targetPos
	tiInstance = treeInterceptor
	if not tiInstance:
		focus = api.getFocusObject()
		tiInstance = getattr(focus, "treeInterceptor", None)

	if curPos is None and tiInstance and getattr(tiInstance, "isAlive", False):
		try:
			ti = getattr(tiInstance, "selection", None)
			if ti and hasattr(ti, "_rangeObj") and ti._rangeObj:
				try:
					curPos = int(ti._rangeObj.Start)
				except Exception:
					pass
		except Exception:
			pass

	if curPos is None:
		try:
			sel = getattr(doc.Application, "Selection", None)
			if sel:
				selStart = getattr(sel, "Start", None)
				if selStart is not None:
					curPos = int(selStart() if callable(selStart) else selStart)
		except Exception:
			pass

	# 1. Jeśli użytkownik nawigował po grafikach klawiszem G / Shift+G:
	# upewnij się, że aktualny kursor wciąż stoi na tej grafice, a nie przesunął się na tekst
	if GlobalPlugin.instance and getattr(GlobalPlugin.instance, "_currentGraphicItem", None):
		cand = GlobalPlugin.instance._currentGraphicItem
		matched = next((it for it in items if it.index == cand.index), None)
		if matched:
			if curPos is not None:
				if (matched.rangeStart - 2) <= curPos <= (matched.rangeEnd + 2):
					return matched
				else:
					GlobalPlugin.instance._currentGraphicItem = None
			else:
				# Jeśli nie można potwierdzić pozycji, nie polegaj na starym cache
				pass

	# 2. Sprawdź, czy w edytorze Word obiekt graficzny jest zaznaczony lub kursor Worda stoi na grafice
	try:
		sel = getattr(doc.Application, "Selection", None)
		if sel:
			selType = getattr(sel, "Type", 0)
			selStart = getattr(sel, "Start", None)
			selStartVal = int(selStart() if callable(selStart) else selStart) if selStart is not None else None
			# 7: wdSelectionInlineShape, 8: wdSelectionShape
			if selType in (7, 8):
				if selStartVal is not None:
					for it in items:
						if abs(it.rangeStart - selStartVal) <= 5 or (it.rangeStart <= selStartVal <= it.rangeEnd):
							return it
				return items[0]
			if selStartVal is not None:
				for it in items:
					if (it.rangeStart <= selStartVal <= it.rangeEnd) or abs(it.rangeStart - selStartVal) <= 1:
						return it
	except Exception as ex:
		log.debug(f"outlookImageHelper: getStrictGraphicUnderCaret selection check error: {ex}")

	# 3. Sprawdź pozycję w buforze przeglądania / dokumencie
	if curPos is not None:
		for it in items:
			if (it.rangeStart <= curPos <= it.rangeEnd) or abs(it.rangeStart - curPos) <= 1:
				return it

	# 4. Jeśli podano targetPos bez bufora przeglądania
	if targetPos is not None:
		for it in items:
			if (it.rangeStart <= targetPos < it.rangeEnd) or (targetPos == it.rangeStart):
				return it

	# 5. Sprawdź, czy obiekt fokusu NVDA to bezpośrednio grafika (poza buforem przeglądania)
	focus = api.getFocusObject()
	if focus:
		role = getattr(focus, "role", None)
		if role in (
			controlTypes.Role.GRAPHIC,
			getattr(controlTypes.Role, "IMAGE", -1),
			getattr(controlTypes.Role, "PICTURE", -1),
		):
			if len(items) == 1:
				return items[0]
			if GlobalPlugin.instance and getattr(GlobalPlugin.instance, "_currentGraphicItem", None):
				cand = GlobalPlugin.instance._currentGraphicItem
				if any(it.index == cand.index for it in items):
					return cand

	return None


def getGraphicAtCaret(
	doc=None,
	strict: bool = False,
	treeInterceptor=None,
	targetPos: int | None = None,
) -> GraphicItem | None:
	"""
	Zwraca obiekt GraphicItem powiązany z pozycją kursora w edytorze WordMail:
	- Jeśli strict=True: wymaga, aby kursor/fokus znajdował się bezpośrednio i ściśle na grafice.
	- Jeśli strict=False: sprawdza grafikę pod kursorem, ostatnio wybraną grafikę lub najbliższą pozycji kursora.
	"""
	strictItem = getStrictGraphicUnderCaret(doc=doc, treeInterceptor=treeInterceptor, targetPos=targetPos)
	if strict:
		return strictItem
	if strictItem:
		return strictItem

	if not doc:
		doc = _getDocFromTreeInterceptorOrFocus(treeInterceptor)
	if not doc:
		return None

	items = getOrCollectGraphics(
		doc,
		minDimension=config.conf[CONFIG_SECTION].get("minDimension", 15),
		fast=True,
	)
	if not items:
		return None

	# 2. Jeśli nie strict, sprawdź ostatnio nawigowaną grafikę
	if GlobalPlugin.instance and getattr(GlobalPlugin.instance, "_currentGraphicItem", None):
		cand = GlobalPlugin.instance._currentGraphicItem
		for it in items:
			if it.index == cand.index:
				return it

	# 3. Wybierz grafikę położoną najbliżej pozycji kursora
	curPos = targetPos if targetPos is not None else 0
	if curPos == 0 and treeInterceptor:
		try:
			ti = getattr(treeInterceptor, "selection", None)
			if ti and hasattr(ti, "_rangeObj") and ti._rangeObj:
				curPos = int(ti._rangeObj.Start)
		except Exception:
			pass

	if curPos == 0:
		try:
			sel = getattr(doc.Application, "Selection", None)
			if sel:
				curPos = int(getattr(sel, "Start", 0))
		except Exception:
			pass

	closest = min(items, key=lambda it: abs(it.rangeStart - curPos))
	return closest


def _isOutlookTreeInterceptor(treeInterceptorInstance=None) -> bool:
	"""Sprawdza, czy dany bufor przeglądania (treeInterceptor) należy do programu Microsoft Outlook."""
	if treeInterceptorInstance:
		root = getattr(treeInterceptorInstance, "rootNVDAObject", None)
		if root:
			appMod = getattr(root, "appModule", None)
			if appMod and getattr(appMod, "appName", "").lower() == "outlook":
				return True
			if appMod and getattr(appMod, "appName", ""):
				return False
	focus = api.getFocusObject()
	if focus:
		fAppMod = getattr(focus, "appModule", None)
		if fAppMod and getattr(fAppMod, "appName", "").lower() == "outlook":
			return True
	fgHwnd = winUser.getForegroundWindow()
	if fgHwnd and "renwnd32" in (winUser.getClassName(fgHwnd) or "").lower():
		return True
	return False


def _getDocFromTreeInterceptorOrFocus(treeInterceptorInstance=None):
	"""Niezawodnie pobiera dokument Word powiązany z treeInterceptorem lub aktywnym oknem wiadomości w Outlooku."""
	if not _isOutlookTreeInterceptor(treeInterceptorInstance):
		return None
	if treeInterceptorInstance:
		root = getattr(treeInterceptorInstance, "rootNVDAObject", None)
		if root:
			doc = getattr(root, "WinwordDocumentObject", None)
			if doc:
				return doc
			hwnd = getattr(root, "windowHandle", 0)
			if hwnd:
				doc = _getWordDocumentFromHwnd(hwnd)
				if doc:
					return doc
	return getActiveWordDocument()


def navigateGraphicInMessage(doc, direction: str = "next", treeInterceptor=None) -> bool:
	"""
	Nawiguje do następnej lub poprzedniej grafiki bezpośrednio w treści wiadomości programu Outlook.
	Wybiera grafikę w dokumencie Word, aktualizuje kursor bufora przeglądania NVDA,
	odtwarza motyw dźwiękowy i wypowiada wyczerpujący, bezbłędny opis grafiki (zamiast 'pusta').
	"""
	if not doc:
		doc = _getDocFromTreeInterceptorOrFocus(treeInterceptor)
	if not doc:
		ui.message(_T("msg_outlook_only"))
		return False

	items = getOrCollectGraphics(
		doc,
		minDimension=config.conf[CONFIG_SECTION].get("minDimension", 15),
		fast=True,
	)
	if not items:
		if isMessageListFocused() or not isMessageBodyFocused():
			ui.message("Aby przejść do grafik w tej wiadomości, otwórz ją klawiszem Enter.")
		else:
			ui.message("W bieżącej wiadomości nie odnaleziono grafik.")
		return True

	total = len(items)

	docKey = None
	try:
		docKey = (
			getattr(doc, "Name", ""),
			getattr(getattr(doc, "ActiveWindow", None), "Hwnd", 0),
		)
	except Exception:
		docKey = id(doc)

	if GlobalPlugin.instance and getattr(GlobalPlugin.instance, "_lastNavigatedDocKey", None) != docKey:
		GlobalPlugin.instance._lastNavigatedDocKey = docKey
		GlobalPlugin.instance._lastNavigatedIndex = -1
		GlobalPlugin.instance._lastNavigatedCaretPos = -9999
		GlobalPlugin.instance._currentGraphicItem = None

	curPos = 0
	if treeInterceptor:
		try:
			ti = getattr(treeInterceptor, "selection", None)
			if ti and hasattr(ti, "_rangeObj") and ti._rangeObj:
				curPos = int(ti._rangeObj.Start)
		except Exception:
			pass

	sel = getattr(doc.Application, "Selection", None)
	if curPos == 0 and sel:
		try:
			curPos = int(getattr(sel, "Start", 0))
		except Exception:
			curPos = 0

	# Ustal bieżący indeks grafiki
	curIdx = -1
	lastNav = getattr(GlobalPlugin.instance, "_lastNavigatedIndex", -1) if GlobalPlugin.instance else -1
	lastCaretPos = getattr(GlobalPlugin.instance, "_lastNavigatedCaretPos", -9999) if GlobalPlugin.instance else -9999

	if 1 <= lastNav <= total:
		lastItem = items[lastNav - 1]
		if abs(curPos - lastCaretPos) <= 5 or abs(curPos - lastItem.rangeStart) <= 10:
			curIdx = lastNav

	if curIdx == -1 and not (lastNav == -1 and curPos == 0):
		for it in items:
			if it.rangeStart <= curPos <= it.rangeEnd or abs(curPos - it.rangeStart) <= 2:
				curIdx = it.index
				break

	targetItem = None
	if direction == "next":
		if curIdx != -1:
			if curIdx < total:
				targetItem = items[curIdx]  # 0-indexed: index == curIdx + 1
			else:
				ui.message("Ostatnia grafika w wiadomości.")
				return True
		else:
			if lastNav == -1 and curPos == 0:
				targetItem = items[0]
			else:
				candidates = [it for it in items if it.rangeStart > curPos]
				if candidates:
					targetItem = candidates[0]
				else:
					ui.message("Ostatnia grafika w wiadomości.")
					return True
	else:  # previous
		if curIdx != -1:
			if curIdx > 1:
				targetItem = items[curIdx - 2]  # index == curIdx - 1
			else:
				ui.message("Pierwsza grafika w wiadomości.")
				return True
		else:
			candidates = [it for it in items if it.rangeStart < curPos]
			if candidates:
				targetItem = candidates[-1]
			else:
				ui.message("Pierwsza grafika w wiadomości.")
				return True

	if not targetItem:
		return False

	# Ustaw wyciszenie zdarzeń kursorowych na czas przemieszczenia i odczytu
	if GlobalPlugin.instance:
		GlobalPlugin.instance._currentGraphicItem = targetItem
		GlobalPlugin.instance._lastNavigatedDocKey = docKey
		GlobalPlugin.instance._lastNavigatedIndex = targetItem.index
		GlobalPlugin.instance._suppressCaretSpeechUntil = time.time() + 0.85

	# Ustaw kursor w dokumencie Word na wybranej grafice jako poprawny zakres tekstowy
	newCaretPos = targetItem.rangeStart
	try:
		if getattr(targetItem, "isInline", True):
			rng = doc.Range(targetItem.rangeStart, max(targetItem.rangeStart + 1, targetItem.rangeEnd))
			rng.Select()
		else:
			rng = doc.Range(targetItem.rangeStart, targetItem.rangeStart)
			rng.Select()
		if sel:
			newCaretPos = int(getattr(sel, "Start", targetItem.rangeStart))
	except Exception:
		try:
			if sel:
				sel.SetRange(targetItem.rangeStart, targetItem.rangeStart)
				newCaretPos = int(getattr(sel, "Start", targetItem.rangeStart))
		except Exception as ex:
			log.debug(f"outlookImageHelper: navigateGraphic select error: {ex}")

	if GlobalPlugin.instance:
		GlobalPlugin.instance._lastNavigatedCaretPos = newCaretPos

	# Spróbuj zaznaczyć bezpośrednio obiekt graficzny, jeśli Word to umożliwia
	try:
		if getattr(targetItem, "rawObject", None) and hasattr(targetItem.rawObject, "Select"):
			targetItem.rawObject.Select()
	except Exception:
		pass

	# Zsynchronizuj pozycję wirtualnego bufora przeglądania NVDA z pozycją grafiki
	if treeInterceptor and getattr(treeInterceptor, "isAlive", False):
		try:
			root = getattr(treeInterceptor, "rootNVDAObject", None)
			if root and hasattr(root, "makeTextInfo"):
				import textInfos
				tiInfo = root.makeTextInfo(textInfos.POSITION_CARET)
				if hasattr(tiInfo, "_rangeObj"):
					tiInfo._rangeObj = doc.Range(targetItem.rangeStart, max(targetItem.rangeStart + 1, targetItem.rangeEnd))
					treeInterceptor.selection = tiInfo
		except Exception as ex:
			log.debug(f"outlookImageHelper: navigateGraphic treeInterceptor sync error: {ex}")

	# Określ czytelny prefiks pozycji (Pierwsza / Ostatnia / Jedyna grafika)
	if total == 1:
		prefixMsg = "Jedyna grafika: "
	elif targetItem.index == 1:
		prefixMsg = "Pierwsza grafika: "
	elif targetItem.index == total:
		prefixMsg = "Ostatnia grafika: "
	else:
		prefixMsg = ""

	# Zbuduj bogaty, dostępny komunikat mowy
	descParts = [f"{prefixMsg}Grafika {targetItem.index} z {total}"]
	if getattr(targetItem, "customName", ""):
		descParts.append(f": '{targetItem.customName}'")
	elif getattr(targetItem, "isBlankOrSolid", False):
		reason = getattr(targetItem, "blankReason", "") or "pusty ozdobnik"
		descParts.append(f" ({reason})")
	elif getattr(targetItem, "isDivider", False):
		descParts.append(" (linia podziału)")
	elif getattr(targetItem, "isScreenshotOrUnlabeled", False):
		descParts.append(" (zrzut ekranu)")
	elif getattr(targetItem, "altText", ""):
		descParts.append(f" ({targetItem.altText})")
	else:
		descParts.append(" (obraz)")

	descParts.append(f", {targetItem.widthPx} na {targetItem.heightPx} px")

	if getattr(targetItem, "isInverted", False):
		descParts.append(", w negatywie")

	if getattr(targetItem, "hyperlinkUrl", ""):
		dom = getCleanDomain(targetItem.hyperlinkUrl)
		descParts.append(f", link: {dom}" if dom else ", zawiera link")

	if getattr(targetItem, "ocrText", "") and targetItem.ocrText != "(Brak tekstu)":
		descParts.append(f", OCR: '{_cleanSnippet(targetItem.ocrText, 35)}'")
	elif getattr(targetItem, "aiDescription", ""):
		descParts.append(f", AI: '{_cleanSnippet(targetItem.aiDescription, 35)}'")

	descParts.append(". Shift+F10: czynności, Alt+R: OCR.")

	fullAnnounce = "".join(descParts)

	# Opóźniona zapowiedź przez callLater, aby zdarzenia kursorowe Worda nie ucięły mowy
	import core
	def _doAnnounce():
		try:
			if GlobalPlugin.instance:
				GlobalPlugin.instance._suppressCaretSpeechUntil = time.time() + 0.40
			playItemTypeSound(targetItem, debounce_ms=0)
			try:
				import speech
				speech.cancelSpeech()
			except Exception:
				pass
			ui.message(fullAnnounce)
		except Exception as ex:
			log.error(f"outlookImageHelper: _doAnnounce error: {ex}", exc_info=True)
			try:
				ui.message(fullAnnounce)
			except Exception:
				pass

	core.callLater(55, _doAnnounce)
	return True


def _treeInterceptor_script_nextGraphic(treeInterceptorInstance, gesture):
	if _isOutlookTreeInterceptor(treeInterceptorInstance):
		try:
			doc = _getDocFromTreeInterceptorOrFocus(treeInterceptorInstance)
			if doc:
				navigateGraphicInMessage(doc, direction="next", treeInterceptor=treeInterceptorInstance)
				return
		except Exception as ex:
			log.warning(f"outlookImageHelper: _treeInterceptor_script_nextGraphic error: {ex}", exc_info=True)
			return

	orig = getattr(treeInterceptorInstance, "_orig_script_nextGraphic", None)
	if orig:
		orig(gesture)
	else:
		gesture.send()


def _treeInterceptor_script_previousGraphic(treeInterceptorInstance, gesture):
	if _isOutlookTreeInterceptor(treeInterceptorInstance):
		try:
			doc = _getDocFromTreeInterceptorOrFocus(treeInterceptorInstance)
			if doc:
				navigateGraphicInMessage(doc, direction="previous", treeInterceptor=treeInterceptorInstance)
				return
		except Exception as ex:
			log.warning(f"outlookImageHelper: _treeInterceptor_script_previousGraphic error: {ex}", exc_info=True)
			return

	orig = getattr(treeInterceptorInstance, "_orig_script_previousGraphic", None)
	if orig:
		orig(gesture)
	else:
		gesture.send()


def _treeInterceptor_script_contextMenu(treeInterceptorInstance, gesture):
	"""Obsługuje Shift+F10 oraz klawisz Aplikacje: dla grafiki otwiera menu czynności dodatku, dla tekstu natywne menu Outlooka."""
	if not _isOutlookTreeInterceptor(treeInterceptorInstance):
		orig = getattr(treeInterceptorInstance, "_orig_script_passThrough", None)
		if orig:
			orig(gesture)
		else:
			gesture.send()
		return

	if GlobalPlugin.instance and getattr(GlobalPlugin.instance, "_bypassNextContextMenu", False):
		GlobalPlugin.instance._bypassNextContextMenu = False
		orig = getattr(treeInterceptorInstance, "_orig_script_passThrough", None)
		if orig:
			orig(gesture)
		else:
			gesture.send()
		return

	# Sprawdź, czy funkcja menu kontekstowego dla grafik jest włączona w konfiguracji
	contextMenuEnabled = True
	try:
		contextMenuEnabled = _to_bool(config.conf[CONFIG_SECTION].get("contextMenuEnabled", True), default=True)
	except Exception:
		pass

	if not contextMenuEnabled:
		if GlobalPlugin.instance:
			GlobalPlugin.instance._bypassNextContextMenu = True
		orig = getattr(treeInterceptorInstance, "_orig_script_passThrough", None)
		if orig:
			orig(gesture)
		else:
			gesture.send()
		return

	try:
		doc = _getDocFromTreeInterceptorOrFocus(treeInterceptorInstance)
		item = None
		if doc:
			# Ścisła weryfikacja: tylko gdy kursor NVDA stoi bezpośrednio na grafice, a NIE na tekście
			item = getStrictGraphicUnderCaret(doc, treeInterceptor=treeInterceptorInstance)

		if item:
			import wx
			wx.CallAfter(showGraphicActionsMenu, item)
			return
	except Exception as ex:
		log.debug(f"outlookImageHelper: _treeInterceptor_script_contextMenu error: {ex}")

	# Kursor NIE stoi ściśle na grafice (jest na tekście, nagłówku lub innym elemencie):
	# Przekaż gest do standardowego menu Outlooka. Ustaw flagę obejścia dla GlobalPlugin.
	if GlobalPlugin.instance:
		GlobalPlugin.instance._bypassNextContextMenu = True

	orig = getattr(treeInterceptorInstance, "_orig_script_passThrough", None)
	if orig:
		orig(gesture)
	else:
		gesture.send()


def _treeInterceptor_speakTextInfo(treeInterceptorInstance, info, *args, **kwargs):
	"""
	Przechwytuje odczyt wiersza w buforze przeglądania wiadomości WordMail/Outlook.
	Jeśli wiersz nie zawiera tekstu (lub zawiera wyłącznie znak obiektu osadzonego \x01, \ufffc)
	i pod kursorem znajduje się grafika, wypowiada pełne informacje o grafice zamiast 'pusta'.
	"""
	if not _isOutlookTreeInterceptor(treeInterceptorInstance):
		orig = getattr(treeInterceptorInstance, "_orig_speakTextInfo", None)
		if orig:
			orig(info, *args, **kwargs)
		return

	if GlobalPlugin.instance and time.time() < getattr(GlobalPlugin.instance, "_suppressCaretSpeechUntil", 0):
		return

	try:
		rawText = getattr(info, "text", "") or ""
		clean = rawText.strip()
		# Sprawdź, czy tekst linii to pusty ciąg lub znak obiektu osadzonego
		if not clean or clean in ("\x01", "\ufffc", "\u200b", "\ufeff"):
			doc = _getDocFromTreeInterceptorOrFocus(treeInterceptorInstance)
			if doc:
				targetPos = None
				try:
					if hasattr(info, "_rangeObj") and info._rangeObj:
						targetPos = int(info._rangeObj.Start)
				except Exception:
					pass

				item = getGraphicAtCaret(
					doc,
					strict=True,
					treeInterceptor=treeInterceptorInstance,
					targetPos=targetPos,
				)
				if item:
					playItemTypeSound(item, debounce_ms=100)
					try:
						import speech
						speech.cancelSpeech()
					except Exception:
						pass

					descParts = [f"Grafika {item.index}"]
					if getattr(item, "customName", ""):
						descParts.append(f": '{item.customName}'")
					elif getattr(item, "isBlankOrSolid", False):
						reason = getattr(item, "blankReason", "") or "pusty ozdobnik"
						descParts.append(f" ({reason})")
					elif getattr(item, "isDivider", False):
						descParts.append(" (linia podziału)")
					elif getattr(item, "isScreenshotOrUnlabeled", False):
						descParts.append(" (zrzut ekranu)")
					elif getattr(item, "altText", ""):
						descParts.append(f" ({item.altText})")
					else:
						descParts.append(f" ({item.widthPx} na {item.heightPx} px)")

					if getattr(item, "hyperlinkUrl", ""):
						dom = getCleanDomain(item.hyperlinkUrl)
						descParts.append(f", link: {dom}" if dom else ", link")

					descParts.append(". Shift+F10: czynności.")
					ui.message("".join(descParts))
					if GlobalPlugin.instance:
						GlobalPlugin.instance._currentGraphicItem = item
					return
	except Exception as ex:
		log.debug(f"outlookImageHelper: _treeInterceptor_speakTextInfo error: {ex}")

	# Jeśli odczytywany wiersz zawiera zwykły tekst, kursor nie stoi na grafice
	if GlobalPlugin.instance and getattr(info, "text", "") and any(ch.isalnum() for ch in str(info.text)):
		GlobalPlugin.instance._currentGraphicItem = None

	orig = getattr(treeInterceptorInstance, "_orig_speakTextInfo", None)
	if orig:
		orig(info, *args, **kwargs)


def _treeInterceptor_handleCaretMove(treeInterceptorInstance):
	"""Zastępuje komunikat 'pusta' przy poruszaniu się strzałkami po linii zawierającej grafikę."""
	if not _isOutlookTreeInterceptor(treeInterceptorInstance):
		orig = getattr(treeInterceptorInstance, "_orig_handleCaretMove", None)
		if orig:
			orig()
		return

	if GlobalPlugin.instance and time.time() < getattr(GlobalPlugin.instance, "_suppressCaretSpeechUntil", 0):
		return

	try:
		ti = getattr(treeInterceptorInstance, "selection", None)
		text = ti.text if ti else ""
		if not text or text.strip() in ("", "\x01", "\ufffc", "\u200b", "\ufeff"):
			doc = _getDocFromTreeInterceptorOrFocus(treeInterceptorInstance)
			if doc:
				item = getGraphicAtCaret(doc, strict=True, treeInterceptor=treeInterceptorInstance)
				if item:
					playItemTypeSound(item, debounce_ms=120)
					try:
						import speech
						speech.cancelSpeech()
					except Exception:
						pass
					descParts = [f"Grafika {item.index}"]
					if getattr(item, "customName", ""):
						descParts.append(f": '{item.customName}'")
					elif getattr(item, "isBlankOrSolid", False):
						reason = getattr(item, "blankReason", "") or "pusty ozdobnik"
						descParts.append(f" ({reason})")
					elif getattr(item, "isDivider", False):
						descParts.append(" (linia podziału)")
					elif getattr(item, "isScreenshotOrUnlabeled", False):
						descParts.append(" (zrzut ekranu)")
					elif getattr(item, "altText", ""):
						descParts.append(f" ({item.altText})")
					else:
						descParts.append(f" ({item.widthPx} na {item.heightPx} px)")

					if getattr(item, "hyperlinkUrl", ""):
						dom = getCleanDomain(item.hyperlinkUrl)
						descParts.append(f", link: {dom}" if dom else ", link")

					descParts.append(". Shift+F10: czynności.")
					ui.message("".join(descParts))
					if GlobalPlugin.instance:
						GlobalPlugin.instance._currentGraphicItem = item
					return
	except Exception:
		pass

	# Jeśli kursor przesunął się na wiersz z tekstem lub poza grafikę, wyczyść zapamiętany stan
	if GlobalPlugin.instance:
		GlobalPlugin.instance._currentGraphicItem = None

	orig = getattr(treeInterceptorInstance, "_orig_handleCaretMove", None)
	if orig:
		orig()


def _patchOutlookTreeInterceptors():
	"""
	Podłącza usprawnioną nawigację po grafikach (G oraz Shift+G),
	obsługę menu kontekstowego (Shift+F10 / Apps) oraz odczyt linii ze zrzutami/obrazami
	do klas przeglądania wiadomości WordMail i Outlook.
	"""
	classesToPatch = []
	try:
		import appModules.outlook
		if hasattr(appModules.outlook, "MailViewerTreeInterceptor"):
			classesToPatch.append(appModules.outlook.MailViewerTreeInterceptor)
	except Exception as ex:
		log.debug(f"outlookImageHelper: import appModules.outlook error: {ex}")

	try:
		import NVDAObjects.window.winword
		if hasattr(NVDAObjects.window.winword, "WordDocumentTreeInterceptor"):
			classesToPatch.append(NVDAObjects.window.winword.WordDocumentTreeInterceptor)
	except Exception as ex:
		log.debug(f"outlookImageHelper: import NVDAObjects.window.winword error: {ex}")

	try:
		import NVDAObjects.UIA.wordDocument
		if hasattr(NVDAObjects.UIA.wordDocument, "WordBrowseModeDocument"):
			classesToPatch.append(NVDAObjects.UIA.wordDocument.WordBrowseModeDocument)
	except Exception as ex:
		log.debug(f"outlookImageHelper: import NVDAObjects.UIA.wordDocument error: {ex}")

	for cls in classesToPatch:
		if not hasattr(cls, "_orig_script_nextGraphic"):
			cls._orig_script_nextGraphic = getattr(cls, "script_nextGraphic", None)
		if not hasattr(cls, "_orig_script_previousGraphic"):
			cls._orig_script_previousGraphic = getattr(cls, "script_previousGraphic", None)
		if not hasattr(cls, "_orig_script_passThrough"):
			cls._orig_script_passThrough = getattr(cls, "script_passThrough", None)
		if not hasattr(cls, "_orig_handleCaretMove"):
			cls._orig_handleCaretMove = getattr(cls, "handleCaretMove", None)
		if not hasattr(cls, "_orig_speakTextInfo"):
			cls._orig_speakTextInfo = getattr(cls, "speakTextInfo", None)

		cls.script_nextGraphic = _treeInterceptor_script_nextGraphic
		cls.script_previousGraphic = _treeInterceptor_script_previousGraphic
		cls.script_oih_contextMenu = _treeInterceptor_script_contextMenu
		if hasattr(cls, "handleCaretMove"):
			cls.handleCaretMove = _treeInterceptor_handleCaretMove
		if hasattr(cls, "speakTextInfo"):
			cls.speakTextInfo = _treeInterceptor_speakTextInfo

	log.info(f"outlookImageHelper: Podłączono obsługę G, Shift+G i nawigacji dla {len(classesToPatch)} klas bufora przeglądania.")


def _unpatchOutlookTreeInterceptors():
	"""Przywraca oryginalne metody nawigacji i zdarzeń przy wyłączaniu dodatku."""
	classesToRestore = []
	try:
		import appModules.outlook
		if hasattr(appModules.outlook, "MailViewerTreeInterceptor"):
			classesToRestore.append(appModules.outlook.MailViewerTreeInterceptor)
	except Exception:
		pass
	try:
		import NVDAObjects.window.winword
		if hasattr(NVDAObjects.window.winword, "WordDocumentTreeInterceptor"):
			classesToRestore.append(NVDAObjects.window.winword.WordDocumentTreeInterceptor)
	except Exception:
		pass
	try:
		import NVDAObjects.UIA.wordDocument
		if hasattr(NVDAObjects.UIA.wordDocument, "WordBrowseModeDocument"):
			classesToRestore.append(NVDAObjects.UIA.wordDocument.WordBrowseModeDocument)
	except Exception:
		pass

	for cls in classesToRestore:
		try:
			if hasattr(cls, "_orig_script_nextGraphic") and cls._orig_script_nextGraphic:
				cls.script_nextGraphic = cls._orig_script_nextGraphic
			if hasattr(cls, "_orig_script_previousGraphic") and cls._orig_script_previousGraphic:
				cls.script_previousGraphic = cls._orig_script_previousGraphic
			if hasattr(cls, "_orig_handleCaretMove") and cls._orig_handleCaretMove:
				cls.handleCaretMove = cls._orig_handleCaretMove
			if hasattr(cls, "_orig_speakTextInfo") and cls._orig_speakTextInfo:
				cls.speakTextInfo = cls._orig_speakTextInfo
			if hasattr(cls, "script_oih_contextMenu"):
				try:
					delattr(cls, "script_oih_contextMenu")
				except Exception:
					pass
			gm = getattr(cls, "_gestureMap", None)
			if isinstance(gm, dict):
				for k, v in list(gm.items()):
					if v == "oih_contextMenu":
						gm.pop(k, None)
		except Exception:
			pass



def _extractFromMapiAttachment(cid: str) -> str | None:
	"""Próbuje wyodrębnić oryginalny plik załącznika inline na podstawie Content-ID lub nazwy pliku bez odwoływania się do PropertyAccessor."""
	if not cid:
		return None
	try:
		cleanCid = cid.replace("cid:", "").strip("<>").strip()
		mailItem = getCurrentMailItem()
		if not mailItem:
			return None
		attachments = getattr(mailItem, "Attachments", None)
		if not attachments:
			return None

		count = int(attachments.Count)
		for a in range(1, count + 1):
			att = attachments.Item(a)
			fn = getattr(att, "FileName", "") or ""
			dn = getattr(att, "DisplayName", "") or ""
			if (fn and (cleanCid in fn or fn in cleanCid)) or (dn and (cleanCid in dn or dn in cleanCid)):
				tempDir = tempfile.gettempdir()
				tempPath = os.path.join(tempDir, f"nvda_oih_att_{int(time.time() * 1000)}.png")
				att.SaveAsFile(tempPath)
				log.info(f"outlookImageHelper: Zapisano załącznik wg nazwy '{fn}' do {tempPath}")
				return tempPath
	except Exception as ex:
		log.debug(f"outlookImageHelper: _extractFromMapiAttachment error: {ex}")
	return None


def checkImageBlankOrSolid(source) -> tuple[bool, str]:
	"""
	Błyskawicznie (w < 1 ms) bada, czy obraz jest pusty, całkowicie czarny, biały,
	przezroczysty, jednolity lub jest niewidocznym separatorem/spacerem.
	Przyjmuje obiekt PIL.Image lub ścieżkę do pliku graficznego.
	Zwraca (True, powód) jeśli obraz nie zawiera widocznej treści, lub (False, "").
	"""
	img = None
	shouldClose = False
	try:
		from PIL import Image, ImageStat
		if isinstance(source, str):
			if not os.path.isfile(source):
				return (False, "")
			fileSize = os.path.getsize(source)
			if fileSize < 130:
				return (True, "Jednolity lub pusty obraz (rozmiar pliku < 130 B)")
			img = Image.open(source)
			shouldClose = True
		else:
			img = source

		if img is None:
			return (False, "")

		w, h = img.size
		if w <= 3 or h <= 3:
			return (True, f"Mikro-ozdobnik/spacer ({w}×{h} px)")

		# 1. Sprawdzenie kanału alfa dla obrazów RGBA / LA / P
		mode = img.mode
		if mode in ("RGBA", "LA") or (mode == "P" and "transparency" in getattr(img, "info", {})):
			if mode == "P":
				img = img.convert("RGBA")
			extrema = img.getextrema()
			alphaExtrema = extrema[3] if len(extrema) == 4 else extrema[1]
			if alphaExtrema[1] <= 5:
				return (True, f"Całkowicie przezroczysty obraz/spacer ({w}×{h} px)")

		# 2. Sprawdzenie jednolitości kolorów
		rgbImg = img.convert("RGB") if mode != "RGB" else img
		extrema = rgbImg.getextrema()
		isPureSolid = (
			extrema[0][0] == extrema[0][1] and
			extrema[1][0] == extrema[1][1] and
			extrema[2][0] == extrema[2][1]
		)
		if isPureSolid:
			r, g, b = extrema[0][0], extrema[1][0], extrema[2][0]
			if r == 0 and g == 0 and b == 0:
				return (True, f"Całkowicie czarny obraz bez treści ({w}×{h} px)")
			elif r == 255 and g == 255 and b == 255:
				return (True, f"Całkowicie biały obraz/tło ({w}×{h} px)")
			else:
				return (True, f"Jednolity blok koloru RGB({r},{g},{b}) ({w}×{h} px)")

		# 3. Sprawdzenie niemal jednolitego koloru (szum, artefakty kompresji DIB / PNG)
		stat = ImageStat.Stat(rgbImg)
		maxStdDev = max(stat.stddev[:3]) if stat.stddev else 0.0
		if maxStdDev < 1.2:
			meanR, meanG, meanB = stat.mean[:3]
			if meanR < 8 and meanG < 8 and meanB < 8:
				return (True, f"Jednolity czarny obraz ({w}×{h} px)")
			elif meanR > 248 and meanG > 248 and meanB > 248:
				return (True, f"Jednolite białe tło ({w}×{h} px)")
			else:
				return (True, f"Jednolity ozdobnik ({w}×{h} px)")

		return (False, "")
	except Exception as ex:
		log.debug(f"outlookImageHelper: checkImageBlankOrSolid błąd: {ex}")
		return (False, "")
	finally:
		if shouldClose and img:
			try:
				img.close()
			except Exception:
				pass


def extractGraphicToPngFile(item: GraphicItem) -> str | None:
	"""
	Kopiuje grafikę z dokumentu Worda i zapisuje ją jako tymczasowy plik PNG.
	UWAGA: Funkcja MUSI być wywoływana z głównego wątku STA (MainThread).
	"""
	# Jeśli obraz został wcześniej pobrany z zewnętrznego serwera WWW lub wyodrębniony do sprawnego pliku
	if getattr(item, "_downloadedTempPath", None) and os.path.isfile(item._downloadedTempPath):
		return item._downloadedTempPath
	if getattr(item, "_extractedTempPath", None) and os.path.isfile(item._extractedTempPath):
		return item._extractedTempPath

	try:
		copied = False
		if item.isInline:
			try:
				# 1. Próba CopyAsPicture (pobiera obraz w pełnej, źródłowej rozdzielczości)
				item.rawObject.Range.CopyAsPicture()
				copied = True
			except Exception as ce0:
				log.debug(f"outlookImageHelper: Range.CopyAsPicture błąd: {ce0}")
				try:
					item.rawObject.Range.Copy()
					copied = True
				except Exception as ce1:
					log.debug(f"outlookImageHelper: Range.Copy błąd: {ce1}")
		else:
			try:
				# Bezpośrednia metoda Copy() na obiekcie Shape bez wywoływania Select()
				# Zabezpiecza przed deadlockiem COM / STA między NVDA a programem Outlook
				item.rawObject.Copy()
				copied = True
			except Exception as se0:
				log.debug(f"outlookImageHelper: Shape.Copy błąd: {se0}")

		if not copied:
			log.warning("outlookImageHelper: Nie udało się skopiować obiektu z dokumentu (brak obsługi Copy).")
			if item.cid:
				fallbackPath = _extractFromMapiAttachment(item.cid)
				if fallbackPath and os.path.isfile(fallbackPath):
					isBlank, blankReason = checkImageBlankOrSolid(fallbackPath)
					if isBlank:
						item.isBlankOrSolid = True
						item.blankReason = blankReason
						item._hasLocalRenderedPixels = False
					else:
						item.isBlankOrSolid = False
						item.blankReason = ""
						item._hasLocalRenderedPixels = True
						item._extractedTempPath = fallbackPath
					return fallbackPath
			return None

		img = None
		for attempt in range(4):
			if attempt > 0:
				time.sleep(0.05)
			try:
				from PIL import ImageGrab
				img = ImageGrab.grabclipboard()
				if img is not None:
					break
			except Exception as pe:
				log.debug(f"outlookImageHelper: ImageGrab próba {attempt} błąd: {pe}")

		if isinstance(img, list) and img:
			firstPath = img[0]
			if isinstance(firstPath, str) and os.path.isfile(firstPath):
				isBlank, blankReason = checkImageBlankOrSolid(firstPath)
				if isBlank:
					item.isBlankOrSolid = True
					item.blankReason = blankReason
					item._hasLocalRenderedPixels = False
				else:
					item.isBlankOrSolid = False
					item.blankReason = ""
					item._hasLocalRenderedPixels = True
					item._extractedTempPath = firstPath
				return firstPath

		if img is None:
			log.warning("outlookImageHelper: Nie udało się pobrać obrazu ze schowka. Próba MAPI...")
			if item.cid:
				fallbackPath = _extractFromMapiAttachment(item.cid)
				if fallbackPath and os.path.isfile(fallbackPath):
					isBlank, blankReason = checkImageBlankOrSolid(fallbackPath)
					if isBlank:
						item.isBlankOrSolid = True
						item.blankReason = blankReason
						item._hasLocalRenderedPixels = False
					else:
						item.isBlankOrSolid = False
						item.blankReason = ""
						item._hasLocalRenderedPixels = True
						item._extractedTempPath = fallbackPath
					return fallbackPath
			return None

		# Uaktualnij rzeczywiste wymiary obiektu z prawdziwego obrazu
		if hasattr(img, "width") and hasattr(img, "height"):
			item.widthPx = img.width
			item.heightPx = img.height

		tempDir = tempfile.gettempdir()
		tempPath = os.path.join(tempDir, f"nvda_oih_img_{int(time.time() * 1000)}.png")
		img.save(tempPath, "PNG")

		isBlank, blankReason = checkImageBlankOrSolid(img)
		if isBlank:
			item.isBlankOrSolid = True
			item.blankReason = blankReason
			item._hasLocalRenderedPixels = False
			log.info(f"outlookImageHelper: Grafika {item.index} zidentyfikowana jako pusta/ozdobnik: {blankReason}")
		else:
			item.isBlankOrSolid = False
			item.blankReason = ""
			item._hasLocalRenderedPixels = True
			item._extractedTempPath = tempPath

		log.info(f"outlookImageHelper: Zapisano tymczasowy obraz: {tempPath} ({img.width}x{img.height})")
		return tempPath
	except Exception as e:
		log.error(f"outlookImageHelper: Błąd podczas wyodrębniania grafiki: {e}", exc_info=True)
		if item.cid:
			fallbackPath = _extractFromMapiAttachment(item.cid)
			if fallbackPath and os.path.isfile(fallbackPath):
				isBlank, blankReason = checkImageBlankOrSolid(fallbackPath)
				if isBlank:
					item.isBlankOrSolid = True
					item.blankReason = blankReason
					item._hasLocalRenderedPixels = False
				else:
					item.isBlankOrSolid = False
					item.blankReason = ""
					item._hasLocalRenderedPixels = True
					item._extractedTempPath = fallbackPath
				return fallbackPath
		return None


def _is_attachment_inline_image(att):
	"""
	Precyzyjnie rozpoznaje, czy załącznik jest osadzony w treści (inline).
	Chroni zwykłe pliki (np. image_budget.pdf) i załączone maile .msg (Type == 5).
	"""
	try:
		fn = (getattr(att, "FileName", "") or "").lower()
		ext = os.path.splitext(fn)[1]
		cid = ""
		flags = 0
		try:
			pa = getattr(att, "PropertyAccessor", None)
			if pa:
				cid = str(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F") or "")
				flags = int(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003") or 0)
		except Exception:
			pass

		img_exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff", ".tif", ".ico", ".svg"}
		if (cid or (flags & 4)) and (ext in img_exts or not ext):
			return True

		attType = int(getattr(att, "Type", 1) or 1)
		if attType == 6 and (fn.startswith("oledata") or not ext):
			return True

		if ext in img_exts and re.match(r'^image\d{3,}\.(?:png|jpg|jpeg|gif|bmp)$', fn) and (cid or not fn):
			return True

		return False
	except Exception:
		return False


def extractAttachmentsFromCurrentMessage(mailItem=None, includeInline=False) -> list[AttachmentItem]:
	"""
	Pobiera listę załączników z bieżącej wiadomości.
	Jeśli podano mailItem, korzysta bezpośrednio z niego.
	W przeciwnym razie pobiera właściwy mailItem przez getCurrentMailItem().
	NIGDY nie pobiera załączników z losowego maila w tle!
	"""
	items: list[AttachmentItem] = []
	try:
		if not mailItem:
			mailItem = getCurrentMailItem()
		if not mailItem:
			return items

		attachments = getattr(mailItem, "Attachments", None)
		if not attachments:
			return items

		count = int(getattr(attachments, "Count", 0))
		realIdx = 1
		for i in range(1, count + 1):
			try:
				att = attachments.Item(i)
				fn = (getattr(att, "FileName", "") or "").strip()
				if not fn:
					continue
				sz = int(getattr(att, "Size", 0) or 0)
				fnLower = fn.lower()
				isInline = _is_attachment_inline_image(att)

				if not includeInline and isInline:
					continue

				item = AttachmentItem(
					index=realIdx,
					fileName=fn,
					sizeBytes=sz,
					rawAttachment=att,
				)
				if isInline:
					item.isInline = True
				ext = os.path.splitext(fnLower)[1]
				if ext in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff", ".ico", ".svg"):
					item.isImage = True

				items.append(item)
				realIdx += 1
			except Exception as ex:
				log.debug(f"outlookImageHelper: Błąd analizy załącznika {i}: {ex}")

	except Exception as e:
		log.debug(f"outlookImageHelper: extractAttachmentsFromCurrentMessage błąd ogólny: {e}")

	return items


_active_temp_attachment_folders = set()


def saveAttachmentToTempFile(item: AttachmentItem) -> str:
	"""
	Zapisuje załącznik do dedykowanego, wyizolowanego podkatalogu tymczasowego.
	Zapobiega kolizjom nazw między różnymi wiadomościami i chroni przed podmianą plików.
	"""
	if item.tempSavedPath and os.path.isfile(item.tempSavedPath):
		return item.tempSavedPath

	baseTempDir = os.path.join(tempfile.gettempdir(), "nvda_oih_attachments")
	try:
		os.makedirs(baseTempDir, exist_ok=True)
	except Exception:
		baseTempDir = tempfile.gettempdir()

	# Unikalny podkatalog dla każdego obiektu/operacji zapobiegający kolizjom
	uniqueSubdir = os.path.join(baseTempDir, f"att_{uuid.uuid4().hex[:12]}")
	try:
		os.makedirs(uniqueSubdir, exist_ok=True)
		_active_temp_attachment_folders.add(uniqueSubdir)
	except Exception:
		uniqueSubdir = baseTempDir

	# Bezpieczna sanitizacja bez mutacji rozszerzenia pliku
	rawFn = (item.fileName or "").strip()
	stem, ext = os.path.splitext(rawFn)
	safeStem = "".join(c if (c.isalnum() or c in " ._-()") else "_" for c in stem).strip().rstrip(". ") or f"zalacznik_{item.index}"
	safeExt = "".join(c if (c.isalnum() or c in "._-") else "_" for c in ext).strip()
	safeName = f"{safeStem}{safeExt}"
	targetPath = os.path.join(uniqueSubdir, f"{item.index}_{safeName}")

	try:
		item.rawAttachment.SaveAsFile(targetPath)
		item.tempSavedPath = targetPath
		log.info(f"outlookImageHelper: Zapisano załącznik #{item.index} do pliku tymczasowego ({item.formattedSize})")
		return targetPath
	except Exception as ex:
		log.error(f"outlookImageHelper: Nie udało się zapisać załącznika #{item.index}: {ex}", exc_info=True)
		return ""


def extractTextFromDocxFile(filePath: str) -> str:
	"""Wyciąga czysty tekst z dokumentu Word (DOCX) bez uruchamiania aplikacji Word."""
	if not os.path.isfile(filePath):
		return ""
	try:
		with zipfile.ZipFile(filePath, "r") as z:
			if "word/document.xml" not in z.namelist():
				return ""
			info = z.getinfo("word/document.xml")
			if info.file_size > 25 * 1024 * 1024:
				log.warning(f"outlookImageHelper: Plik word/document.xml przekracza bezpieczny rozmiar 25 MB ({info.file_size} B).")
				return ""
			xmlContent = z.read("word/document.xml")
			tree = ET.fromstring(xmlContent)
			paragraphs = []
			body = None
			for child in tree.iter():
				if child.tag.endswith("body"):
					body = child
					break
			if body is None:
				body = tree

			for elem in body:
				tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
				if tag == "p":
					pTexts = [child.text for child in elem.iter() if child.text and (child.tag.endswith("t") or child.tag.endswith("delText"))]
					if pTexts:
						paragraphs.append("".join(pTexts))
				elif tag == "tbl":
					for tr in elem.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr"):
						cells = []
						for tc in tr.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc"):
							cTexts = [child.text for child in tc.iter() if child.text and (child.tag.endswith("t") or child.tag.endswith("delText"))]
							cells.append(" ".join("".join(cTexts).split()))
						if any(cells):
							paragraphs.append(" | ".join(cells))
				elif tag == "sdt":
					for p in elem.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
						pTexts = [child.text for child in p.iter() if child.text and (child.tag.endswith("t") or child.tag.endswith("delText"))]
						if pTexts:
							paragraphs.append("".join(pTexts))

			return "\n\n".join(paragraphs).strip()
	except Exception as e:
		log.debug(f"outlookImageHelper: extractTextFromDocxFile błąd: {e}")
		return ""


def isPdfEncrypted(pdfPath: str) -> bool:
	"Sprawdza, czy plik PDF jest zaszyfrowany/chroniony hasłem."
	if not pdfPath or not os.path.isfile(pdfPath):
		return False
	try:
		with open(pdfPath, 'rb') as f:
			chunk = f.read(1024 * 1024)
			if b'/Encrypt' in chunk:
				return True
	except Exception as ex:
		log.error(f'outlookImageHelper: błąd odczytu binarnego PDF: {ex}')
	if pypdf is not None:
		try:
			reader = pypdf.PdfReader(pdfPath)
			if reader.is_encrypted:
				return True
		except Exception as ex:
			log.error(f'outlookImageHelper: błąd pypdf reader.is_encrypted: {ex}')
	return False


def unlockPdfFile(pdfPath: str, password: str = "") -> tuple[bool, str]:
	"""
	Odszyfrowuje plik PDF podanym hasłem i zapisuje wersję odbezpieczoną do dedykowanego pliku tymczasowego.
	Zwraca (sukces: bool, ścieżka_lub_błąd: str).
	"""
	if not pdfPath or not os.path.isfile(pdfPath):
		return False, "Plik PDF nie istnieje."
	if pypdf is None:
		return False, "Brak biblioteki pypdf do obsługi zaszyfrowanych plików PDF."
	try:
		reader = pypdf.PdfReader(pdfPath)
		if not reader.is_encrypted:
			return True, pdfPath

		res = reader.decrypt(password)
		pw_type = getattr(pypdf, "PasswordType", None)
		not_decrypted = getattr(pw_type, "NOT_DECRYPTED", 0) if pw_type else 0
		if res == not_decrypted or int(res) == 0:
			return False, "Podane hasło jest niepoprawne."

		unlockedDir = os.path.dirname(pdfPath)
		baseName = os.path.splitext(os.path.basename(pdfPath))[0]
		unlockedPath = os.path.join(unlockedDir, f"{baseName}_unlocked.pdf")

		writer = pypdf.PdfWriter()
		writer.append(reader)
		with open(unlockedPath, "wb") as f_out:
			writer.write(f_out)

		log.info(f"outlookImageHelper: Plik PDF pomyślnie odblokowany i zapisany do: {unlockedPath}")
		return True, unlockedPath
	except Exception as ex:
		log.error(f"outlookImageHelper: Błąd podczas odszyfrowywania PDF: {ex}", exc_info=True)
		return False, str(ex)


class PdfPasswordDialog(wx.Dialog):
	"""
	W pełni dostępne (WCAG 2.1/2.2 AA) okno dialogowe z prośbą o hasło
	do zabezpieczonego dokumentu PDF.
	"""
	def __init__(self, parent, fileName: str, pdfPath: str):
		super().__init__(
			parent,
			title="Hasło do dokumentu PDF",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.fileName = fileName
		self.pdfPath = pdfPath
		self.unlockedPath = ""
		self.enteredPassword = ""

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# Informacja o pliku
		infoMsg = (
			f"Dokument PDF „{self.fileName}” jest zabezpieczony hasłem.\n\n"
			"Podaj hasło do pliku, aby umożliwić jego odczyt, analizę "
			"i podsumowanie przez sztuczną inteligencję (AI):"
		)
		self.lblInfo = wx.StaticText(self, label=infoMsg)
		mainSizer.Add(self.lblInfo, 0, wx.ALL | wx.EXPAND, 15)

		# Pole wprowadzania hasła
		lblPwd = wx.StaticText(self, label="&Hasło do pliku PDF:")
		mainSizer.Add(lblPwd, 0, wx.LEFT | wx.RIGHT | wx.TOP, 15)

		self.pwdCtrl = wx.TextCtrl(self, style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER, size=(400, -1))
		self.pwdCtrl.SetName("Hasło do pliku PDF")
		mainSizer.Add(self.pwdCtrl, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 15)

		# Przyciski
		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnUnlock = wx.Button(self, wx.ID_OK, label="&Odblokuj i kontynuuj")
		self.btnUnlock.SetDefault()
		btnSizer.Add(self.btnUnlock, 0, wx.ALL, 5)

		self.btnCancel = wx.Button(self, wx.ID_CANCEL, label="&Anuluj")
		btnSizer.Add(self.btnCancel, 0, wx.ALL, 5)

		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 12)

		self.SetSizer(mainSizer)
		self.Fit()
		self.CenterOnScreen()

		# Zdarzenia
		self.btnUnlock.Bind(wx.EVT_BUTTON, self.onUnlock)
		self.pwdCtrl.Bind(wx.EVT_TEXT_ENTER, self.onUnlock)
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

		wx.CallAfter(self.pwdCtrl.SetFocus)

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		else:
			event.Skip()

	def onUnlock(self, event):
		pwd = self.pwdCtrl.GetValue()
		if not pwd:
			showAccessibleMessageBox(
				'Pole hasła jest puste.\n\nWprowadź hasło do dokumentu PDF, aby kontynuować, lub wybierz Anuluj.',
				title='Brak hasła',
				parent=self,
			)
			self.pwdCtrl.SetFocus()
			return

		ok, result = unlockPdfFile(self.pdfPath, pwd)
		if not ok:
			showAccessibleMessageBox(
				f'Nie udało się odblokować dokumentu PDF.\n\nPowód:\n{result}\n\nSprawdź poprawność hasła i spróbuj ponownie, lub zamknij to okno wybierając Anuluj.',
				title='Błąd hasła do dokumentu PDF',
				parent=self,
			)
			self.pwdCtrl.SelectAll()
			self.pwdCtrl.SetFocus()
			return

		self.unlockedPath = result
		self.enteredPassword = pwd
		self.EndModal(wx.ID_OK)
def ensurePdfUnlocked(item: AttachmentItem, parentWin=None) -> str | None:
	"""
	Upewnia się, że dokument PDF jest odbezpieczony.
	Jeśli plik jest zaszyfrowany i wymaga hasła, wyświetla przystępne okno dialogowe z prośbą o hasło.
	Zwraca ścieżkę do odbezpieczonego pliku tymczasowego lub None w przypadku anulowania.
	"""
	if not getattr(item, "isPdf", False):
		return getattr(item, "tempSavedPath", "")

	# Jeśli już wcześniej został odblokowany w bieżącej sesji:
	curUnlocked = getattr(item, "unlockedPdfPath", "")
	if curUnlocked and os.path.isfile(curUnlocked):
		return curUnlocked

	targetPath = getattr(item, "tempSavedPath", "")
	if not targetPath or not os.path.isfile(targetPath):
		targetPath = saveAttachmentToTempFile(item)
		if not targetPath or not os.path.isfile(targetPath):
			ui.message(f"Błąd: Nie udało się pobrać pliku załącznika {item.fileName}.")
			return None

	if not isPdfEncrypted(targetPath):
		return targetPath

	# Próba automatycznego odblokowania (puste hasło otwarcia)
	ok, res = unlockPdfFile(targetPath, "")
	if ok and res and os.path.isfile(res):
		item.unlockedPdfPath = res
		return res

	# Plik wymaga podania hasła przez użytkownika
	parent = parentWin or gui.mainFrame
	gui.mainFrame.prePopup()
	try:
		dlg = PdfPasswordDialog(parent, fileName=item.fileName, pdfPath=targetPath)
		ret = dlg.ShowModal()
		if ret == wx.ID_OK and dlg.unlockedPath:
			item.unlockedPdfPath = dlg.unlockedPath
			item.pdfPassword = dlg.enteredPassword
			return dlg.unlockedPath
		else:
			ui.message(f"Anulowano analizę zabezpieczonego dokumentu {item.fileName}.")
			return None
	finally:
		dlg.Destroy()
		gui.mainFrame.postPopup()


_cached_ocr_script: str | None = None


def _getOcrScriptContent() -> str:
	"""Zwraca treść skryptu PowerShell dla silnika OCR z pamięci podręcznej RAM."""
	global _cached_ocr_script
	if _cached_ocr_script is None:
		if not os.path.isfile(OCR_SCRIPT_PATH):
			raise FileNotFoundError(f"Nie odnaleziono skryptu OCR: {OCR_SCRIPT_PATH}")
		with open(OCR_SCRIPT_PATH, "r", encoding="utf-8") as f:
			_cached_ocr_script = f.read()
	return _cached_ocr_script


def runWindowsOcr(imagePath: str, lang: str = "pl-PL") -> str:
	"""
	Uruchamia systemowy Windows Media OCR w skrypcie ocr_worker.ps1 wykonywanym
	bezpośrednio w pamięci procesora PowerShell (-EncodedCommand).
	Dzięki temu skrypt działa niezawodnie na każdym komputerze, nawet przy restrykcyjnych
	zasadach ExecutionPolicy (Restricted, AllSigned, RemoteSigned) i braku podpisu cyfrowego.
	"""
	with BusySoundTicker(startTone=True, interval=0.28, baseFreq=620, endTone=True):
		script_content = _getOcrScriptContent()

		escaped_path = imagePath.replace("'", "''")
		escaped_lang = (lang or "pl-PL").replace("'", "''")
		cmd_text = f"""$ProgressPreference = 'SilentlyContinue'
& {{
{script_content}
}} -ImagePath '{escaped_path}' -Lang '{escaped_lang}'
"""
		encoded_cmd = base64.b64encode(cmd_text.encode("utf-16le")).decode("ascii")

		powershell_exe = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
		if not os.path.isfile(powershell_exe):
			powershell_exe = "powershell.exe"

		cmd = [
			powershell_exe,
			"-NoProfile",
			"-NonInteractive",
			"-ExecutionPolicy",
			"Bypass",
			"-EncodedCommand",
			encoded_cmd,
		]

		startupinfo = subprocess.STARTUPINFO()
		startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
		startupinfo.wShowWindow = 0

		proc = subprocess.Popen(
			cmd,
			stdout=subprocess.PIPE,
			stderr=subprocess.PIPE,
			startupinfo=startupinfo,
			creationflags=0x08000000,
		)

		try:
			stdout, stderr = proc.communicate(timeout=25)
		except subprocess.TimeoutExpired:
			try:
				proc.kill()
			except Exception:
				pass
			try:
				proc.communicate()
			except Exception:
				pass
			raise TimeoutError("Przekroczono limit czasu (25 s) operacji OCR systemu Windows.")

		if proc.returncode != 0:
			errText = stderr.decode("utf-8", errors="replace").strip()
			if "#< CLIXML" in errText:
				clean_lines = [l for l in errText.splitlines() if not l.startswith("<") and not l.startswith("#<")]
				if clean_lines:
					errText = "\n".join(clean_lines).strip()
			raise RuntimeError(f"Błąd silnika OCR (kod {proc.returncode}): {errText}")

		output = stdout.decode("utf-8", errors="replace").strip()
		if output == "[BRAK_TEKSTU]":
			return ""
		return output


_RETIRED_GEMINI_MODELS = {
	"gemini-1.5-flash",
	"gemini-1.5-pro",
	"gemini-2.0-flash",
	"gemini-2.5-flash",
	"gemini-2.5-flash-lite",
	"gemini-3.5-flash",
	"gemini-3.5-flash-lite",
	"gemini-3.6-flash",
	"gemini-3.7-flash",
}
_DEFAULT_GEMINI_MODEL = "gemini-3.8-flash"
_FALLBACK_GEMINI_MODELS = (
	"gemini-3.8-flash",
	"gemini-flash-lite-latest",
	"gemini-flash-latest",
)


def _getGeminiCandidateModels(preferredModel: str = "") -> list[str]:
	"""Zwraca uporządkowaną listę aktywnych modeli Gemini z automatyczną migracją wycofanych nazw (np. gemini-2.0-flash)."""
	m = (preferredModel or "").strip()
	if not m or m in _RETIRED_GEMINI_MODELS:
		m = _DEFAULT_GEMINI_MODEL
	res = [m]
	for fb in _FALLBACK_GEMINI_MODELS:
		if fb not in res:
			res.append(fb)
	return res


def describeWithGemini(
	imagePath: str = "",
	apiKey: str = "",
	model: str = "gemini-3.8-flash",
	customPrompt: str = "",
	contextText: str = "",
	mimeType: str = "image/png",
	prefaceText: str = "",
	base64Data: str | None = None,
) -> str:
	"""Wysyła obraz lub załączony dokument (PDF/obraz) do API Google Gemini i zwraca szczegółowy opis dla osoby niewidomej."""
	if not apiKey:
		raise ValueError(
			"Brak klucza API Google Gemini. Skonfiguruj go w oknie 'Klucz AI' lub ustawieniach dodatku."
		)

	imgData = base64Data or ""
	if not imgData and imagePath and os.path.isfile(imagePath):
		with open(imagePath, "rb") as f:
			imgData = base64.b64encode(f.read()).decode("ascii")

	lang = getEffectiveLanguage()
	if customPrompt:
		prompt = customPrompt.strip()
	else:
		if lang == "en":
			contextHint = (
				f"Context from the email message body: \"{contextText}\".\n"
				if contextText else ""
			)
			prompt = (
				"You are a precise assistant for a blind person. "
				"Thoroughly and comprehensively describe and analyze the attached visual/document material in English.\n"
				f"{contextHint}"
				"- If it is a software screenshot, UI, or table: connect visible data with email context, specify column headers, row values, and conclusions.\n"
				"- If it is a PDF document: specify subject, sender, recipient, dates, amounts, tabular breakdown, and summary.\n"
				"- If it is a chart: describe axes, values, and trends.\n"
				"- If it is a photo: describe scene, objects, people, and colors."
			)
		else:
			contextHint = (
				f"Kontekst z treści wiadomości e-mail: \"{contextText}\".\n"
				if contextText else ""
			)
			prompt = (
				"Jesteś precyzyjnym asystentem osoby niewidomej. "
				"Dokładnie i wyczerpująco przeanalizuj załączony materiał w języku polskim.\n"
				f"{contextHint}"
				"- Jeśli to zrzut ekranu lub tabela: połącz widoczne dane, podaj nagłówki kolumn, wartości i wnioski.\n"
				"- Jeśli to dokument PDF: wskaż temat, nadawcę, odbiorcę, daty, kwoty, tabelaryczne zestawienia i podsumowanie.\n"
				"- Jeśli to wykres: opisz osie, wartości i trendy.\n"
				"- Jeśli to zdjęcie: opisz scenę, obiekty, ludzi i kolory."
			)

	if prefaceText:
		prompt = f"{prefaceText}\n\n{prompt}"

	parts = [{"text": prompt}]
	if imgData:
		parts.append({
			"inline_data": {
				"mime_type": mimeType,
				"data": imgData,
			}
		})

	payload = {
		"contents": [
			{
				"parts": parts
			}
		]
	}

	modelsToTry = _getGeminiCandidateModels(model)
	lastError = None
	payloadBytes = json.dumps(payload).encode("utf-8")

	for candidateModel in modelsToTry:
		url = f"https://generativelanguage.googleapis.com/v1beta/models/{candidateModel}:generateContent"
		req = urllib.request.Request(
			url,
			data=payloadBytes,
			headers={
				"Content-Type": "application/json",
				"x-goog-api-key": apiKey,
			},
		)
		try:
			with urllib.request.urlopen(req, timeout=28) as resp:
				resData = json.loads(resp.read().decode("utf-8"))

			candidates = resData.get("candidates", [])
			if not candidates:
				raise RuntimeError("Model AI nie zwrócił żadnej odpowiedzi.")

			resParts = candidates[0].get("content", {}).get("parts", [])
			texts = [p.get("text", "") for p in resParts if "text" in p]
			result = "\n".join(texts).strip()
			if result:
				log.info(f"outlookImageHelper: Pomyślnie wygenerowano opis AI modelem '{candidateModel}'.")
				return result
		except urllib.error.HTTPError as he:
			lastError = he
			try:
				errBody = he.read().decode("utf-8", errors="replace")
				errJson = json.loads(errBody)
				apiMsg = errJson.get("error", {}).get("message", "")
				if apiMsg:
					log.warning(f"outlookImageHelper: Gemini API HTTP {he.code} szczegóły: {apiMsg}")
					lastError = RuntimeError(f"HTTP Error {he.code}: {apiMsg}")
			except Exception:
				pass
			if he.code in (404, 429, 500, 502, 503, 504):
				log.warning(f"outlookImageHelper: Model '{candidateModel}' zwrócił HTTP {he.code}. Próba kolejnego...")
				continue
			raise lastError
		except Exception as e:
			lastError = e
			log.warning(f"outlookImageHelper: Błąd zapytania dla modelu '{candidateModel}': {e}.")
			continue

	if lastError:
		raise lastError
	raise RuntimeError("Nie udało się uzyskać odpowiedzi od żadnego modelu Gemini.")


def callGeminiVisionApi(
	prompt: str,
	imagePath: str | None = None,
	base64Data: str | None = None,
	mimeType: str = "image/png",
	apiKey: str | None = None,
	model: str = "gemini-3.8-flash",
) -> str:
	"""Wywołuje model wizyjny Google Gemini z obsługą obrazu na dysku lub zakodowanego w base64."""
	return describeWithGemini(
		imagePath=imagePath or "",
		apiKey=apiKey or getActiveAiApiKey(),
		model=model,
		customPrompt=prompt,
		mimeType=mimeType,
		base64Data=base64Data,
	)


def askGeminiMultiTurn(
	chatHistory: list[dict],
	apiKey: str,
	model: str = "gemini-3.8-flash",
) -> str:
	"""
	Wysyła historię wieloturowej rozmowy (wraz z obrazem/dokumentem w pierwszej turze)
	do API Google Gemini i zwraca odpowiedź modelu.
	"""
	if not apiKey:
		raise ValueError("Brak klucza API Google Gemini.")

	contents = []
	for turn in chatHistory:
		parts = []
		textVal = turn.get("text", "").strip()
		if textVal:
			parts.append({"text": textVal})
		imgB64 = turn.get("imageData")
		mimeType = turn.get("mimeType", "image/png")
		if imgB64:
			parts.append({
				"inline_data": {
					"mime_type": mimeType,
					"data": imgB64,
				}
			})
		if parts:
			contents.append({
				"role": turn.get("role", "user"),
				"parts": parts,
			})

	payload = {"contents": contents}
	payloadBytes = json.dumps(payload).encode("utf-8")

	modelsToTry = _getGeminiCandidateModels(model)
	lastError = None
	for candidateModel in modelsToTry:
		url = f"https://generativelanguage.googleapis.com/v1beta/models/{candidateModel}:generateContent"
		req = urllib.request.Request(
			url,
			data=payloadBytes,
			headers={
				"Content-Type": "application/json",
				"x-goog-api-key": apiKey,
			},
		)
		try:
			with urllib.request.urlopen(req, timeout=30) as resp:
				resData = json.loads(resp.read().decode("utf-8"))

			candidates = resData.get("candidates", [])
			if not candidates:
				raise RuntimeError("Model AI nie zwrócił żadnej odpowiedzi.")

			resParts = candidates[0].get("content", {}).get("parts", [])
			texts = [p.get("text", "") for p in resParts if "text" in p]
			result = "\n".join(texts).strip()
			if result:
				log.info(f"outlookImageHelper: Odpowiedź w wątku wieloturowym z modelu '{candidateModel}'.")
				return result
		except urllib.error.HTTPError as he:
			lastError = he
			try:
				errBody = he.read().decode("utf-8", errors="replace")
				errJson = json.loads(errBody)
				apiMsg = errJson.get("error", {}).get("message", "")
				if apiMsg:
					log.warning(f"outlookImageHelper: Gemini API HTTP {he.code} szczegóły: {apiMsg}")
					lastError = RuntimeError(f"HTTP Error {he.code}: {apiMsg}")
			except Exception:
				pass
			if he.code in (404, 429, 500, 502, 503, 504):
				log.warning(f"outlookImageHelper: Model '{candidateModel}' zwrócił HTTP {he.code}. Próba kolejnego...")
				continue
			raise lastError
		except Exception as e:
			lastError = e
			log.warning(f"outlookImageHelper: Błąd zapytania dla modelu '{candidateModel}': {e}.")
			continue

	if lastError:
		raise lastError
	raise RuntimeError("Nie udało się uzyskać odpowiedzi od żadnego modelu Gemini.")


# ==============================================================================
# OBSŁUGA OPENAI (CHATGPT) ORAZ UNIFIKACJA DOSTAWCÓW AI
# ==============================================================================

def getActiveAiProvider() -> str:
	try:
		sec = config.conf.get(CONFIG_SECTION, {})
		prov = str(sec.get("aiProvider", "gemini")).strip().lower()
		return "openai" if prov == "openai" else "gemini"
	except Exception:
		return "gemini"


def getActiveAiModel() -> str:
	try:
		sec = config.conf.get(CONFIG_SECTION, {})
		prov = getActiveAiProvider()
		if prov == "openai":
			return str(sec.get("openaiModel", "gpt-4o-mini") or "gpt-4o-mini").strip()
		m = str(sec.get("geminiModel", _DEFAULT_GEMINI_MODEL) or _DEFAULT_GEMINI_MODEL).strip()
		if m in _RETIRED_GEMINI_MODELS:
			m = _DEFAULT_GEMINI_MODEL
		return m
	except Exception:
		return _DEFAULT_GEMINI_MODEL


def getActiveAiApiKey() -> str:
	try:
		sec = config.conf.get(CONFIG_SECTION, {})
		prov = getActiveAiProvider()
		if prov == "openai":
			return str(sec.get("openaiApiKey", "") or "").strip()
		return str(sec.get("geminiApiKey", "") or "").strip()
	except Exception:
		return ""


def callOpenAiVisionApi(
	prompt: str,
	imagePath: str | None = None,
	base64Data: str | None = None,
	mimeType: str = "image/png",
	apiKey: str | None = None,
	model: str = "gpt-4o-mini",
	timeout: int = 35,
	maxTokens: int = 1500,
) -> str:
	"""Wykonuje zapytanie do API OpenAI (ChatGPT) z dołączonym obrazem lub dokumentem."""
	key = (apiKey or getActiveAiApiKey()).strip()
	if not key:
		raise ValueError("Brak klucza API OpenAI (ChatGPT). Skonfiguruj go w ustawieniach dodatku.")

	content = []
	if prompt:
		content.append({"type": "text", "text": prompt})

	b64 = base64Data
	if not b64 and imagePath and os.path.isfile(imagePath):
		with open(imagePath, "rb") as f:
			b64 = base64.b64encode(f.read()).decode("ascii")

	if b64:
		if mimeType == "application/pdf" or (imagePath and imagePath.lower().endswith(".pdf")):
			fn = os.path.basename(imagePath) if imagePath else "document.pdf"
			content.append({
				"type": "file",
				"file": {
					"filename": fn,
					"file_data": f"data:application/pdf;base64,{b64}"
				}
			})
		else:
			content.append({
				"type": "image_url",
				"image_url": {
					"url": f"data:{mimeType};base64,{b64}",
					"detail": "high"
				}
			})

	payload = {
		"model": model or "gpt-4o-mini",
		"messages": [{"role": "user", "content": content}],
		"max_tokens": maxTokens,
	}
	payloadBytes = json.dumps(payload).encode("utf-8")

	req = urllib.request.Request(
		"https://api.openai.com/v1/chat/completions",
		data=payloadBytes,
		headers={
			"Content-Type": "application/json",
			"Authorization": f"Bearer {key}",
		},
	)

	lastErr = None
	for attempt in range(3):
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				resData = json.loads(resp.read().decode("utf-8"))
			choices = resData.get("choices", [])
			if not choices:
				raise RuntimeError("Model OpenAI nie zwrócił żadnej odpowiedzi.")
			msgContent = choices[0].get("message", {}).get("content", "")
			return (msgContent or "").strip()
		except urllib.error.HTTPError as he:
			lastErr = he
			if he.code in (429, 502, 503, 504) and attempt < 2:
				time.sleep(1.8 * (attempt + 1))
				continue
			raise he
		except Exception as ex:
			lastErr = ex
			if attempt < 2:
				time.sleep(1.5)
				continue
			raise ex

	if lastErr:
		raise lastErr
	raise RuntimeError("Brak odpowiedzi od API OpenAI.")


def describeWithOpenAi(
	filePath: str,
	apiKey: str | None = None,
	model: str = "gpt-4o-mini",
	customPrompt: str | None = None,
	mimeType: str = "image/png",
	timeout: int = 40,
) -> str:
	"""Adapter wywołujący OpenAI Chat Completions dla pliku obrazu lub dokumentu PDF."""
	prompt = customPrompt or "Dokładnie opisz ten plik."
	return callOpenAiVisionApi(
		prompt=prompt,
		imagePath=filePath,
		mimeType=mimeType,
		apiKey=apiKey,
		model=model,
		timeout=timeout,
	)


def askOpenAiMultiTurn(
	chatHistory: list[dict],
	apiKey: str | None = None,
	model: str = "gpt-4o-mini",
	timeout: int = 40,
) -> str:
	"""Obsługuje wieloturowy czat konwersacyjny z modelami OpenAI (ChatGPT)."""
	key = (apiKey or getActiveAiApiKey()).strip()
	if not key:
		raise ValueError("Brak klucza API OpenAI (ChatGPT).")

	messages = []
	for turn in chatHistory:
		role = "assistant" if turn.get("role") == "model" else "user"
		textVal = turn.get("text", "").strip()
		imgB64 = turn.get("imageData")
		mimeType = turn.get("mimeType", "image/png")

		if imgB64 and role == "user":
			content = []
			if textVal:
				content.append({"type": "text", "text": textVal})
			if mimeType == "application/pdf":
				content.append({
					"type": "file",
					"file": {
						"filename": turn.get("fileName", "document.pdf"),
						"file_data": f"data:application/pdf;base64,{imgB64}"
					}
				})
			else:
				content.append({
					"type": "image_url",
					"image_url": {
						"url": f"data:{mimeType};base64,{imgB64}",
						"detail": "high"
					}
				})
			messages.append({"role": role, "content": content})
		else:
			messages.append({"role": role, "content": textVal})

	payload = {
		"model": model or "gpt-4o-mini",
		"messages": messages,
		"max_tokens": 1800,
	}
	payloadBytes = json.dumps(payload).encode("utf-8")

	req = urllib.request.Request(
		"https://api.openai.com/v1/chat/completions",
		data=payloadBytes,
		headers={
			"Content-Type": "application/json",
			"Authorization": f"Bearer {key}",
		},
	)

	lastErr = None
	for attempt in range(3):
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				resData = json.loads(resp.read().decode("utf-8"))
			choices = resData.get("choices", [])
			if not choices:
				raise RuntimeError("Model OpenAI nie zwrócił żadnej odpowiedzi.")
			msgContent = choices[0].get("message", {}).get("content", "")
			return (msgContent or "").strip()
		except urllib.error.HTTPError as he:
			lastErr = he
			if he.code in (429, 502, 503, 504) and attempt < 2:
				time.sleep(1.8 * (attempt + 1))
				continue
			raise he
		except Exception as ex:
			lastErr = ex
			if attempt < 2:
				time.sleep(1.5)
				continue
			raise ex

	if lastErr:
		raise lastErr
	raise RuntimeError("Brak odpowiedzi od API OpenAI.")


def testOpenAiApiKey(apiKey: str, model: str = "gpt-4o-mini") -> tuple[bool, str]:
	"""Testuje poprawność klucza API OpenAI."""
	key = (apiKey or "").strip()
	if not key:
		return False, "Klucz API OpenAI jest pusty."

	payload = {
		"model": model or "gpt-4o-mini",
		"messages": [{"role": "user", "content": "Hello"}],
		"max_tokens": 5,
	}
	req = urllib.request.Request(
		"https://api.openai.com/v1/chat/completions",
		data=json.dumps(payload).encode("utf-8"),
		headers={
			"Content-Type": "application/json",
			"Authorization": f"Bearer {key}",
		},
	)
	try:
		with urllib.request.urlopen(req, timeout=12) as resp:
			resData = json.loads(resp.read().decode("utf-8"))
		if "choices" in resData:
			return True, f"Klucz API OpenAI jest prawidłowy! Połączono pomyślnie z modelem '{model}'."
	except urllib.error.HTTPError as he:
		if he.code == 401:
			return False, "Nieprawidłowy klucz API OpenAI (HTTP 401 Unauthorized). Sprawdź czy klucz został poprawnie skopiowany."
		elif he.code == 429:
			return True, "Klucz API jest poprawny, lecz wystąpił limit zapytań lub brak środków na koncie OpenAI (HTTP 429 Too Many Requests)."
		return False, f"Błąd serwera OpenAI: HTTP {he.code}"
	except Exception as ex:
		return False, f"Błąd sieciowy podczas testu OpenAI: {ex}"
	return False, "Brak odpowiedzi od serwera OpenAI."


def callAiVisionApi(
	prompt: str,
	imagePath: str | None = None,
	base64Data: str | None = None,
	mimeType: str = "image/png",
	parentDialog=None,
) -> str:
	"""Uniwersalna funkcja wywołująca zapytanie wizyjne do wybranego w konfiguracji dostawcy (Gemini lub OpenAI)."""
	prov = getActiveAiProvider()
	key = getActiveAiApiKey()
	if not key:
		provName = "OpenAI (ChatGPT)" if prov == "openai" else "Google Gemini"
		raise ValueError(f"Brak klucza API dla usługi {provName}. Skonfiguruj klucz w ustawieniach dodatku.")
	model = getActiveAiModel()
	if prov == "openai":
		return callOpenAiVisionApi(prompt, imagePath=imagePath, base64Data=base64Data, mimeType=mimeType, apiKey=key, model=model)
	else:
		return callGeminiVisionApi(prompt, imagePath=imagePath, base64Data=base64Data, mimeType=mimeType, apiKey=key, model=model)


def askAiMultiTurn(chatHistory: list[dict], parentDialog=None) -> str:
	"""Uniwersalna funkcja obsługująca wieloturową rozmowę dla aktywnego dostawcy AI."""
	prov = getActiveAiProvider()
	key = getActiveAiApiKey()
	if not key:
		provName = "OpenAI (ChatGPT)" if prov == "openai" else "Google Gemini"
		raise ValueError(f"Brak klucza API dla usługi {provName}. Skonfiguruj klucz w ustawieniach dodatku.")
	model = getActiveAiModel()
	if prov == "openai":
		return askOpenAiMultiTurn(chatHistory, apiKey=key, model=model)
	else:
		return askGeminiMultiTurn(chatHistory, apiKey=key, model=model)


def getDefaultSaveDirectory(category: str = "chat") -> str:
	"""Zwraca właściwy, intuicyjny folder zapisu dopasowany do preferencji użytkownika i języka."""
	baseCandidates = [
		r"G:\Mój dysk\Antigravity\Nvda dodatki",
		r"C:\Antigravity\Nvda dodatki",
	]
	baseDir = None
	for cand in baseCandidates:
		if os.path.isdir(cand):
			baseDir = cand
			break

	if not baseDir:
		docs = os.path.join(os.path.expanduser("~"), "Documents")
		if not os.path.isdir(docs):
			docs = os.path.expanduser("~")
		baseDir = os.path.join(docs, "Outlook Image Helper")

	lang = getEffectiveLanguage()
	prov = getActiveAiProvider()
	provName = "ChatGPT" if prov == "openai" else "Gemini"

	if category == "chat":
		folderName = f"Zapisane rozmowy z {provName}" if lang == "pl" else f"Saved chats with {provName}"
	elif category == "images":
		folderName = "Zapisane grafiki" if lang == "pl" else "Saved images"
	else:
		folderName = "Raporty" if lang == "pl" else "Reports"

	targetDir = os.path.join(baseDir, folderName)
	try:
		os.makedirs(targetDir, exist_ok=True)
	except Exception as ex:
		log.warning(f"outlookImageHelper: Nie udało się utworzyć folderu {targetDir}: {ex}")
		return baseDir
	return targetDir


def showGeminiKeyInstructions(parent=None):
	"""Wyświetla dostępne okno instrukcji uzyskania bezpłatnego klucza Google Gemini z linkiem."""
	url = "https://aistudio.google.com/app/apikey"
	msg = (
		"INSTRUKCJA: JAK UZYSKAĆ BEZPŁATNY KLUCZ API GOOGLE GEMINI\n"
		+ "=" * 60 + "\n\n"
		"Klucz API Google Gemini jest bezpłatny do użytku osobistego. Aby go uzyskać:\n\n"
		f"1. Otwórz w przeglądarce stronę Google AI Studio:\n   {url}\n\n"
		"2. Zaloguj się na swoje konto Google (Gmail).\n\n"
		"3. Kliknij przycisk 'Create API key' (Utwórz klucz API).\n\n"
		"4. Wybierz projekt Google Cloud lub zaakceptuj domyślny nowy projekt i zatwierdź.\n\n"
		"5. Wygenerowany klucz (długi ciąg liter i cyfr) skopiuj do schowka przyciskiem 'Copy'.\n\n"
		"6. Wróć do ustawień dodatku Outlook Image Helper w NVDA, wklej klucz w polu 'Klucz API Google Gemini' i kliknij przycisk 'Testuj klucz'.\n\n"
		"Możesz nacisnąć Tab i użyć przycisku 'Otwórz stronę w przeglądarce' lub naciśnij Enter/Esc, aby zamknąć tę instrukcję."
	)
	def _openUrl(evt):
		try:
			os.startfile(url)
		except Exception as e:
			ui.message(f"Błąd otwierania przeglądarki: {e}")

	showAccessibleMessageBox(
		msg,
		title="Instrukcja: Klucz API Google Gemini",
		parent=parent,
		buttonLabel="&Zamknij (Esc)",
		actionBtnLabel="&Otwórz stronę w przeglądarce",
		actionCallback=_openUrl,
	)


def showOpenAiKeyInstructions(parent=None):
	"""Wyświetla dostępne okno instrukcji uzyskania klucza OpenAI (ChatGPT) z linkiem."""
	url = "https://platform.openai.com/api-keys"
	msg = (
		"INSTRUKCJA: JAK UZYSKAĆ KLUCZ API OPENAI (CHATGPT)\n"
		+ "=" * 60 + "\n\n"
		"Aby korzystać z modeli ChatGPT (np. GPT-4o, GPT-4o-mini) do analizy obrazów i rozmowy:\n\n"
		f"1. Otwórz w przeglądarce platformę deweloperską OpenAI:\n   {url}\n\n"
		"2. Zaloguj się na swoje konto OpenAI lub utwórz nowe konto.\n\n"
		"3. Kliknij przycisk '+ Create new secret key' (Utwórz nowy klucz tajny).\n\n"
		"4. Nadaj kluczowi dowolną nazwę (np. 'NVDA Outlook') i zatwierdź.\n\n"
		"5. Skopiuj wyświetlony klucz (zaczyna się od 'sk-...') do schowka — pamiętaj, że klucz wyświetla się tylko raz!\n\n"
		"6. Wróć do ustawień dodatku w NVDA, wybierz dostawcę 'OpenAI / ChatGPT', wklej klucz w polu 'Klucz API OpenAI' i kliknij 'Testuj klucz'.\n\n"
		"Możesz nacisnąć Tab i użyć przycisku 'Otwórz stronę w przeglądarce' lub naciśnij Enter/Esc, aby zamknąć tę instrukcję."
	)
	def _openUrl(evt):
		try:
			os.startfile(url)
		except Exception as e:
			ui.message(f"Błąd otwierania przeglądarki: {e}")

	showAccessibleMessageBox(
		msg,
		title="Instrukcja: Klucz API OpenAI (ChatGPT)",
		parent=parent,
		buttonLabel="&Zamknij (Esc)",
		actionBtnLabel="&Otwórz stronę w przeglądarce",
		actionCallback=_openUrl,
	)



def getOrSaveSourceImageFile(item: GraphicItem | AttachmentItem) -> str | None:
	"""
	Zwraca ścieżkę do pliku grafiki źródłowej na dysku.
	Jeśli plik jeszcze nie istnieje na dysku, zapisuje dane z item.base64Image lub wyodrębnia z Worda do pliku tymczasowego PNG.
	Gwarantuje, że Microsoft Word / Outlook wklei obraz z lokalnego pliku URI file:///.
	"""
	if not item:
		return None
	try:
		if isinstance(item, AttachmentItem):
			if getattr(item, "isImage", False):
				return saveAttachmentToTempFile(item)
			return None
		b64 = getattr(item, "base64Image", "") or getattr(item, "base64Data", "")
		if b64:
			tempDir = tempfile.gettempdir()
			img_hash = hashlib.sha256(b64.encode("ascii", errors="ignore")).hexdigest()[:16]
			filePath = os.path.join(tempDir, f"nvda_oih_clip_{item.index}_{img_hash}.png")
			if not os.path.isfile(filePath):
				with open(filePath, "wb") as f:
					f.write(base64.b64decode(b64))
			return filePath
		return extractGraphicToPngFile(item)
	except Exception as ex:
		log.warning(f"outlookImageHelper: getOrSaveSourceImageFile błąd: {ex}")
		return None


def copyHtmlAndTextToClipboard(
	plainText: str,
	htmlBody: str | None = None,
	imageFilePath: str | None = None
) -> bool:
	"""
	Kopiuje jednocześnie sformatowany kod HTML, bitmapę oraz czysty tekst do schowka systemu Windows.
	Dzięki temu aplikacje bogate (Outlook, Word) wklejają sformatowaną treść z grafiką,
	programy graficzne wklejają bitmapę, a edytory tekstowe (Notatnik) wklejają czysty tekst.
	"""
	try:
		if wx.TheClipboard.Open():
			try:
				composite = wx.DataObjectComposite()
				cleanPlain = plainText.replace("\r\n", "\n").replace("\n", "\r\n")
				composite.Add(wx.TextDataObject(cleanPlain))
				if htmlBody:
					composite.Add(wx.HTMLDataObject(htmlBody))
				if imageFilePath and os.path.isfile(imageFilePath):
					try:
						bmp = wx.Bitmap(imageFilePath, wx.BITMAP_TYPE_PNG)
						if bmp.IsOk():
							composite.Add(wx.BitmapDataObject(bmp))
					except Exception as bex:
						log.debug(f"outlookImageHelper: composite.Add BitmapDataObject error: {bex}")
				wx.TheClipboard.SetData(composite)
				return True
			finally:
				wx.TheClipboard.Close()
	except Exception as e:
		log.error(f"outlookImageHelper: copyHtmlAndTextToClipboard error: {e}", exc_info=True)
	return False


def extractAnswerSections(answerText: str) -> list[tuple[str, str]]:
	"""
	Wyodrębnia poszczególne podrozdziały (sekcje) z odpowiedzi Markdown.
	Wyszukuje nagłówki Markdown (#, ##, ###, ####), nagłówki pogrubione oraz numerowane punkty główne.
	Zwraca listę krotek (tytuł_sekcji, pełna_treść_sekcji).
	"""
	if not answerText:
		return []

	lines = answerText.splitlines()
	sections = []
	curTitle = None
	curLines = []

	def _isHeadingLine(line: str) -> tuple[bool, str]:
		s = line.strip()
		if not s:
			return False, ""
		m = re.match(r"^#{1,5}\s+(.+)$", s)
		if m:
			return True, m.group(1).strip()
		m2 = re.match(r"^\*\*([^*]{3,80}?)\*\*:?$", s)
		if m2:
			return True, m2.group(1).strip()
		m3 = re.match(r"^(\d+\.\s+[A-ZĄĆĘŁŃÓŚŹŻ][^:\n]{3,60}:?)$", s)
		if m3 and not s.endswith("."):
			return True, m3.group(1).strip()
		return False, ""

	for line in lines:
		isHdr, title = _isHeadingLine(line)
		if isHdr:
			if curTitle is not None and curLines:
				sections.append((curTitle, "\n".join(curLines).strip()))
			curTitle = title
			curLines = [line]
		else:
			if curTitle is not None:
				curLines.append(line)

	if curTitle is not None and curLines:
		sections.append((curTitle, "\n".join(curLines).strip()))

	return sections


def markdownToHtml(mdText: str) -> str:
	"""Prosty, bezpieczny konwerter Markdown na HTML ze strukturą nagłówków, list i tabel dla czytników ekranu."""
	if not mdText:
		return ""
	import html as pyHtml
	escaped = pyHtml.escape(mdText)
	lines = escaped.splitlines()
	out = []
	inUl = False
	inTable = False

	for line in lines:
		sline = line.strip()
		if sline.startswith("### "):
			if inUl: out.append("</ul>"); inUl = False
			if inTable: out.append("</table>"); inTable = False
			out.append(f"<h4>{sline[4:]}</h4>")
		elif sline.startswith("## "):
			if inUl: out.append("</ul>"); inUl = False
			if inTable: out.append("</table>"); inTable = False
			out.append(f"<h4>{sline[3:]}</h4>")
		elif sline.startswith("# "):
			if inUl: out.append("</ul>"); inUl = False
			if inTable: out.append("</table>"); inTable = False
			out.append(f"<h4>{sline[2:]}</h4>")
		elif sline.startswith("#### ") or sline.startswith("##### "):
			if inUl: out.append("</ul>"); inUl = False
			if inTable: out.append("</table>"); inTable = False
			cleanHdr = re.sub(r"^#{4,5}\s*", "", sline)
			out.append(f"<h5>{cleanHdr}</h5>")
		elif sline.startswith("- ") or sline.startswith("* "):
			if inTable: out.append("</table>"); inTable = False
			if not inUl:
				out.append("<ul style='margin: 6px 0; padding-left: 20px;'>")
				inUl = True
			out.append(f"<li>{sline[2:]}</li>")
		elif sline.startswith("|") and sline.endswith("|"):
			if inUl: out.append("</ul>"); inUl = False
			if "---" in sline:
				continue
			cells = [c.strip() for c in sline[1:-1].split("|")]
			if not inTable:
				out.append("<table border='1' cellpadding='6' cellspacing='0' style='border-collapse: collapse; margin: 10px 0;'>")
				inTable = True
				out.append("<tr>" + "".join(f"<th style='background: #eef2f7;'>{c}</th>" for c in cells) + "</tr>")
			else:
				out.append("<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>")
		else:
			if inTable: out.append("</table>"); inTable = False
			if inUl: out.append("</ul>"); inUl = False
			if sline:
				formatted = sline
				formatted = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', formatted)
				formatted = re.sub(r'\*(.+?)\*', r'<em>\1</em>', formatted)
				formatted = re.sub(r'`(.+?)`', r'<code>\1</code>', formatted)
				out.append(f"<p style='margin: 6px 0;'>{formatted}</p>")

	if inTable: out.append("</table>")
	if inUl: out.append("</ul>")
	return "\n".join(out)


def buildHtmlSectionReport(
	item: GraphicItem | AttachmentItem,
	sectionTitle: str,
	sectionContent: str,
	includeImage: bool,
	imageFilePath: str | None = None
) -> str:
	"""Tworzy dokument HTML dla wybranej pojedynczej sekcji odpowiedzi AI z opcjonalnym obrazem."""
	imgTag = ""
	if includeImage:
		b64 = getattr(item, "base64Data", "") or getattr(item, "base64Image", "")
		if not b64:
			imgPath = imageFilePath or getOrSaveSourceImageFile(item)
			if imgPath and os.path.isfile(imgPath):
				try:
					with open(imgPath, "rb") as f:
						b64 = base64.b64encode(f.read()).decode("ascii")
				except Exception:
					pass
		if b64:
			imgTag = f"""
			<div style="margin: 15px 0; text-align: left;">
				<img src="data:image/png;base64,{b64}" alt="Źródłowy obraz" style="max-width: 100%; height: auto; border: 1px solid #ccc; border-radius: 4px;" />
			</div>
			"""
	formatted = markdownToHtml(sectionContent)
	titleEsc = html.escape(sectionTitle)
	return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>{titleEsc}</title></head>
<body style="font-family: Segoe UI, sans-serif; line-height: 1.6; color: #222; max-width: 850px; margin: 20px auto; padding: 15px;">
	<h3 style="color: #196f3d; border-bottom: 2px solid #27ae60; padding-bottom: 6px;">{titleEsc}</h3>
	{imgTag}
	<div style="margin-bottom: 14px; padding: 12px 16px; background: #ffffff; border-left: 4px solid #27ae60; border-radius: 4px;">
		<div style="font-size: 1.05em; line-height: 1.5;">{formatted}</div>
	</div>
</body>
</html>"""


def buildHtmlReport(
	item: GraphicItem | AttachmentItem,
	turns: list[dict],
	includeImage: bool,
	imageFilePath: str | None = None
) -> str:
	"""Tworzy kompletny, uniwersalny dokument HTML raportu z nagłówkami i wbudowanym obrazem (base64)."""
	imgTag = ""
	b64 = getattr(item, "base64Data", "") or getattr(item, "base64Image", "")
	if includeImage:
		if not b64:
			imgPath = imageFilePath or getOrSaveSourceImageFile(item)
			if imgPath and os.path.isfile(imgPath):
				try:
					with open(imgPath, "rb") as f:
						b64 = base64.b64encode(f.read()).decode("ascii")
				except Exception:
					pass
		if b64 and not getattr(item, "isPdf", False):
			imgTag = f"""
			<div style="margin: 15px 0; text-align: left;">
				<img src="data:image/png;base64,{b64}" alt="Źródłowy obraz" style="max-width: 100%; height: auto; border: 1px solid #ccc; border-radius: 4px;" />
			</div>
			"""

	prov = getActiveAiProvider()
	provName = "ChatGPT" if prov == "openai" else "Gemini"
	provFullName = "OpenAI (ChatGPT)" if prov == "openai" else "Google Gemini"
	continueUrl = "https://chatgpt.com" if prov == "openai" else "https://gemini.google.com/app"

	turnsHtml = []
	qIdx = 0
	aIdx = 0
	for turn in turns:
		isUser = (turn.get("role") == "user")
		if isUser:
			qIdx += 1
			role = f"Twoje pytanie {qIdx}" if getEffectiveLanguage() == "pl" else f"Your Question {qIdx}"
			hdrTag = "h2"
			hdrColor = "#1a5276"
		else:
			aIdx += 1
			role = f"Odpowiedź {provName} nr {aIdx}" if getEffectiveLanguage() == "pl" else f"{provName} Answer #{aIdx}"
			hdrTag = "h3"
			hdrColor = "#196f3d"
		bgColor = "#f0f4f8" if isUser else "#ffffff"
		borderColor = "#2980b9" if isUser else "#27ae60"
		timeStr = f" ({turn['timestamp']})" if turn.get("timestamp") else ""
		text = turn.get("text", "")
		formatted = markdownToHtml(text)
		turnsHtml.append(f"""
		<div style="margin-bottom: 14px; padding: 12px 16px; background: {bgColor}; border-left: 4px solid {borderColor}; border-radius: 4px;">
			<{hdrTag} style="margin: 0 0 8px 0; color: {hdrColor}; font-size: {'1.2em' if isUser else '1.1em'};">{role}{timeStr}</{hdrTag}>
			<div style="font-size: 1.05em; line-height: 1.5;">{formatted}</div>
		</div>
		""")

	isAtt = isinstance(item, AttachmentItem)
	rawTitle = f"Raport analizy załącznika: {item.fileName}" if isAtt else f"Raport analizy grafiki {item.index}"
	titleText = html.escape(rawTitle)
	detailMeta = f"<p><strong>Typ:</strong> {html.escape(item.typeName)} | <strong>Rozmiar:</strong> {html.escape(item.formattedSize)}</p>" if isAtt else f"<p><strong>Wymiary:</strong> {item.widthPx} × {item.heightPx} px</p>"
	ctx_val = getattr(item, "contextText", "") or ""
	ctxStr = f"<p><strong>Kontekst:</strong> {html.escape(ctx_val)}</p>" if ctx_val and not isAtt else ""

	link_val = getattr(item, "hyperlinkUrl", "") or ""
	if link_val:
		safe_url = urllib.parse.quote(link_val, safe=":/?#[]@!$&'()*+,;=-_.~%")
		parsed_u = urllib.parse.urlparse(safe_url)
		if parsed_u.scheme.lower() in ("http", "https", "mailto"):
			linkStr = f'<p><strong>Link:</strong> <a href="{html.escape(safe_url)}">{html.escape(link_val)}</a></p>'
		else:
			linkStr = f'<p><strong>Link:</strong> {html.escape(link_val)}</p>'
	else:
		linkStr = ""

	continueBox = f"""
	<div style="margin-top: 26px; padding: 16px; background: #e8f0fe; border-left: 5px solid #1a73e8; border-radius: 4px; font-family: Segoe UI, sans-serif;">
		<h3 style="margin-top: 0; color: #1a73e8; font-size: 1.1em;">Kontynuacja tej rozmowy w serwisie AI:</h3>
		<p style="margin-bottom: 8px;">Aby kontynuować tę rozmowę z inną osobą lub bezpośrednio w przeglądarce, otwórz poniższy link i wklej treść rozmowy:</p>
		<p style="margin: 0;"><a href="{continueUrl}" style="font-weight: bold; color: #1a73e8; font-size: 1.05em;" target="_blank">{continueUrl} ({provFullName})</a></p>
	</div>
	"""

	htmlReport = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>{titleText}</title></head>
<body style="font-family: Segoe UI, sans-serif; line-height: 1.6; color: #222; max-width: 850px; margin: 20px auto; padding: 15px;">
	<h1 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 6px; font-size: 1.5em;">{titleText}</h1>
	{detailMeta}
	{ctxStr}
	{linkStr}
	{imgTag}
	<div style="font-weight: bold; font-size: 1.15em; color: #444; margin: 20px 0 10px 0;">Zapis rozmowy i opisu:</div>
	{"".join(turnsHtml)}
	{continueBox}
</body>
</html>"""
	return htmlReport


def buildPlainTextReport(item: GraphicItem | AttachmentItem, turns: list[dict], includeImage: bool) -> str:
	"""Tworzy czytelny, tekstowy raport z analizy grafiki lub załącznika."""
	prov = getActiveAiProvider()
	provFullName = "Sztuczna inteligencja (OpenAI ChatGPT)" if prov == "openai" else "Sztuczna inteligencja (Google Gemini)"
	continueUrl = "https://chatgpt.com" if prov == "openai" else "https://gemini.google.com/app"

	isAtt = isinstance(item, AttachmentItem)
	titleStr = f"RAPORT ANALIZY ZAŁĄCZNIKA: {item.fileName}" if isAtt else f"RAPORT ANALIZY OBRAZU: Grafika {item.index}"
	lines = [
		"=" * 70,
		titleStr + (f" [{item.contextText}]" if getattr(item, "contextText", "") and not isAtt else ""),
		f"Typ: {item.typeName} ({item.formattedSize})" if isAtt else f"Wymiary: {item.widthPx} x {item.heightPx} px",
	]
	if getattr(item, "hyperlinkUrl", ""):
		lines.append(f"Powiązany link: {item.hyperlinkUrl}")
	b64 = getattr(item, "base64Data", "") or getattr(item, "base64Image", "")
	if includeImage and b64:
		lines.append("[Załączono źródłowy materiał (wersja HTML schowka zawiera obraz)]")
	lines.append("=" * 70 + "\r\n")

	for turn in turns:
		role = "Użytkownik" if turn.get("role") == "user" else provFullName
		timeStr = f" ({turn['timestamp']})" if turn.get("timestamp") else ""
		lines.append(f"[{role}{timeStr}]:")
		lines.append(turn.get("text", "").strip())
		lines.append("-" * 60)

	lines.append("\r\n" + "-" * 60)
	lines.append("KONTYNUACJA TEJ ROZMOWY:")
	lines.append("Aby kontynuować tę rozmowę z inną osobą lub w oknie przeglądarki, otwórz:")
	lines.append(f"{continueUrl} ({provFullName})")
	lines.append("i wklej treść powyższego raportu.")
	lines.append("-" * 60)

	return "\r\n".join(lines)



# ==============================================================================
# OKNO DIALOGOWE: WYBÓR I ZARZĄDZANIE PROMPTAMI AI
# ==============================================================================


class AiPromptSelectionDialog(wx.Dialog):
	"""
	Dostępne okno dialogowe umożliwiające wybór, edycję i zarządzanie szablonami promptów AI.
	Użytkownik może wybrać szablon z listy, edytować go przed wysłaniem, dodać nowy lub usunąć stary.
	"""

	def __init__(self, parent, item: GraphicItem | AttachmentItem):
		isAtt = isinstance(item, AttachmentItem)
		titleStr = f"Wybór instrukcji dla AI — Załącznik {item.index}: {item.fileName}" if isAtt else f"Wybór instrukcji dla AI — Grafika {item.index}"
		super().__init__(
			parent,
			title=titleStr,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.item = item
		self.presets = loadPromptPresets()
		self.selectedPrompt = ""

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		if isAtt:
			infoStr = f"Załącznik {item.index}: {item.fileName} ({item.formattedSize})"
		else:
			infoStr = f"Grafika {item.index}"
			if getattr(item, "contextText", ""):
				infoStr += f": [{item.contextText}]"
			infoStr += f" ({item.widthPx}×{item.heightPx} px)"
		lblInfo = wx.StaticText(self, label=f"Wybrany element: {infoStr}")
		lblInfo.SetFont(lblInfo.GetFont().Bold())
		mainSizer.Add(lblInfo, 0, wx.ALL, 8)

		lblCombo = wx.StaticText(self, label="&Wybierz szablon instrukcji z listy:")
		mainSizer.Add(lblCombo, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)

		titles = [p["title"] for p in self.presets]
		self.comboPresets = wx.Choice(self, choices=titles)
		if titles:
			self.comboPresets.SetSelection(0)
		mainSizer.Add(self.comboPresets, 0, wx.EXPAND | wx.ALL, 8)

		lblEdit = wx.StaticText(self, label="&Treść instrukcji dla AI (możesz zmodyfikować przed uruchomieniem):")
		mainSizer.Add(lblEdit, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)

		initPrompt = self.presets[0]["prompt"] if self.presets else ""
		self.promptEdit = wx.TextCtrl(
			self,
			value=initPrompt,
			style=wx.TE_MULTILINE | wx.HSCROLL,
			size=(-1, 140),
		)
		mainSizer.Add(self.promptEdit, 1, wx.EXPAND | wx.ALL, 8)

		manageSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnSaveAsNew = wx.Button(self, label="Zapisz jako nowy szablon...")
		manageSizer.Add(self.btnSaveAsNew, 0, wx.RIGHT, 5)

		self.btnDeletePreset = wx.Button(self, label="Usuń wybrany szablon")
		manageSizer.Add(self.btnDeletePreset, 0, wx.RIGHT, 5)
		mainSizer.Add(manageSizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnOk = wx.Button(self, wx.ID_OK, label="Generuj opis AI")
		self.btnOk.SetDefault()
		btnSizer.Add(self.btnOk, 0, wx.ALL, 5)

		self.btnCancel = wx.Button(self, wx.ID_CANCEL, label="Anuluj")
		btnSizer.Add(self.btnCancel, 0, wx.ALL, 5)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 8)

		self.SetSizer(mainSizer)
		self.Fit()
		self.SetSize((650, 460))
		self.CenterOnScreen()

		self.comboPresets.Bind(wx.EVT_CHOICE, self.onPresetChanged)
		self.btnSaveAsNew.Bind(wx.EVT_BUTTON, self.onSaveAsNew)
		self.btnDeletePreset.Bind(wx.EVT_BUTTON, self.onDeletePreset)
		self.btnOk.Bind(wx.EVT_BUTTON, self.onOk)
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

	def onPresetChanged(self, event):
		idx = self.comboPresets.GetSelection()
		if 0 <= idx < len(self.presets):
			self.promptEdit.SetValue(self.presets[idx]["prompt"])

	def onSaveAsNew(self, event):
		text = self.promptEdit.GetValue().strip()
		if not text:
			ui.message("Wprowadź treść instrukcji przed zapisaniem szablonu.")
			return
		with wx.TextEntryDialog(
			self,
			message="Podaj nazwę dla nowego szablonu prompta:",
			caption="Zapisz nowy szablon prompta",
			value="Mój własny prompt",
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				title = dlg.GetValue().strip()
				if title:
					newPreset = {
						"id": f"custom_{int(time.time())}",
						"title": title,
						"prompt": text,
					}
					self.presets.append(newPreset)
					savePromptPresets(self.presets)
					self.comboPresets.Append(title)
					self.comboPresets.SetSelection(len(self.presets) - 1)
					ui.message(f"Zapisano nowy szablon: {title}")

	def onDeletePreset(self, event):
		idx = self.comboPresets.GetSelection()
		if idx < 0 or idx >= len(self.presets):
			return
		preset = self.presets[idx]
		if len(self.presets) <= 1:
			ui.message("Nie można usunąć jedynego szablonu.")
			return
		if not showAccessibleConfirmBox(
			f"Czy na pewno chcesz usunąć szablon '{preset['title']}'?",
			title="Potwierdzenie usunięcia",
			parent=self,
			defaultYes=False
		):
			return
		del self.presets[idx]
		savePromptPresets(self.presets)
		self.comboPresets.Delete(idx)
		newIdx = min(idx, len(self.presets) - 1)
		self.comboPresets.SetSelection(newIdx)
		self.promptEdit.SetValue(self.presets[newIdx]["prompt"])
		ui.message("Usunięto szablon.")

	def onOk(self, event):
		self.selectedPrompt = self.promptEdit.GetValue().strip()
		self.EndModal(wx.ID_OK)

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		else:
			event.Skip()


# ==============================================================================
# KLASA WSKAŹNIKA POSTĘPU I FEEDBACKU DLA OPERACJI AI
# ==============================================================================


class AiProgressFeedback:
	"""Zapewnia ciągłe, subtelne powiadomienia dźwiękowo-głosowe podczas operacji AI i testów połączeń."""

	def __init__(self, startMessage: str = "Trwa łączenie ze sztuczną inteligencją Gemini, proszę czekać...", ongoingMessage: str = ""):
		self.running = True
		self.step = 0
		self.ongoingMessage = ongoingMessage
		if startMessage:
			ui.message(startMessage)
		try:
			tones.beep(880, 40)
		except Exception:
			pass
		self._scheduleNext()

	def _scheduleNext(self):
		if not self.running:
			return
		wx.CallLater(1200, self._onTick)

	def _onTick(self):
		if not self.running:
			return
		self.step += 1
		try:
			# Dyskretny, krótki ton (25 ms) informujący o aktywnym przetwarzaniu lub sprawdzaniu połączenia
			freq = 700 + (self.step % 3) * 120
			tones.beep(freq, 25)
		except Exception:
			pass
		if self.step == 4:
			msg = self.ongoingMessage or "Nadal trwa łączenie ze sztuczną inteligencją, proszę czekać..."
			ui.message(msg)
		elif self.step == 8:
			ui.message("Oczekiwanie na odpowiedź z serwera...")
		self._scheduleNext()

	def stop(self, finishMessage: str = ""):
		self.running = False
		if finishMessage:
			ui.message(finishMessage)


def playSuccessChime():
	try:
		playAlertSound("social_media", context="Sygnał ukończenia operacji")
	except Exception:
		pass


# ==============================================================================
# OKNO DIALOGOWE: ROZMOWA Z AI (JEDNO DOSTĘPNE OKNO CZATU GEMINI)
# ==============================================================================


class HeadingTarget:
	"""Reprezentuje cel nawigacji po nagłówkach wewnątrz okna rozmowy z AI."""

	def __init__(self, role: str, num: int, title: str, level: int, charOffset: int = 0):
		self.role = role  # "user" lub "model"
		self.num = num  # numer pytania lub odpowiedzi (1, 2, ...)
		self.title = title  # etykieta czytana przez syntezator
		self.level = level  # poziom nagłówka (2=pytanie, 3=odpowiedź, 4=wypunktowanie/sekcja, 5=podpunkt)
		self.charOffset = charOffset  # przesunięcie znakowe wewnątrz self.historyCtrl


class LinkTarget:
	"""Reprezentuje interaktywny link, adres e-mail, telefon lub pozycję GEO w oknie rozmowy z AI."""

	def __init__(
		self,
		kind: str,
		label: str,
		target: str,
		startOffset: int = 0,
		endOffset: int = 0,
		displaySnippet: str = ""
	):
		self.kind = kind  # "url", "email", "phone", "geo"
		self.label = label  # np. "Confluence", "ttst.zfss@ttst.pl", "+48 123 456 789"
		self.target = target  # np. "https://confluence.atlassian.com", "mailto:...", "tel:...", "https://maps.google.com/..."
		self.startOffset = startOffset
		self.endOffset = endOffset
		self.displaySnippet = displaySnippet or label


def isDarkColour(c: wx.Colour) -> bool:
	"""Zwraca True, jeśli kolor tła jest ciemny (dla doboru wysokiego kontrastu WCAG AA/AAA)."""
	if not c or not c.IsOk():
		return False
	lum = 0.2126 * c.Red() + 0.7152 * c.Green() + 0.0722 * c.Blue()
	return lum < 128


_RE_MARKDOWN_LINK = re.compile(
	r'\[(?P<label>[^\]\r\n]+)\]\((?P<url>[^\s\)\r\n]+)(?:\s+"[^"]*")?\)'
)
_RE_RAW_URL = re.compile(
	r'(?:https?://|www\.)[^\s<>"\'\(\)]+(?:\([^\s<>"\'\(\)]*\)[^\s<>"\'\(\)]*)*',
	re.IGNORECASE
)
_RE_EMAIL = re.compile(
	r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b',
	re.IGNORECASE
)
_RE_PHONE_PREFIXED = re.compile(
	r'(?i)\b(?:tel\.?|telefon|infolinia|kom\.?|mob\.?|fax)\s*:?\s*(?P<phone>\+?[0-9\s\(\)-]{7,20})\b'
)
_RE_PHONE_INTL = re.compile(
	r'(?<![\w@\d])\+\d{1,3}[\s-]?(?:\(?\d{2,4}\)?[\s-]?)?\d{3}[\s-]?\d{2,3}[\s-]?\d{2,3}(?![\w@\d])'
)
_RE_PHONE_PL = re.compile(
	r'(?<![\w@\d])(?:\+48[\s-]?)?(?:[1-9]\d{2}[\s-]?\d{3}[\s-]?\d{3}|[1-9]\d[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2})(?![\w@\d])'
)
_RE_GEO_DECIMAL = re.compile(
	r'(?<![\w\d])(?P<lat>[-+]?(?:[1-8]?\d(?:\.\d{3,})|90(?:\.0+)?))\s*°?\s*(?P<latDir>[NSns])?[,\s/]+\s*(?P<lon>[-+]?(?:(?:1[0-7]\d|\d{1,2})(?:\.\d{3,})|180(?:\.0+)?))\s*°?\s*(?P<lonDir>[EWew])?(?![\w\d])'
)
_RE_GEO_URI = re.compile(
	r'\bgeo:[-+]?[0-9.]+,[-+]?[0-9.]+(?:[^\s<>"\'\(\)]*)?',
	re.IGNORECASE
)


def extractInteractiveLinks(text: str) -> list[LinkTarget]:
	"""Wykrywa i indeksuje wszystkie interaktywne linki (URL, e-maile, telefony, GEO) w tekście historii rozmowy."""
	if not text:
		return []

	results: list[LinkTarget] = []
	occupiedRanges: list[tuple[int, int]] = []

	def isOverlapping(s: int, e: int) -> bool:
		return any(max(s, occS) < min(e, occE) for occS, occE in occupiedRanges)

	def addTarget(t: LinkTarget):
		results.append(t)
		occupiedRanges.append((t.startOffset, t.endOffset))

	# 1. Linki w formacie Markdown [etykieta](url)
	for m in _RE_MARKDOWN_LINK.finditer(text):
		s, e = m.span()
		label = m.group("label").strip()
		rawUrl = m.group("url").strip()
		urlLower = rawUrl.lower()
		if urlLower.startswith("mailto:"):
			kind = "email"
			target = rawUrl
		elif urlLower.startswith("tel:"):
			kind = "phone"
			target = rawUrl
		elif urlLower.startswith("geo:"):
			kind = "geo"
			target = rawUrl
		else:
			kind = "url"
			target = rawUrl if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", rawUrl) else f"https://{rawUrl}"
		addTarget(LinkTarget(kind=kind, label=label, target=target, startOffset=s, endOffset=e))

	# 2. Samodzielne adresy URL (http://, https://, www.)
	for m in _RE_RAW_URL.finditer(text):
		s, e = m.span()
		matchedText = m.group(0)
		while matchedText and matchedText[-1] in ".,;:!?*_)`'\"<>":
			matchedText = matchedText[:-1]
			e -= 1
		if not matchedText or isOverlapping(s, e):
			continue
		target = matchedText if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", matchedText) else f"https://{matchedText}"
		addTarget(LinkTarget(kind="url", label=matchedText, target=target, startOffset=s, endOffset=e))

	# 3. Adresy e-mail (np. ttst.zfss@ttst.pl)
	for m in _RE_EMAIL.finditer(text):
		s, e = m.span()
		email = m.group(0).rstrip(".,;")
		e = s + len(email)
		if isOverlapping(s, e):
			continue
		addTarget(LinkTarget(kind="email", label=email, target=f"mailto:{email}", startOffset=s, endOffset=e))

	# 4. URI geograficzne geo:
	for m in _RE_GEO_URI.finditer(text):
		s, e = m.span()
		geoUri = m.group(0).rstrip(".,;)")
		e = s + len(geoUri)
		if isOverlapping(s, e):
			continue
		coords = geoUri[4:]
		target = f"https://www.google.com/maps/search/?api=1&query={coords}"
		addTarget(LinkTarget(kind="geo", label=f"Pozycja GEO: {coords}", target=target, startOffset=s, endOffset=e))

	# 5. Współrzędne geograficzne w formacie dziesiętnym (np. 52.2297, 21.0122)
	for m in _RE_GEO_DECIMAL.finditer(text):
		s, e = m.span()
		if isOverlapping(s, e):
			continue
		try:
			latVal = float(m.group("lat"))
			lonVal = float(m.group("lon"))
			if m.group("latDir") and m.group("latDir").upper() == "S":
				latVal = -abs(latVal)
			if m.group("lonDir") and m.group("lonDir").upper() == "W":
				lonVal = -abs(lonVal)
			if -90 <= latVal <= 90 and -180 <= lonVal <= 180:
				label = f"{latVal:.5f}, {lonVal:.5f}"
				target = f"https://www.google.com/maps/search/?api=1&query={latVal:.6f},{lonVal:.6f}"
				addTarget(LinkTarget(kind="geo", label=label, target=target, startOffset=s, endOffset=e))
		except Exception:
			pass

	# 6. Numery telefonów z prefiksem (np. tel. +48 123 456 789)
	for m in _RE_PHONE_PREFIXED.finditer(text):
		s, e = m.span("phone")
		phoneStr = m.group("phone").strip().rstrip(".,;)")
		e = s + len(phoneStr)
		if isOverlapping(s, e):
			continue
		digits = re.sub(r"[^\d+]", "", phoneStr)
		if len(re.sub(r"\D", "", digits)) >= 7:
			addTarget(LinkTarget(kind="phone", label=phoneStr, target=f"tel:{digits}", startOffset=s, endOffset=e))

	# 7. Numery telefonów międzynarodowe i polskie
	for regex in (_RE_PHONE_INTL, _RE_PHONE_PL):
		for m in regex.finditer(text):
			s, e = m.span()
			phoneStr = m.group(0).strip().rstrip(".,;)")
			e = s + len(phoneStr)
			if isOverlapping(s, e):
				continue
			afterSnippet = text[e:e+10].strip().lower()
			if re.match(r"^(zł|pln|eur|usd|\$|€|px|%|kg|km|m\b|godz|dni|r\b)", afterSnippet):
				continue
			digits = re.sub(r"[^\d+]", "", phoneStr)
			if 7 <= len(re.sub(r"\D", "", digits)) <= 15:
				addTarget(LinkTarget(kind="phone", label=phoneStr, target=f"tel:{digits}", startOffset=s, endOffset=e))

	results.sort(key=lambda t: t.startOffset)
	return results


class AiConversationDialog(wx.Dialog):
	"""
	Zunifikowane, w pełni dostępne okno rozmowy z AI (Google Gemini).
	Działa analogicznie do czatu Gemini w przeglądarce:
	- Każde pytanie i odpowiedź posiada własny nagłówek (nawigacja klawiszami H oraz Shift+H),
	- Pod każdą z odpowiedzi AI znajdują się dedykowane przyciski: kopiowania tej odpowiedzi oraz zapisu do pliku,
	- Klawisz B / Shift+B pozwala błyskawicznie przeskakiwać po przyciskach pod odpowiedziami,
	- Klawisz E przenosi bezpośrednio do pola wprowadzania kolejnego pytania,
	- Klawisz C kopiuje aktualną odpowiedź, klawisz S zapisuje ją do pliku,
	- Pole wprowadzania kolejnego pytania (Enter lub Ctrl+Enter wysyła, Shift+Enter nowa linia),
	- Dźwiękowy i głosowy wskaźnik postępu (AiProgressFeedback) podczas generowania odpowiedzi,
	- Przyciski zbiorcze na dole: kopiuj całą rozmowę, eksportuj do pliku, dołącz grafikę źródłową, zamknij.
	"""

	def __init__(self, parent, item: GraphicItem | AttachmentItem):
		self.item = item
		self.headingTargets: list[HeadingTarget] = []
		self.turnRanges: list[tuple[str, int, int, int]] = []  # (role, num, startOffset, endOffset)
		self.linkTargets: list[LinkTarget] = []

		isAtt = isinstance(item, AttachmentItem)
		if isAtt:
			target = f"{item.fileName} ({item.formattedSize})"
		else:
			cName = getattr(item, "customName", "")
			if cName:
				target = f"[{cName}]"
			else:
				svc = getattr(item, "linkServiceName", "")
				target = f"{item.index}" + (f" ({svc})" if svc else "")
			target += f" ({item.widthPx}×{item.heightPx} px)"
		titleHeader = _T("chat_title", target=target)

		super().__init__(
			parent,
			title=titleHeader,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
			size=(860, 700),
		)
		self.SetMinSize((640, 480))

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# 1. Ciągły, zunifikowany bufor historii pytań i odpowiedzi (pełna dostępność Ctrl+Home i strzałek)
		self.lblHistory = wx.StaticText(
			self,
			label=_T("chat_history_lbl"),
		)
		mainSizer.Add(self.lblHistory, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)

		self.historyCtrl = wx.TextCtrl(
			self,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.TE_AUTO_URL,
			size=(-1, 380),
		)
		self.historyCtrl.SetName(_T("chat_history_lbl"))
		mainSizer.Add(self.historyCtrl, 1, wx.EXPAND | wx.ALL, 8)

		# 2. Pole nowego pytania
		self.lblQuestion = wx.StaticText(
			self,
			label=_T("chat_question_lbl"),
		)
		mainSizer.Add(self.lblQuestion, 0, wx.LEFT | wx.RIGHT | wx.TOP, 4)

		self.questionCtrl = wx.TextCtrl(
			self,
			style=wx.TE_MULTILINE,
			size=(-1, 75),
		)
		self.questionCtrl.SetName(_T("chat_question_lbl"))
		mainSizer.Add(self.questionCtrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

		# 3. Pasek przycisków głównych czynności zbiorczych
		btnSizer = wx.WrapSizer(wx.HORIZONTAL)

		self.btnSend = wx.Button(self, label=_T("chat_btn_send"))
		self.btnSend.SetDefault()
		btnSizer.Add(self.btnSend, 0, wx.ALL, 4)

		self.btnCopyLast = wx.Button(self, label=_T("chat_btn_copy_last"))
		btnSizer.Add(self.btnCopyLast, 0, wx.ALL, 4)

		self.btnCopyAll = wx.Button(self, label=_T("chat_btn_copy_all"))
		btnSizer.Add(self.btnCopyAll, 0, wx.ALL, 4)

		self.btnExportAll = wx.Button(self, label="Zapisz / eksportuj rozmowę...")
		btnSizer.Add(self.btnExportAll, 0, wx.ALL, 4)

		self.btnAnswersMenu = wx.Button(self, label="Czynności dla wybranej odpowiedzi...")
		btnSizer.Add(self.btnAnswersMenu, 0, wx.ALL, 4)

		self.chkIncludeImage = wx.CheckBox(
			self,
			label=_T("chat_chk_include_img"),
		)
		self.chkIncludeImage.SetValue(True)
		btnSizer.Add(self.chkIncludeImage, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 4)

		self.btnHelp = wx.Button(self, label=_T("chat_btn_help"))
		btnSizer.Add(self.btnHelp, 0, wx.ALL, 4)

		self.btnClose = wx.Button(self, wx.ID_CANCEL, label=_T("chat_btn_close"))
		btnSizer.Add(self.btnClose, 0, wx.ALL, 4)

		mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

		self.SetSizer(mainSizer)
		self.CenterOnScreen()

		# Zdarzenia kontrolek
		self.btnSend.Bind(wx.EVT_BUTTON, lambda evt: self.onSend())
		self.btnCopyLast.Bind(wx.EVT_BUTTON, self.onCopyLast)
		self.btnCopyAll.Bind(wx.EVT_BUTTON, self.onCopyAll)
		self.btnExportAll.Bind(wx.EVT_BUTTON, self.onExportMenu)
		self.btnAnswersMenu.Bind(wx.EVT_BUTTON, self.onAnswersMenu)
		self.btnHelp.Bind(wx.EVT_BUTTON, lambda evt: showContextHelp("chat", parent=self))
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		self.historyCtrl.Bind(wx.EVT_TEXT_URL, self.onTextUrl)
		self.historyCtrl.Bind(wx.EVT_LEFT_DCLICK, self.onHistoryDClick)
		self.historyCtrl.Bind(wx.EVT_CONTEXT_MENU, self.onHistoryContextMenu)

		self.updateHistoryDisplay(scrollToEnd=True)
		if not self.item.chatHistory:
			wx.CallAfter(self.questionCtrl.SetFocus)
		else:
			wx.CallAfter(self.historyCtrl.SetFocus)

	def getAllButtons(self) -> list[wx.Button]:
		"""Zwraca listę wszystkich aktywnych przycisków w oknie w logicznej kolejności fokusu."""
		btns = [
			self.btnSend, self.btnCopyLast, self.btnCopyAll,
			self.btnExportAll, self.btnAnswersMenu, self.btnHelp, self.btnClose
		]
		return [b for b in btns if b and b.IsEnabled()]

	def updateHistoryDisplay(self, scrollToEnd: bool = True):
		"""Aktualizuje zawartość jednolitej historii rozmowy, generując hierarchię nagłówków Markdown (H2, H3, H4, H5) oraz linków."""
		self.headingTargets = []
		self.turnRanges = []
		self.linkTargets = []

		if not self.item.chatHistory:
			placeholder = (
				"ROZMOWA Z AI JEST PUSTA\n"
				+ "=" * 60 + "\n\n"
				"Wpisz pytanie w polu edycji poniżej i naciśnij Enter lub Ctrl+Enter, "
				"aby zapytać sztuczną inteligencję o tę grafikę lub załącznik."
			)
			self.historyCtrl.SetValue(placeholder)
			self.btnCopyLast.Disable()
			self.btnCopyAll.Disable()
			self.btnExportAll.Disable()
			self.btnAnswersMenu.Disable()
			return

		self.btnCopyLast.Enable()
		self.btnCopyAll.Enable()
		self.btnExportAll.Enable()
		self.btnAnswersMenu.Enable()

		contentParts = []
		currentOffset = 0
		qCount = 0
		aCount = 0

		# Sprawdź, czy pierwsza tura to automatyczny opis wygenerowany przez system (placeholder w nawiasach)
		firstTurn = self.item.chatHistory[0] if self.item.chatHistory else {}
		firstIsSynthetic = (
			firstTurn.get("role") == "user" and
			firstTurn.get("text", "").strip().startswith("(") and
			firstTurn.get("text", "").strip().endswith(")")
		)

		for turnIdx, turn in enumerate(self.item.chatHistory):
			role = turn.get("role")
			timestamp = turn.get("timestamp", "")
			timeStr = f" ({timestamp})" if timestamp else ""
			rawText = turn.get("text", "").strip()

			isIntroTurn = (firstIsSynthetic and turnIdx <= 1)

			if role == "user":
				if isIntroTurn:
					hdrTitle = f"Opis początkowy AI{timeStr}"
					displayTitle = f"## OPIS POCZĄTKOWY AI{timeStr.upper()}"
				else:
					qCount += 1
					hdrTitle = f"Twoje pytanie {qCount}{timeStr}"
					displayTitle = f"## TWOJE PYTANIE {qCount}{timeStr.upper()}"

				blockHeader = f"{displayTitle}\n"
				blockBody = f"{rawText}\n\n"
				fullBlock = blockHeader + blockBody

				self.headingTargets.append(
					HeadingTarget(role="user", num=qCount, title=hdrTitle, level=2, charOffset=currentOffset)
				)
				self.turnRanges.append(("user", qCount, currentOffset, currentOffset + len(fullBlock)))
				contentParts.append(fullBlock)
				currentOffset += len(fullBlock)

			else:  # role == "model"
				aCount += 1
				if isIntroTurn and aCount == 1:
					hdrTitle = f"Odpowiedź — Opis początkowy{timeStr}"
					displayTitle = f"### ODPOWIEDŹ — OPIS POCZĄTKOWY{timeStr.upper()}"
				else:
					hdrTitle = f"Odpowiedź Gemini #{aCount}{timeStr}"
					displayTitle = f"### ODPOWIEDŹ GEMINI #{aCount}{timeStr.upper()}"

				blockHeader = f"{displayTitle}\n"
				self.headingTargets.append(
					HeadingTarget(role="model", num=aCount, title=hdrTitle, level=3, charOffset=currentOffset)
				)

				# Analiza śródnagłówków i wypunktowań wewnątrz odpowiedzi (H4 / H5)
				lines = rawText.split("\n")
				lineOffset = currentOffset + len(blockHeader)
				for line in lines:
					sline = line.strip()
					if sline.startswith("### "):
						st = sline[4:].strip().replace("**", "").replace("__", "")
						if st:
							self.headingTargets.append(HeadingTarget(role="model", num=aCount, title=st, level=4, charOffset=lineOffset))
					elif sline.startswith("## "):
						st = sline[3:].strip().replace("**", "").replace("__", "")
						if st:
							self.headingTargets.append(HeadingTarget(role="model", num=aCount, title=st, level=4, charOffset=lineOffset))
					elif sline.startswith("# "):
						st = sline[2:].strip().replace("**", "").replace("__", "")
						if st:
							self.headingTargets.append(HeadingTarget(role="model", num=aCount, title=st, level=4, charOffset=lineOffset))
					elif sline.startswith("#### ") or sline.startswith("##### "):
						st = re.sub(r"^#{4,5}\s*", "", sline).strip().replace("**", "").replace("__", "")
						if st:
							self.headingTargets.append(HeadingTarget(role="model", num=aCount, title=st, level=5, charOffset=lineOffset))
					elif re.match(r"^\s*\d+[.)]\s+", line):
						mNum = re.match(r"^(\s*)(\d+[.)])\s+(.*)$", line)
						if mNum:
							indent = len(mNum.group(1))
							numPrefix = mNum.group(2)
							rest = mNum.group(3).strip()
							lvl = 5 if indent >= 2 else 4
							mBold = re.match(r"^\*\*([^*]+)\*\*[:\s]*(.*)$", rest)
							if mBold:
								cleanTitle = f"{numPrefix} {mBold.group(1).strip()}"
							elif ":" in rest and len(rest.split(":", 1)[0]) <= 50:
								cleanTitle = f"{numPrefix} {rest.split(':', 1)[0].strip()}"
							else:
								cleanTitle = f"{numPrefix} {_cleanSnippet(rest, maxLen=42)}"
							cleanTitle = cleanTitle.replace("**", "").replace("__", "").rstrip(":")
							self.headingTargets.append(HeadingTarget(role="model", num=aCount, title=cleanTitle, level=lvl, charOffset=lineOffset))
					elif re.match(r"^\s*[-*•]\s+", line):
						mBul = re.match(r"^(\s*)[-*•]\s+(.*)$", line)
						if mBul:
							indent = len(mBul.group(1))
							rest = mBul.group(2).strip()
							lvl = 5 if indent >= 2 else 4
							mBold = re.match(r"^\*\*([^*]+)\*\*[:\s]*(.*)$", rest)
							if mBold:
								cleanTitle = mBold.group(1).strip()
							elif ":" in rest and len(rest.split(":", 1)[0]) <= 50:
								cleanTitle = rest.split(":", 1)[0].strip()
							else:
								cleanTitle = _cleanSnippet(rest, maxLen=42)
							cleanTitle = cleanTitle.replace("**", "").replace("__", "").rstrip(":")
							if cleanTitle:
								self.headingTargets.append(HeadingTarget(role="model", num=aCount, title=cleanTitle, level=lvl, charOffset=lineOffset))
					elif re.match(r"^\s*\*\*([^*]{3,60})\*\*[:\s]*$", line):
						mHdr = re.match(r"^\s*\*\*([^*]{3,60})\*\*[:\s]*$", line)
						if mHdr:
							cleanTitle = mHdr.group(1).strip().rstrip(":")
							self.headingTargets.append(HeadingTarget(role="model", num=aCount, title=cleanTitle, level=4, charOffset=lineOffset))

					lineOffset += len(line) + 1  # +1 dla nowej linii

				divider = "=" * 60 + "\n\n"
				fullBlock = blockHeader + f"{rawText}\n\n" + divider
				self.turnRanges.append(("model", aCount, currentOffset, currentOffset + len(fullBlock)))
				contentParts.append(fullBlock)
				currentOffset += len(fullBlock)

		fullText = "".join(contentParts).rstrip()
		self.historyCtrl.SetValue(fullText)
		self.linkTargets = extractInteractiveLinks(fullText)
		self.applyLinkStyles()

		if scrollToEnd and self.headingTargets:
			# Ustaw pozycję na ostatnim nagłówku
			lastTarget = self.headingTargets[-1]
			self.historyCtrl.SetInsertionPoint(lastTarget.charOffset)
			self.historyCtrl.ShowPosition(lastTarget.charOffset)
			self.historyCtrl.SetFocus()

	def applyLinkStyles(self):
		"""Nakłada wyróżnienie wizualne (podkreślenie i wysoki kontrast WCAG AA/AAA) na wszystkie wykryte linki."""
		if not self.linkTargets:
			return
		isDark = isDarkColour(self.historyCtrl.GetBackgroundColour())
		linkColor = wx.Colour(100, 180, 255) if isDark else wx.Colour(0, 90, 180)
		font = self.historyCtrl.GetFont()
		font.SetUnderlined(True)
		attr = wx.TextAttr()
		attr.SetTextColour(linkColor)
		attr.SetFont(font)
		for t in self.linkTargets:
			try:
				self.historyCtrl.SetStyle(t.startOffset, t.endOffset, attr)
			except Exception:
				pass

	def getCurrentHeadingIndex(self, forward: bool) -> int:
		"""Wyznacza indeks najbliższego nagłówka w zależności od aktualnej pozycji kursora w historii."""
		curOffset = self.historyCtrl.GetInsertionPoint()
		matchedIdx = -1
		for idx, t in enumerate(self.headingTargets):
			if t.charOffset <= curOffset:
				matchedIdx = idx
			else:
				break
		if matchedIdx != -1:
			return matchedIdx
		return -1 if forward else len(self.headingTargets)

	def navigateHeading(self, forward: bool, targetLevel: int = None):
		"""Nawiguje płynnie do kolejnego lub poprzedniego nagłówka (H / Shift+H lub cyfry 1-6) z automatycznym zawijaniem."""
		if not self.headingTargets:
			ui.message("Brak nagłówków w rozmowie.")
			return

		curOffset = self.historyCtrl.GetInsertionPoint()

		if targetLevel is not None:
			matching = [
				idx for idx, t in enumerate(self.headingTargets)
				if t.level == targetLevel
			]
			if not matching:
				ui.message(f"Brak nagłówków poziomu {targetLevel} w tej rozmowie.")
				return

			chosenIdx = None
			if forward:
				for idx in matching:
					if self.headingTargets[idx].charOffset > curOffset:
						chosenIdx = idx
						break
				if chosenIdx is None:
					ui.message(f"Koniec czatu (brak dalszych nagłówków poziomu {targetLevel})")
					return
			else:
				for idx in reversed(matching):
					if self.headingTargets[idx].charOffset < curOffset:
						chosenIdx = idx
						break
				if chosenIdx is None:
					ui.message(f"Początek czatu (brak wcześniejszych nagłówków poziomu {targetLevel})")
					return

			nextIdx = chosenIdx
		else:
			chosenIdx = None
			if forward:
				for idx, t in enumerate(self.headingTargets):
					if t.charOffset > curOffset:
						chosenIdx = idx
						break
				if chosenIdx is None:
					ui.message("Koniec czatu (ostatni nagłówek)")
					return
			else:
				for idx in reversed(range(len(self.headingTargets))):
					if self.headingTargets[idx].charOffset < curOffset:
						chosenIdx = idx
						break
				if chosenIdx is None:
					ui.message("Początek czatu (pierwszy nagłówek)")
					return
			nextIdx = chosenIdx

		target = self.headingTargets[nextIdx]
		self.historyCtrl.SetFocus()
		self.historyCtrl.SetInsertionPoint(target.charOffset)
		self.historyCtrl.ShowPosition(target.charOffset)
		ui.message(f"Nagłówek poziom {target.level}: {target.title}")

	def navigateLink(self, forward: bool):
		"""Nawiguje płynnie do kolejnego lub poprzedniego interaktywnego linku (K / Shift+K)."""
		if not self.linkTargets:
			ui.message("Brak linków w rozmowie.")
			return

		curOffset = self.historyCtrl.GetInsertionPoint()
		chosenIdx = None

		if forward:
			for idx, t in enumerate(self.linkTargets):
				if t.startOffset > curOffset:
					chosenIdx = idx
					break
			if chosenIdx is None:
				chosenIdx = 0
				ui.message("Zawijanie do pierwszego linku.")
		else:
			for idx in reversed(range(len(self.linkTargets))):
				if self.linkTargets[idx].startOffset < curOffset:
					chosenIdx = idx
					break
			if chosenIdx is None:
				chosenIdx = len(self.linkTargets) - 1
				ui.message("Zawijanie do ostatniego linku.")

		target = self.linkTargets[chosenIdx]
		self.historyCtrl.SetFocus()
		self.historyCtrl.SetInsertionPoint(target.startOffset)
		self.historyCtrl.SetSelection(target.startOffset, target.endOffset)
		self.historyCtrl.ShowPosition(target.startOffset)

		if target.kind == "email":
			playAlertSound("email", context="Nawigacja: e-mail")
		elif target.kind == "phone":
			playAlertSound("phone", context="Nawigacja: telefon")
		else:
			playAlertSound("link", context="Nawigacja: link")

		kindLabels = {
			"url": "Link",
			"email": "Link e-mail",
			"phone": "Link telefon",
			"geo": "Link lokalizacja",
		}
		kLbl = kindLabels.get(target.kind, "Link")
		total = len(self.linkTargets)
		posStr = f"({chosenIdx + 1} z {total})"

		if target.kind == "url" and target.label and target.label != target.target:
			msg = f"{kLbl} {posStr}: {target.label}, {target.target}"
		else:
			msg = f"{kLbl} {posStr}: {target.label}"

		ui.message(f"{msg}. Naciśnij Enter, aby otworzyć.")

	def findLinkAtOffset(self, offset: int) -> LinkTarget | None:
		"""Zwraca interaktywny link pod podanym przesunięciem lub zaznaczeniem."""
		selFrom, selTo = self.historyCtrl.GetSelection()
		if selFrom != selTo:
			for t in self.linkTargets:
				if max(selFrom, t.startOffset) < min(selTo, t.endOffset):
					return t

		for t in self.linkTargets:
			if t.startOffset <= offset <= t.endOffset:
				return t
			if abs(offset - t.startOffset) <= 1 or abs(offset - t.endOffset) <= 1:
				return t

		return None

	def activateLink(self, link: LinkTarget):
		"""Uruchamia domyślną aplikację powiązaną z wybranym linkiem (przeglądarka, poczta, telefon, mapa)."""
		if not link:
			return

		if link.kind == "email":
			playAlertSound("email", context="Otwieranie poczty")
			ui.message(f"Otwieram program pocztowy: {link.label}")
		elif link.kind == "phone":
			playAlertSound("phone", context="Wybieranie telefonu")
			ui.message(f"Wybieram numer telefonu: {link.label}")
		elif link.kind == "geo":
			playAlertSound("link", context="Otwieranie mapy")
			ui.message(f"Otwieram mapę: {link.label}")
		else:
			playAlertSound("link", context="Otwieranie linku")
			ui.message(f"Otwieram: {link.label}")

		openInteractiveTarget(link.target, kind=link.kind)

	def copyLinkToClipboard(self, link: LinkTarget):
		"""Kopiuje docelowy adres linku do schowka."""
		if not link:
			return
		cleanTarget = link.target
		if cleanTarget.startswith("mailto:"):
			cleanTarget = cleanTarget[7:]
		elif cleanTarget.startswith("tel:"):
			cleanTarget = cleanTarget[4:]
		if hasattr(api, "copyToClip") and api.copyToClip(cleanTarget):
			ui.message(f"Skopiowano do schowka: {cleanTarget}")
		elif copyHtmlAndTextToClipboard(cleanTarget):
			ui.message(f"Skopiowano do schowka: {cleanTarget}")

	def showLinksMenu(self):
		"""Wyświetla podręczne menu wszystkich wykrytych linków i kontaktów w rozmowie (klawisz L)."""
		if not self.linkTargets:
			ui.message("Brak linków w tej rozmowie.")
			return

		menu = wx.Menu()
		for idx, lt in enumerate(self.linkTargets, 1):
			idOpen = wx.NewIdRef()
			kindTag = {
				"url": "Link",
				"email": "E-mail",
				"phone": "Tel",
				"geo": "Mapa",
			}.get(lt.kind, "Link")
			display = f"{idx}. [{kindTag}] {lt.label}"
			if len(display) > 60:
				display = display[:57] + "..."
			menu.Append(idOpen, display)
			self.Bind(wx.EVT_MENU, lambda evt, l=lt: self.activateLink(l), id=idOpen)

		self.PopupMenu(menu)
		menu.Destroy()

	def onTextUrl(self, event):
		mouseEvt = event.GetMouseEvent()
		if mouseEvt and mouseEvt.LeftUp():
			pos = event.GetURLStart()
			link = self.findLinkAtOffset(pos)
			if link:
				self.activateLink(link)
			else:
				rawUrl = self.historyCtrl.GetRange(event.GetURLStart(), event.GetURLEnd()).strip()
				if rawUrl:
					openInteractiveTarget(rawUrl, kind="url")
		event.Skip()

	def onHistoryDClick(self, event):
		curOffset = self.historyCtrl.GetInsertionPoint()
		link = self.findLinkAtOffset(curOffset)
		if link:
			self.activateLink(link)
			return
		event.Skip()

	def onHistoryContextMenu(self, event):
		curOffset = self.historyCtrl.GetInsertionPoint()
		link = self.findLinkAtOffset(curOffset)
		menu = wx.Menu()

		if link:
			idAct = wx.NewIdRef()
			actLabels = {
				"url": "Otwórz w przeglądarce",
				"email": "Napisz wiadomość e-mail",
				"phone": "Wybierz numer telefonu",
				"geo": "Pokaż lokalizację na mapie",
			}
			menu.Append(idAct, f"{actLabels.get(link.kind, 'Otwórz')}: {link.label}")
			self.Bind(wx.EVT_MENU, lambda evt, l=link: self.activateLink(l), id=idAct)

			idCopy = wx.NewIdRef()
			menu.Append(idCopy, "Kopiuj adres linku do schowka")
			self.Bind(wx.EVT_MENU, lambda evt, l=link: self.copyLinkToClipboard(l), id=idCopy)
			menu.AppendSeparator()

		if self.linkTargets:
			sub = wx.Menu()
			for idx, lt in enumerate(self.linkTargets, 1):
				idSub = wx.NewIdRef()
				sub.Append(idSub, f"{idx}. [{lt.kind.upper()}] {lt.label}")
				self.Bind(wx.EVT_MENU, lambda evt, l=lt: self.activateLink(l), id=idSub)
			menu.AppendSubMenu(sub, f"Wszystkie linki w rozmowie ({len(self.linkTargets)})...	L")
			menu.AppendSeparator()

		idCopySel = wx.NewIdRef()
		menu.Append(idCopySel, "Kopiuj zaznaczony tekst\tCtrl+C")
		self.Bind(wx.EVT_MENU, lambda evt: self.historyCtrl.Copy(), id=idCopySel)

		self.PopupMenu(menu)
		menu.Destroy()


	def navigateButton(self, forward: bool):
		"""Nawiguje do następnego lub poprzedniego przycisku (B / Shift+B)."""
		btns = self.getAllButtons()
		if not btns:
			return

		focused = self.FindFocus()
		curIdx = -1
		for idx, b in enumerate(btns):
			if b == focused:
				curIdx = idx
				break

		if curIdx == -1:
			nextIdx = 0 if forward else len(btns) - 1
		else:
			nextIdx = curIdx + 1 if forward else curIdx - 1
			if nextIdx >= len(btns):
				nextIdx = 0
			elif nextIdx < 0:
				nextIdx = len(btns) - 1

		targetBtn = btns[nextIdx]
		targetBtn.SetFocus()
		ui.message(f"{targetBtn.GetLabel()}, przycisk")

	def focusQuestionEdit(self):
		"""Przenosi fokus bezpośrednio do pola edycji pytania (klawisz E)."""
		self.questionCtrl.SetFocus()
		ui.message("Pole edycji: Zadaj pytanie")

	def getTurnAtCaret(self) -> tuple[int | None, dict | None]:
		"""Zwraca wypowiedź modelu AI odpowiadającą aktualnej pozycji kursora w historii."""
		curOffset = self.historyCtrl.GetInsertionPoint()
		modelTurns = [t for t in self.item.chatHistory if t.get("role") == "model"]
		for role, num, start, end in self.turnRanges:
			if start <= curOffset <= end and role == "model":
				if 0 <= num - 1 < len(modelTurns):
					return num, modelTurns[num - 1]
		if modelTurns:
			return len(modelTurns), modelTurns[-1]
		return None, None

	def onCopyKey(self):
		"""Klawisz C: kopiuje odpowiedź pod kursorem lub ostatnią odpowiedź."""
		num, turn = self.getTurnAtCaret()
		if turn and num:
			self.onCopySpecificAnswer(num, turn)
		else:
			self.onCopyLast(None)

	def onSaveKey(self):
		"""Klawisz S: zapisuje odpowiedź pod kursorem lub całą rozmowę."""
		num, turn = self.getTurnAtCaret()
		if turn and num:
			self.onSaveSpecificAnswer(num, turn)
		else:
			self.onExport(None)

	def onAnswersMenu(self, event=None):
		"""Wyświetla menu czynności umożliwiające skopiowanie lub zapisanie całej odpowiedzi lub poszczególnych jej sekcji oraz linków."""
		modelTurns = [(idx + 1, t) for idx, t in enumerate(self.item.chatHistory) if t.get("role") == "model"]
		if not modelTurns:
			ui.message("Brak odpowiedzi AI w tej rozmowie.")
			return

		menu = wx.Menu()

		# Jeśli w rozmowie są wykryte linki, dodaj na samej górze sekcję szybkich linków
		if self.linkTargets:
			linksSub = wx.Menu()
			for idx, lt in enumerate(self.linkTargets, 1):
				idL = wx.NewIdRef()
				kindTag = {
					"url": "Link",
					"email": "E-mail",
					"phone": "Tel",
					"geo": "Mapa",
				}.get(lt.kind, "Link")
				display = f"{idx}. [{kindTag}] {lt.label}"
				if len(display) > 60:
					display = display[:57] + "..."
				linksSub.Append(idL, display)
				self.Bind(wx.EVT_MENU, lambda evt, l=lt: self.activateLink(l), id=idL)
			menu.AppendSubMenu(linksSub, f"Wszystkie linki i kontakty ({len(self.linkTargets)})...	L")
			menu.AppendSeparator()

		if len(modelTurns) == 1:
			aNum, turn = modelTurns[0]
			text = turn.get("text", "")
			sections = extractAnswerSections(text)

			idCopyAll = wx.NewIdRef()
			menu.Append(idCopyAll, "Kopiuj całą treść odpowiedzi\tC")
			self.Bind(wx.EVT_MENU, lambda evt, n=aNum, t=turn: self.onCopySpecificAnswer(n, t), id=idCopyAll)

			idSaveAll = wx.NewIdRef()
			menu.Append(idSaveAll, "Zapisz całą odpowiedź do pliku...\tS")
			self.Bind(wx.EVT_MENU, lambda evt, n=aNum, t=turn: self.onSaveSpecificAnswer(n, t), id=idSaveAll)

			if sections:
				menu.AppendSeparator()
				for sIdx, (secTitle, secContent) in enumerate(sections, 1):
					snipTitle = secTitle[:60] + ("..." if len(secTitle) > 60 else "")
					idSec = wx.NewIdRef()
					menu.Append(idSec, f"Kopiuj sekcję {sIdx}: \"{snipTitle}\"")
					self.Bind(
						wx.EVT_MENU,
						lambda evt, st=secTitle, sc=secContent: self.onCopySection(st, sc),
						id=idSec
					)
		else:
			for aNum, turn in modelTurns:
				text = turn.get("text", "")
				sections = extractAnswerSections(text)
				snip = text.strip()[:35].replace("\n", " ")

				ansSub = wx.Menu()
				idCopyWhole = wx.NewIdRef()
				ansSub.Append(idCopyWhole, f"Kopiuj całą odpowiedź nr {aNum}")
				self.Bind(wx.EVT_MENU, lambda evt, n=aNum, t=turn: self.onCopySpecificAnswer(n, t), id=idCopyWhole)

				idSaveWhole = wx.NewIdRef()
				ansSub.Append(idSaveWhole, f"Zapisz odpowiedź nr {aNum} do pliku...")
				self.Bind(wx.EVT_MENU, lambda evt, n=aNum, t=turn: self.onSaveSpecificAnswer(n, t), id=idSaveWhole)

				if sections:
					ansSub.AppendSeparator()
					for sIdx, (secTitle, secContent) in enumerate(sections, 1):
						snipTitle = secTitle[:50] + ("..." if len(secTitle) > 50 else "")
						idSec = wx.NewIdRef()
						ansSub.Append(idSec, f"Kopiuj sekcję {sIdx}: \"{snipTitle}\"")
						self.Bind(
							wx.EVT_MENU,
							lambda evt, st=secTitle, sc=secContent: self.onCopySection(st, sc),
							id=idSec
						)

				menu.AppendSubMenu(ansSub, f"Odpowiedź #{aNum} (\"{snip}...\")")

		self.PopupMenu(menu)
		menu.Destroy()

	def onCopySection(self, sectionTitle: str, sectionContent: str):
		"""Kopiuje konkretną wybraną sekcję (podrozdział) odpowiedzi AI do schowka."""
		if not sectionContent:
			ui.message("Wybrana sekcja nie zawiera treści.")
			return
		includeImg = self.chkIncludeImage.GetValue()
		imgPath = getOrSaveSourceImageFile(self.item) if includeImg else None

		plain = f"{sectionTitle}\r\n\r\n{sectionContent.strip()}"
		html = buildHtmlSectionReport(self.item, sectionTitle, sectionContent, includeImage=includeImg, imageFilePath=imgPath)
		if copyHtmlAndTextToClipboard(plain, html, imageFilePath=imgPath):
			msg = f"Skopiowano sekcję '{sectionTitle}' do schowka."
			if includeImg:
				msg += " (z grafiką źródłową)"
			ui.message(msg)

	def onCharHook(self, event):
		key = event.GetKeyCode()
		focused = self.FindFocus()

		# OCHRONA POLSKIEJ KLAWIATURY (Prawy Alt / AltGr):
		isAltGr = (event.ControlDown() and event.AltDown()) or bool(winUser.getKeyState(winUser.VK_RMENU) & 0x8000)
		if isAltGr:
			event.Skip()
			return

		# Obsługa Ctrl+Home i Ctrl+End — ZAWSZE przenosi na sam początek lub koniec rozmowy
		if event.ControlDown() and not event.AltDown() and not event.ShiftDown():
			if key == wx.WXK_HOME:
				self.historyCtrl.SetFocus()
				self.historyCtrl.SetInsertionPoint(0)
				self.historyCtrl.ShowPosition(0)
				firstHdr = self.headingTargets[0].title if self.headingTargets else "początek"
				ui.message(f"Początek rozmowy: {firstHdr}")
				return
			elif key == wx.WXK_END:
				self.historyCtrl.SetFocus()
				lastPos = self.historyCtrl.GetLastPosition()
				self.historyCtrl.SetInsertionPoint(lastPos)
				self.historyCtrl.ShowPosition(lastPos)
				ui.message("Koniec rozmowy")
				return

		# Klawisz Escape zamyka okno, F1 otwiera pomoc kontekstową
		if key == wx.WXK_ESCAPE:
			self.Close()
			return
		elif key == wx.WXK_F1:
			showContextHelp("chat", parent=self)
			return

		# Gdy fokus znajduje się w polu edycji pytania (pisanie tekstu)
		if focused == self.questionCtrl:
			if (key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER) and not event.ShiftDown()) or (key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER) and event.ControlDown()):
				self.onSend()
				return
			event.Skip()
			return

		# Obsługa klawisza Enter / Return w oknie historii (otwieranie linku pod kursorem)
		if focused == self.historyCtrl and key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
			curOffset = self.historyCtrl.GetInsertionPoint()
			link = self.findLinkAtOffset(curOffset)
			if link:
				self.activateLink(link)
				return
			else:
				ui.message("Kursor nie znajduje się na linku. Użyj klawisza K, aby przejść do linku, a następnie naciśnij Enter.")
				return

		# Szybka nawigacja w historii (działa WYŁĄCZNIE gdy fokus jest POZA polem pytania i BEZ modyfikatorów Ctrl/Alt):
		if not event.ControlDown() and not event.AltDown():
			if key in (ord('H'), ord('h')):
				self.navigateHeading(forward=not event.ShiftDown())
				return
			elif key in (ord('K'), ord('k')):
				self.navigateLink(forward=not event.ShiftDown())
				return
			elif key in (ord('L'), ord('l')):
				self.showLinksMenu()
				return
			elif key in (
				ord('1'), ord('2'), ord('3'), ord('4'), ord('5'), ord('6'),
				wx.WXK_NUMPAD1, wx.WXK_NUMPAD2, wx.WXK_NUMPAD3,
				wx.WXK_NUMPAD4, wx.WXK_NUMPAD5, wx.WXK_NUMPAD6,
			):
				levelMap = {
					ord('1'): 1, ord('2'): 2, ord('3'): 3, ord('4'): 4, ord('5'): 5, ord('6'): 6,
					wx.WXK_NUMPAD1: 1, wx.WXK_NUMPAD2: 2, wx.WXK_NUMPAD3: 3,
					wx.WXK_NUMPAD4: 4, wx.WXK_NUMPAD5: 5, wx.WXK_NUMPAD6: 6,
				}
				lvl = levelMap.get(key)
				if lvl:
					self.navigateHeading(forward=not event.ShiftDown(), targetLevel=lvl)
					return
			elif key in (ord('B'), ord('b')):
				self.navigateButton(forward=not event.ShiftDown())
				return
			elif key in (ord('E'), ord('e')):
				self.focusQuestionEdit()
				return
			elif key in (ord('C'), ord('c')):
				self.onCopyKey()
				return
			elif key in (ord('S'), ord('s')):
				self.onSaveKey()
				return

		event.Skip()

	def onCopySpecificAnswer(self, aNum: int, turn: dict):
		"""Kopiuje konkretną odpowiedź AI (tekst + opcjonalny obraz) do schowka."""
		if not turn or not turn.get("text"):
			ui.message("Brak treści odpowiedzi do skopiowania.")
			return
		includeImg = self.chkIncludeImage.GetValue()
		imgPath = getOrSaveSourceImageFile(self.item) if includeImg else None
		plain = turn["text"].strip()
		html = buildHtmlReport(self.item, [turn], includeImage=includeImg, imageFilePath=imgPath)
		if copyHtmlAndTextToClipboard(plain, html, imageFilePath=imgPath):
			ui.message(f"Skopiowano odpowiedź nr {aNum} do schowka." + (" (z grafiką źródłową)" if includeImg else ""))

	def onSaveSpecificAnswer(self, aNum: int, turn: dict):
		"""Zapisuje konkretną odpowiedź AI jako plik HTML, Markdown lub TXT."""
		if not turn or not turn.get("text"):
			ui.message("Brak treści odpowiedzi do zapisania.")
			return
		includeImg = self.chkIncludeImage.GetValue()
		isAtt = isinstance(self.item, AttachmentItem)
		if isAtt:
			cleanName = sanitizeFilename(os.path.splitext(self.item.fileName)[0])
			defaultName = f"Odpowiedz_{aNum}_{cleanName}.html"
		else:
			cName = getattr(self.item, "customName", "")
			if cName:
				cleanName = sanitizeFilename(cName)
				defaultName = f"Odpowiedz_{aNum}_{cleanName}.html"
			else:
				defaultName = f"Odpowiedz_{aNum}_Grafika_{self.item.index:02d}.html"

		defaultDir = getDefaultSaveDirectory(category="chat")
		with wx.FileDialog(
			self,
			message=f"Zapisz odpowiedź nr {aNum}",
			defaultDir=defaultDir,
			defaultFile=defaultName,
			wildcard="Plik raportu HTML (*.html)|*.html|Dokument Markdown (*.md)|*.md|Plik tekstowy (*.txt)|*.txt",
			style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
		) as dlg:
			if dlg.ShowModal() == wx.ID_CANCEL:
				return
			path = dlg.GetPath()

		try:
			if path.lower().endswith(".md"):
				content = buildPlainTextReport(self.item, [turn], includeImage=includeImg)
				with open(path, "w", encoding="utf-8") as f:
					f.write(content)
			elif path.lower().endswith(".txt"):
				with open(path, "w", encoding="utf-8") as f:
					f.write(turn.get("text", "").strip())
			else:
				htmlContent = buildHtmlReport(self.item, [turn], includeImage=includeImg)
				with open(path, "w", encoding="utf-8") as f:
					f.write(htmlContent)

			ui.message(f"Pomyślnie zapisano odpowiedź nr {aNum}: {os.path.basename(path)}")
		except Exception as ex:
			ui.message(f"Błąd podczas zapisu odpowiedzi: {ex}")

	def onSend(self):
		"""Wysyła kolejne pytanie do modelu AI i dołącza odpowiedź do historii."""
		question = self.questionCtrl.GetValue().strip()
		if not question:
			ui.message("Wpisz pytanie przed wysłaniem.")
			return

		if not getActiveAiApiKey():
			if not checkAndPromptAiApiKey(parentWin=self):
				return

		self.btnSend.Disable()
		self.questionCtrl.Disable()

		isFirst = len(self.item.chatHistory) == 0
		mimeType = "image/png"

		if isFirst:
			if isinstance(self.item, AttachmentItem):
				tempPath = saveAttachmentToTempFile(self.item)
				if self.item.isPdf:
					tempPath = ensurePdfUnlocked(self.item, parentWin=self)
					if not tempPath:
						ui.message("Operacja przerwana: dokument PDF jest zabezpieczony hasłem.")
						self.btnSend.Enable()
						self.questionCtrl.Enable()
						return
					mimeType = "application/pdf"
					if tempPath and os.path.isfile(tempPath):
						try:
							with open(tempPath, "rb") as f:
								self.item.base64Data = base64.b64encode(f.read()).decode("ascii")
						except Exception:
							pass
				elif self.item.isImage:
					mimeType = "image/png" if self.item.ext == ".png" else "image/jpeg"
					if tempPath and os.path.isfile(tempPath):
						try:
							with open(tempPath, "rb") as f:
								self.item.base64Data = base64.b64encode(f.read()).decode("ascii")
						except Exception:
							pass
				elif self.item.isDocx and tempPath:
					docxText = extractTextFromDocxFile(tempPath)
					if docxText:
						question = f"Oto treść załączonego dokumentu Word ({self.item.fileName}):\n\n{docxText}\n\n{question}"
				elif self.item.isText and tempPath:
					try:
						with open(tempPath, "r", encoding="utf-8", errors="replace") as f:
							txtContent = f.read(30000)
						question = f"Oto treść załączonego pliku ({self.item.fileName}):\n\n{txtContent}\n\n{question}"
					except Exception:
						pass
			elif isinstance(self.item, GraphicItem):
				if not self.item.base64Image:
					tempPng = extractGraphicToPngFile(self.item)
					if tempPng and os.path.isfile(tempPng):
						try:
							with open(tempPng, "rb") as f:
								self.item.base64Image = base64.b64encode(f.read()).decode("ascii")
						finally:
							try:
								os.remove(tempPng)
							except Exception:
								pass

		userTurn = {
			"role": "user",
			"text": question,
			"timestamp": time.strftime("%H:%M"),
		}
		if isFirst:
			b64 = getattr(self.item, "base64Data", "") or getattr(self.item, "base64Image", "")
			if b64:
				userTurn["imageData"] = b64
				userTurn["mimeType"] = mimeType

		self.item.chatHistory.append(userTurn)
		self.updateHistoryDisplay(scrollToEnd=True)
		self.questionCtrl.SetValue("")

		prov = getActiveAiProvider()
		provName = "ChatGPT" if prov == "openai" else "Gemini"
		feedback = AiProgressFeedback(f"Wysyłanie pytania do {provName}, proszę czekać...")

		def worker():
			try:
				ans = askAiMultiTurn(self.item.chatHistory, parentDialog=self)
				self.item.chatHistory.append({
					"role": "model",
					"text": ans,
					"timestamp": time.strftime("%H:%M"),
				})
				self.item.aiDescription = ans

				def updateUi():
					feedback.stop()
					playSuccessChime()
					if not self or not winUser.isWindow(getattr(self, "Handle", 0)):
						return
					try:
						self.btnSend.Enable()
						self.questionCtrl.Enable()
						self.updateHistoryDisplay(scrollToEnd=True)
					except Exception:
						pass
					firstSent = ans.split(".")[0] if "." in ans else ans[:80]
					ui.message(f"{provName}: {firstSent}")

				wx.CallAfter(updateUi)
			except Exception as ex:
				errStr = str(ex)
				log.error(f"outlookImageHelper: Błąd czatu AI: {errStr}", exc_info=True)
				if self.item.chatHistory:
					self.item.chatHistory.pop()

				def failUi(capturedErr=errStr):
					feedback.stop(f"Błąd AI: {capturedErr}")
					if not self or not winUser.isWindow(getattr(self, "Handle", 0)):
						return
					try:
						self.btnSend.Enable()
						self.questionCtrl.Enable()
						self.questionCtrl.SetValue(question)
						self.updateHistoryDisplay(scrollToEnd=True)
						self.questionCtrl.SetFocus()
					except Exception:
						pass

					if "503" in capturedErr or "Service Unavailable" in capturedErr:
						msg = (
							"BŁĄD POŁĄCZENIA Z SERWEREM AI (HTTP 503 SERVICE UNAVAILABLE)\n"
							+ "=" * 60 + "\n\n"
							f"Serwery usługi {provName} są w tej chwili tymczasowo przeciążone lub niedostępne.\n\n"
							"Co możesz zrobić:\n"
							"1. Odczekaj chwilę (ok. 15-30 sekund) i spróbuj zadać pytanie ponownie.\n"
							"2. W ustawieniach dodatku możesz wybrać inny model (np. Gemini 1.5 Flash lub Gemini 2.0 Flash) albo przełączyć się na dostawcę OpenAI (ChatGPT).\n"
							"3. Sprawdź czy połączenie internetowe działa stabilnie."
						)
					elif "429" in errStr:
						msg = (
							f"PRZEKROCZONO LIMIT ZAPYTAŃ DLA {provName.upper()} (HTTP 429 TOO MANY REQUESTS)\n"
							+ "=" * 60 + "\n\n"
							"Wyczerpano dopuszczalny limit bezpłatnych zapytań na minutę dla Twojego klucza API.\n"
							"Odczekaj minutę przed wysłaniem kolejnego pytania."
						)
					elif "401" in errStr or "403" in errStr:
						msg = (
							f"BŁĄD AUTORYZACJI KLUCZA API DLA {provName.upper()} (HTTP 401 / 403)\n"
							+ "=" * 60 + "\n\n"
							"Twój klucz API jest nieprawidłowy lub utracił ważność.\n"
							"Sprawdź i zaktualizuj klucz API w ustawieniach dodatku."
						)
					else:
						msg = f"Wystąpił błąd podczas komunikacji z modelem AI:\n\n{errStr}"
					showAccessibleMessageBox(msg, title=f"Błąd odpowiedzi AI ({provName})", parent=self)

				wx.CallAfter(failUi)

		threading.Thread(target=worker, daemon=True).start()

	def onCopyLast(self, event=None):
		if not self.item.chatHistory:
			ui.message("Brak odpowiedzi do skopiowania.")
			return
		modelTurns = [t for t in self.item.chatHistory if t.get("role") == "model"]
		if not modelTurns:
			ui.message("Brak odpowiedzi modelu AI.")
			return
		lastTurn = modelTurns[-1]
		includeImg = self.chkIncludeImage.GetValue()
		imgPath = getOrSaveSourceImageFile(self.item) if includeImg else None
		plain = lastTurn["text"]
		html = buildHtmlReport(self.item, [lastTurn], includeImage=includeImg, imageFilePath=imgPath)
		if copyHtmlAndTextToClipboard(plain, html, imageFilePath=imgPath):
			ui.message("Skopiowano ostatnią odpowiedź AI do schowka." + (" (z obiektem źródłowym)" if includeImg else ""))

	def onCopyAll(self, event=None):
		if not self.item.chatHistory:
			ui.message("Brak historii rozmowy do skopiowania.")
			return
		includeImg = self.chkIncludeImage.GetValue()
		imgPath = getOrSaveSourceImageFile(self.item) if includeImg else None
		plain = buildPlainTextReport(self.item, self.item.chatHistory, includeImage=includeImg)
		html = buildHtmlReport(self.item, self.item.chatHistory, includeImage=includeImg, imageFilePath=imgPath)
		if copyHtmlAndTextToClipboard(plain, html, imageFilePath=imgPath):
			ui.message(f"Skopiowano całą rozmowę ({len(self.item.chatHistory)} wypowiedzi) do schowka." + (" (z obiektem źródłowym)" if includeImg else ""))

	def onExportMenu(self, event=None):
		"""Wyświetla menu czynności eksportu: zapis do pliku, otwarcie folderu lub wysłanie e-mailem."""
		if not self.item.chatHistory:
			ui.message("Brak danych do wyeksportowania.")
			return
		menu = wx.Menu()

		idSave = wx.NewIdRef()
		menu.Append(idSave, "Zapisz całą rozmowę do pliku...\tS")
		self.Bind(wx.EVT_MENU, lambda evt: self.onExport(), id=idSave)

		idOpenDir = wx.NewIdRef()
		menu.Append(idOpenDir, "Otwórz folder z zapisanymi rozmowami w Eksploratorze")
		self.Bind(wx.EVT_MENU, lambda evt: self.onOpenSavedChatsFolder(), id=idOpenDir)

		idSendEmail = wx.NewIdRef()
		menu.Append(idSendEmail, "Wyślij rozmowę w nowej wiadomości e-mail (Outlook)")
		self.Bind(wx.EVT_MENU, lambda evt: self.onSendChatViaEmail(), id=idSendEmail)

		self.PopupMenu(menu)
		menu.Destroy()

	def onOpenSavedChatsFolder(self, event=None):
		folder = getDefaultSaveDirectory(category="chat")
		try:
			os.startfile(folder)
			ui.message(f"Otwarto folder: {os.path.basename(folder)}")
		except Exception as ex:
			ui.message(f"Błąd otwierania folderu: {ex}")

	def onSendChatViaEmail(self, event=None):
		if not self.item.chatHistory:
			ui.message("Brak treści do wysłania.")
			return
		includeImg = self.chkIncludeImage.GetValue()
		imgPath = getOrSaveSourceImageFile(self.item) if includeImg else None
		htmlContent = buildHtmlReport(self.item, self.item.chatHistory, includeImage=includeImg, imageFilePath=imgPath)
		plainContent = buildPlainTextReport(self.item, self.item.chatHistory, includeImage=includeImg)

		prov = getActiveAiProvider()
		provName = "ChatGPT" if prov == "openai" else "Gemini"
		isAtt = isinstance(self.item, AttachmentItem)
		subject = f"Rozmowa z {provName}: {self.item.fileName}" if isAtt else f"Rozmowa z {provName}: Grafika {self.item.index}"

		try:
			olApp = getOutlookApp()
			if olApp:
				mail = olApp.CreateItem(0)
				mail.Subject = subject
				mail.HTMLBody = htmlContent
				if includeImg and imgPath and os.path.isfile(imgPath):
					try:
						mail.Attachments.Add(imgPath)
					except Exception:
						pass
				mail.Display()
				ui.message("Utworzono nową wiadomość w programie Outlook z treścią rozmowy.")
				return
		except Exception as ex:
			log.warning(f"outlookImageHelper: onSendChatViaEmail COM error: {ex}")

		if copyHtmlAndTextToClipboard(plainContent, htmlContent, imageFilePath=imgPath):
			ui.message("Nie udało się bezpośrednio utworzyć wiadomości w Outlooku. Treść rozmowy skopiowano do schowka.")

	def onExport(self, event=None):
		if not self.item.chatHistory:
			ui.message("Brak danych do wyeksportowania.")
			return
		includeImg = self.chkIncludeImage.GetValue()
		isAtt = isinstance(self.item, AttachmentItem)
		defaultName = f"Raport_{sanitizeFilename(self.item.fileName)}" if isAtt else f"Raport_{self.item.index:02d}"
		defaultName += ".html"
		defaultDir = getDefaultSaveDirectory(category="chat")

		with wx.FileDialog(
			self,
			message="Zapisz raport całej analizy",
			defaultDir=defaultDir,
			defaultFile=defaultName,
			wildcard="Plik raportu HTML (*.html)|*.html|Dokument Markdown (*.md)|*.md",
			style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
		) as dlg:
			if dlg.ShowModal() == wx.ID_CANCEL:
				return
			path = dlg.GetPath()

		try:
			if path.lower().endswith(".md"):
				content = buildPlainTextReport(self.item, self.item.chatHistory, includeImage=includeImg)
				with open(path, "w", encoding="utf-8") as f:
					f.write(content)
			else:
				htmlContent = buildHtmlReport(self.item, self.item.chatHistory, includeImage=includeImg)
				with open(path, "w", encoding="utf-8") as f:
					f.write(htmlContent)

			ui.message(f"Pomyślnie wyeksportowano raport: {os.path.basename(path)}")
		except Exception as ex:
			ui.message(f"Błąd podczas eksportu raportu: {ex}")


# Alias wstecznej zgodności
AiImageChatDialog = AiConversationDialog


# ==============================================================================
# OKNO DIALOGOWE: KREATOR I TESTER KLUCZA API GEMINI
# ==============================================================================


def testGeminiApiKey(apiKey: str, model: str = "gemini-3.8-flash") -> tuple[bool, str]:
	"""Weryfikuje poprawność klucza API Google Gemini poprzez weryfikację uprawnień klucza i test modelu."""
	key = (apiKey or "").strip()
	if not key:
		return False, _T("wizard_key_empty")

	keyVerifiedViaList = False
	availModels = set()
	try:
		listReq = urllib.request.Request(
			"https://generativelanguage.googleapis.com/v1beta/models",
			headers={"x-goog-api-key": key},
		)
		with urllib.request.urlopen(listReq, timeout=8) as resp:
			listData = json.loads(resp.read().decode("utf-8"))
			if "models" in listData:
				keyVerifiedViaList = True
				for mInfo in listData.get("models", []):
					mName = str(mInfo.get("name", "")).replace("models/", "")
					if mName:
						availModels.add(mName)
	except urllib.error.HTTPError as he:
		if he.code == 400:
			return False, _T("wizard_test_err_invalid_key")
		elif he.code == 403:
			return False, _T("wizard_test_err_forbidden")
		elif he.code == 429:
			return True, _T("wizard_test_warn_quota")
	except Exception:
		pass

	payload = {
		"contents": [
			{
				"parts": [{"text": "Hello"}]
			}
		]
	}
	payloadBytes = json.dumps(payload).encode("utf-8")
	modelsToTry = _getGeminiCandidateModels(model)
	if availModels:
		filtered = [m for m in modelsToTry if m in availModels]
		if filtered:
			modelsToTry = filtered

	lastErr = ""
	for m in modelsToTry[:3]:
		url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent"
		req = urllib.request.Request(
			url,
			data=payloadBytes,
			headers={
				"Content-Type": "application/json",
				"x-goog-api-key": key,
			},
		)
		try:
			with urllib.request.urlopen(req, timeout=10) as resp:
				resData = json.loads(resp.read().decode("utf-8"))
				if "candidates" in resData:
					return True, _T("wizard_test_success", model=m)
		except urllib.error.HTTPError as he:
			if he.code == 400:
				return False, _T("wizard_test_err_invalid_key")
			elif he.code == 403:
				return False, _T("wizard_test_err_forbidden")
			elif he.code == 429:
				return True, _T("wizard_test_warn_quota")
			elif he.code in (404, 500, 502, 503, 504):
				lastErr = f"HTTP {he.code}"
				continue
			else:
				if keyVerifiedViaList:
					return True, _T("wizard_test_success", model=modelsToTry[0])
				return False, _T("wizard_test_err_http", code=he.code)
		except Exception as ex:
			lastErr = str(ex)
			continue

	if keyVerifiedViaList:
		activeM = modelsToTry[0] if modelsToTry else _DEFAULT_GEMINI_MODEL
		return True, _T("wizard_test_success", model=activeM)

	return False, _T("wizard_test_err_failed", err=lastErr or "brak odpowiedzi")


class GeminiApiKeyWizardDialog(wx.Dialog):
	"""Dostępny kreator umożliwiający pozyskanie, przetestowanie i zapisanie bezpłatnego klucza API Google Gemini."""

	def __init__(self, parent, onKeySavedCallback=None):
		super().__init__(
			parent,
			title=_T("wizard_title"),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.onKeySavedCallback = onKeySavedCallback

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		self.lblInstructions = wx.StaticText(self, label=_T("wizard_instructions_lbl"))
		mainSizer.Add(self.lblInstructions, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)

		self.infoTextCtrl = wx.TextCtrl(
			self,
			value=_T("wizard_info"),
			style=wx.TE_MULTILINE | wx.TE_READONLY,
			size=(520, 160),
		)
		self.infoTextCtrl.SetName(_T("wizard_instructions_lbl"))
		mainSizer.Add(self.infoTextCtrl, 0, wx.EXPAND | wx.ALL, 10)

		self.btnOpenBrowser = wx.Button(self, label=_T("wizard_btn_open"))
		mainSizer.Add(self.btnOpenBrowser, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		lblKey = wx.StaticText(self, label=_T("wizard_lbl_key"))
		mainSizer.Add(lblKey, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)

		currentKey = config.conf[CONFIG_SECTION].get("geminiApiKey", "")
		self.keyCtrl = wx.TextCtrl(self, value=currentKey, size=(520, -1), style=wx.TE_PASSWORD)
		self.keyCtrl.SetName(_T("wizard_lbl_key"))
		mainSizer.Add(self.keyCtrl, 0, wx.EXPAND | wx.ALL, 10)

		self.keyCtrlPlain = wx.TextCtrl(self, value=currentKey, size=(520, -1))
		self.keyCtrlPlain.SetName(_T("wizard_lbl_key"))
		self.keyCtrlPlain.Hide()
		mainSizer.Add(self.keyCtrlPlain, 0, wx.EXPAND | wx.ALL, 10)

		self.chkShowKey = wx.CheckBox(self, label="&Pokaż klucz API" if getEffectiveLanguage() == "pl" else "&Show API key")
		self.chkShowKey.Bind(wx.EVT_CHECKBOX, self.onToggleShowKey)
		mainSizer.Add(self.chkShowKey, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		actionSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnPaste = wx.Button(self, label=_T("wizard_btn_paste"))
		actionSizer.Add(self.btnPaste, 0, wx.RIGHT, 10)

		self.btnTest = wx.Button(self, label=_T("wizard_btn_test"))
		actionSizer.Add(self.btnTest, 0, wx.RIGHT, 10)
		mainSizer.Add(actionSizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnSave = wx.Button(self, label=_T("wizard_btn_save"))
		self.btnSave.SetDefault()
		btnSizer.Add(self.btnSave, 0, wx.ALL, 5)

		self.btnCancel = wx.Button(self, wx.ID_CANCEL, label="&Anuluj" if getEffectiveLanguage() == "pl" else "&Cancel")
		btnSizer.Add(self.btnCancel, 0, wx.ALL, 5)

		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
		self.SetSizer(mainSizer)
		self.Fit()
		self.CenterOnScreen()

		self.btnOpenBrowser.Bind(wx.EVT_BUTTON, self.onOpenBrowser)
		self.btnPaste.Bind(wx.EVT_BUTTON, self.onPaste)
		self.btnTest.Bind(wx.EVT_BUTTON, self.onTest)
		self.btnSave.Bind(wx.EVT_BUTTON, self.onSave)
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		self.infoTextCtrl.SetFocus()

	def onToggleShowKey(self, event):
		if self.chkShowKey.GetValue():
			self.keyCtrlPlain.SetValue(self.keyCtrl.GetValue())
			self.keyCtrl.Hide()
			self.keyCtrlPlain.Show()
			self.keyCtrlPlain.SetFocus()
		else:
			self.keyCtrl.SetValue(self.keyCtrlPlain.GetValue())
			self.keyCtrlPlain.Hide()
			self.keyCtrl.Show()
			self.keyCtrl.SetFocus()
		self.Layout()

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.Close()
		else:
			event.Skip()

	def onOpenBrowser(self, event):
		try:
			webbrowser.open("https://aistudio.google.com/app/apikey")
			ui.message("Otwarto stronę Google AI Studio w przeglądarce.")
		except Exception as e:
			ui.message(f"Nie udało się otworzyć przeglądarki: {e}")

	def onPaste(self, event):
		try:
			pasted = ""
			if wx.TheClipboard.Open():
				try:
					textData = wx.TextDataObject()
					if wx.TheClipboard.GetData(textData):
						pasted = textData.GetText().strip()
				finally:
					wx.TheClipboard.Close()
			if pasted:
				self.keyCtrl.SetValue(pasted)
				self.keyCtrlPlain.SetValue(pasted)
				ui.message(_T("wizard_pasted_msg", chars=len(pasted)))
			else:
				ui.message(_T("wizard_clipboard_empty"))
		except Exception as e:
			log.debug(f"outlookImageHelper: onPaste error: {e}")

	def onTest(self, event):
		val = (self.keyCtrlPlain.GetValue() if self.chkShowKey.GetValue() else self.keyCtrl.GetValue()).strip()
		if not val:
			ui.message(_T("wizard_key_empty"))
			return
		testMsg = _T("wizard_testing") or "Testowanie połączenia z Google Gemini, proszę czekać..."
		progress = AiProgressFeedback(
			startMessage=testMsg,
			ongoingMessage="Nadal trwa testowanie klucza API Google Gemini, proszę czekać..."
		)
		self.btnTest.Disable()
		model = config.conf[CONFIG_SECTION].get("geminiModel", _DEFAULT_GEMINI_MODEL)

		def worker():
			ok = False
			msg = ""
			try:
				ok, msg = testGeminiApiKey(val, model=model)
			except Exception as ex:
				msg = f"Błąd testowania: {ex}"
			finally:
				def finish():
					self.btnTest.Enable()
					progress.stop(finishMessage=msg)
					try:
						if ok:
							tones.beep(880, 50)
							tones.beep(1175, 70)
						else:
							tones.beep(330, 150)
					except Exception:
						pass
				wx.CallAfter(finish)

		threading.Thread(target=worker, daemon=True).start()

	def onSave(self, event):
		val = (self.keyCtrlPlain.GetValue() if self.chkShowKey.GetValue() else self.keyCtrl.GetValue()).strip()
		if not val:
			ui.message(_T("wizard_key_empty"))
			return
		config.conf[CONFIG_SECTION]["geminiApiKey"] = val
		try:
			config.conf.save()
		except Exception:
			pass
		self.Close()
		ui.message(_T("wizard_saved_msg"))
		if self.onKeySavedCallback:
			def doCallback():
				try:
					self.onKeySavedCallback(val)
				except TypeError:
					self.onKeySavedCallback()
			wx.CallLater(100, doCallback)


# Alias wstecznej zgodności
GeminiApiKeyDialog = GeminiApiKeyWizardDialog


def checkAndPromptAiApiKey(parentWin=None, onKeyReadyCallback=None) -> bool:
	"""
	Sprawdza, czy klucz API dla aktywnego dostawcy AI jest skonfigurowany.
	Jeśli nie, wyświetla przyjazne i w pełni dostępne pytanie z opcją natychmiastowego
	otwarcia modalnego kreatora bezpłatnego klucza Google Gemini (lub OpenAI).
	Zwraca True jeśli klucz jest gotowy, False jeśli użytkownik anulował konfigurację.
	"""
	prov = getActiveAiProvider()
	apiKey = getActiveAiApiKey()
	if apiKey:
		return True

	provName = "OpenAI (ChatGPT)" if prov == "openai" else "Google Gemini"
	dlgParent = parentWin or gui.mainFrame

	msg = (
		f"Do wykonania tej analizy wymagany jest klucz API sztucznej inteligencji ({provName}).\n\n"
		"Osobisty klucz Google Gemini jest całkowicie BEZPŁATNY dla każdego konta Google.\n\n"
		"Czy chcesz teraz otworzyć prosty kreator i wprowadzić swój bezpłatny klucz API?"
	)

	dlg = wx.MessageDialog(
		dlgParent,
		message=msg,
		caption=f"Wymagany klucz API — {provName}",
		style=wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION
	)
	res = dlg.ShowModal()
	dlg.Destroy()

	if res == wx.ID_YES:
		if prov == "openai":
			showOpenAiKeyInstructions(dlgParent)
			return False
		else:
			def onSaved(newKey=None):
				if onKeyReadyCallback:
					wx.CallAfter(onKeyReadyCallback)
			wizard = GeminiApiKeyWizardDialog(dlgParent, onKeySavedCallback=onSaved)
			wizard.ShowModal()
			wizard.Destroy()
			return bool(getActiveAiApiKey())
	else:
		ui.message(f"Analiza AI anulowana — brak skonfigurowanego klucza API dla {provName}.")
		return False



# ==============================================================================
# LOKALIZACJA I GENEROWANIE WERSJI JĘZYKOWYCH PRZEZ AI / JSON
# ==============================================================================


def getLanguageTemplate(targetCode: str = "xx", targetName: str = "Example Language") -> dict:
	"""Zwraca słownik szablonu językowego ze wszystkimi kluczami interfejsu, pomocy i manifestu."""
	plStrings = dict(STRINGS.get("pl", {}))
	helpDict = {
		"list_title": "Pomoc: Lista grafik i załączników w wiadomości",
		"list_text": (
			"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC OKNA LISTY\n"
			+ "=" * 60 + "\n\n"
			"KATEGORIE ELEMENTÓW (Przełączaj grupą opcji lub strzałkami):\n"
			"  1. Grafiki i zrzuty ekranu w treści — rzeczywiste obrazy, wykresy i schematy\n"
			"  2. Linie podziału i separatory — poziome kreski i linie oddzielające w stopce\n"
			"  3. Załączniki do wiadomości — pliki załączone do maila (PDF, DOCX, XLSX, obrazy)\n\n"
			"SKRÓTY KLAWISZOWE W WIDOKU LISTY:\n"
			"  Enter         - Szybkie rozpoznanie tekstu (OCR) lub odczytanie treści pliku\n"
			"  P             - Powiększony podgląd grafiki (okno podglądu dla słabowidzących)\n"
			"  I             - Przełącz inwersję kolorów (negatyw) dla wybranej grafiki\n"
			"  Alt+C / Menu  - Menu czynności (OCR, opisywanie AI, zmiana nazwy, zapis, linki)\n"
			"  F2            - Nadaj własną, czytelną nazwę grafice (zapamiętywana w tej wiadomości)\n"
			"  Ctrl+Enter    - Otwórz link internetowy w przeglądarce lub otwórz plik załącznika\n"
			"  Tab           - Przejdź do pola szczegółów technicznych (wymiary, link, OCR, AI) i przycisków\n"
			"  Alt+O         - Szybki opis przez sztuczną inteligencję (AI)\n"
			"  Alt+R         - Rozpoznanie tekstu przez OCR\n"
			"  Alt+Z         - Zapisz wybrany element jako plik\n"
			"  Alt+W         - Zapisz wszystkie elementy do folderu\n"
			"  F1            - Wyświetl tę pomoc kontekstową\n"
			"  Escape        - Zamknij okno listy\n"
		),
		"chat_title": "Pomoc: Rozmowa z AI o grafice",
		"chat_text": (
			"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC ROZMOWY Z AI\n"
			+ "=" * 60 + "\n\n"
			"NAWIGACJA PO HISTORII ROZMOWY:\n"
			"  H / Shift+H   - Następny / Poprzedni nagłówek (dowolnego poziomu)\n"
			"  K / Shift+K   - Następny / Poprzedni link, e-mail, telefon lub pozycja GEO\n"
			"  Enter         - Otwórz link pod kursorem w domyślnej aplikacji\n"
			"  L             - Wyświetl listę wszystkich linków i kontaktów w rozmowie\n"
			"  2 / Shift+2   - Następne / Poprzednie pytanie (nagłówek H2)\n"
			"  3 / Shift+3   - Następna / Poprzednia odpowiedź AI (nagłówek H3)\n"
			"  4 / Shift+4   - Następne / Poprzednie wypunktowanie lub sekcja (nagłówek H4)\n"
			"  5 / Shift+5   - Następny / Poprzedni podpunkt (nagłówek H5)\n"
			"  B / Shift+B   - Następny / Poprzedni przycisk (kopiowanie i zapis pod odpowiedziami)\n"
			"  E             - Natychmiastowy skok do pola wprowadzania pytania\n"
			"  C             - Kopiuj wybraną lub ostatnią odpowiedź do schowka\n"
			"  S             - Zapisz wybraną lub ostatnią odpowiedź do pliku (HTML, MD, TXT)\n"
			"  Strzałki      - Czytaj treść odpowiedzi wiersz po wierszu\n\n"
			"ZADAWANIE PYTAŃ (w polu edycji pytania):\n"
			"  Enter / Ctrl+Enter - Wyślij pytanie do AI\n"
			"  Shift+Enter        - Nowa linia w treści pytania\n"
			"  Polskie znaki diakrytyczne (Prawy Alt+C, Prawy Alt+A itd.) są w pełni chronione.\n\n"
			"PRZYCISKI NA DOLE OKNA:\n"
			"  Kopiuj ostatnią odpowiedź, Kopiuj całą rozmowę, Eksportuj do pliku,\n"
			"  Dołącz grafikę źródłową (pole wyboru), Zamknij (Esc).\n\n"
			"  F1            - Wyświetl tę pomoc kontekstową\n"
			"  Escape        - Zamknij okno rozmowy\n"
		),
		"preview_title": "Pomoc: Powiększony podgląd grafiki",
		"preview_text": (
			"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC PODGLĄDU OBRAZU\n"
			+ "=" * 60 + "\n\n"
			"STEROWANIE POWIĘKSZENIEM I KONTRASTEM DLA OSÓB SŁABOWIDZĄCYCH:\n"
			"  + / = / Num+  - Zwiększ powiększenie (przybliż)\n"
			"  - / Num-      - Zmniejsz powiększenie (oddal)\n"
			"  0 / D         - Dopasuj cały obraz do rozmiaru okna\n"
			"  1             - Skala 100% (oryginalna rozdzielczość pikselowa)\n"
			"  I             - Przełącz inwersję kolorów (tryb negatywu / wysoki kontrast)\n"
			"  O             - Otwórz obraz w domyślnej przeglądarce graficznej Windows\n"
			"  Strzałki / PgUp / PgDn - Przewijaj powiększony obraz\n"
			"  F1            - Wyświetl tę pomoc kontekstową\n"
			"  Escape        - Zamknij okno podglądu\n"
		),
		"settings_title": "Pomoc: Ustawienia dodatku",
		"settings_text": (
			"OUTLOOK POMOCNIK OBRAZÓW I ZAŁĄCZNIKÓW — POMOC USTAWIEŃ\n"
			+ "=" * 60 + "\n\n"
			"DOSTĘPNE OPCJE:\n"
			"- Język interfejsu dodatku: Automatyczny (zgodny z NVDA), Polski, English oraz dodane języki\n"
			"- Domyślny język silnika OCR: Wybór języka rozpoznawania Windows Media OCR\n"
			"- Motyw dźwiękowy powiadomień: Skojarzeniowy, Minimalistyczny, Instrumentalny\n"
			"- Klucz API Google Gemini: Wymagany do multimodalnych opisów AI, czatu oraz tłumaczenia interfejsu\n"
			"- Model Gemini: Wybór modelu (domyślnie: gemini-3.8-flash)\n"
			"- Generator języków AI: Błyskawiczne tłumaczenie całego dodatku na dowolny język świata\n"
			"- Eksport / Import: Wygodna wymiana i edycja plików tłumaczeń JSON\n"
			"- Przycisk 'Skróty klawiszowe (Zdarzenia wejścia)': Konfiguracja globalnych gestów NVDA\n"
			"- Przycisk 'Instrukcja obsługi': Otwiera pełną dokumentację HTML w wybranym języku\n"
		)
	}
	manifestDict = {
		"summary": "Outlook - Pomocnik obrazów, załączników i OCR",
		"description": "Wykrywa, listuje, kopiuje oraz wykonuje OCR i interaktywne opisywanie/rozmowę przez AI o grafikach, zrzutach ekranu oraz załącznikach (pliki PDF, DOCX, obrazy itp.) w programie Microsoft Outlook Classic. Oferuje 3 kategorie elementów (grafiki, linie podziału, załączniki), powiększony podgląd z inwersją kolorów (negatyw) dla słabowidzących (P/I), pomoc kontekstową (F1), semantyczną nawigację nagłówkami i tabelami po odpowiedziach AI, menu czynności, szablony promptów, obsługę hiperłączy, czat wieloturowy, uniwersalny eksport oraz seryjny zapis z czytelnymi nazwami plików."
	}
	return {
		"langCode": targetCode,
		"langName": targetName,
		"strings": plStrings,
		"help": helpDict,
		"manifest": manifestDict,
	}


def exportLanguageTemplate(targetFilePath: str) -> bool:
	"""Eksportuje szablon tłumaczeniowy JSON do wskazanego pliku."""
	try:
		tmpl = getLanguageTemplate(targetCode="custom", targetName="Custom Language")
		with open(targetFilePath, "w", encoding="utf-8") as f:
			json.dump(tmpl, f, ensure_ascii=False, indent=2)
		return True
	except Exception as ex:
		log.error(f"outlookImageHelper: Błąd eksportu szablonu: {ex}")
		return False


def importLanguageFile(sourceFilePath: str) -> tuple[bool, str, str]:
	"""
	Importuje plik JSON z definicją języka i zapisuje go w CUSTOM_LANGUAGES_DIR.
	Zawiera ścisłą walidację kodu języka oraz weryfikację ścieżki (ochrona przed path traversal).
	Zwraca (sukces, kodJęzyka, nazwaJęzyka).
	"""
	try:
		with open(sourceFilePath, "r", encoding="utf-8") as f:
			data = json.load(f)
		if not isinstance(data, dict):
			return False, "", "Plik nie zawiera poprawnego obiektu JSON."
		code = str(data.get("langCode") or "").strip().lower()
		name = str(data.get("langName") or "").strip()
		strings = data.get("strings")
		if not code or not isinstance(strings, dict):
			return False, "", "Brak wymaganego pola 'langCode' lub słownika 'strings' w pliku JSON."

		# Ścisła walidacja kodu języka (np. 'pl', 'en-us', 'pt-br')
		if not re.match(r"^[a-z0-9]{2,5}(?:-[a-z0-9]{2,8})?$", code):
			return False, "", "Nieprawidłowy kod języka. Dozwolone są wyłącznie małe litery, cyfry i pojedynczy łącznik (np. 'es' lub 'pt-br')."

		if not name:
			name = code

		absDir = os.path.abspath(CUSTOM_LANGUAGES_DIR)
		os.makedirs(absDir, exist_ok=True)
		destPath = os.path.abspath(os.path.join(absDir, f"{code}.json"))

		# Weryfikacja czy ścieżka docelowa znajduje się ściśle w wyznaczonym folderze
		try:
			if os.path.commonpath([destPath, absDir]) != absDir:
				return False, "", "Odrzucono niebezpieczną ścieżkę zapisu pliku językowego."
		except Exception:
			if not destPath.startswith(absDir + os.sep):
				return False, "", "Odrzucono niebezpieczną ścieżkę zapisu pliku językowego."

		with open(destPath, "w", encoding="utf-8") as df:
			json.dump(data, df, ensure_ascii=False, indent=2)
		loadCustomLanguages()
		return True, code, name
	except Exception as ex:
		return False, "", str(ex)


def callGeminiTextOnly(
	prompt: str,
	apiKey: str,
	model: str = "gemini-3.8-flash",
	jsonMode: bool = False,
	timeout: int = 90
) -> str:
	"""Wysyła tekstowe zapytanie do API Google Gemini."""
	if not apiKey:
		raise ValueError("Brak klucza API Google Gemini.")
	payload = {
		"contents": [{"parts": [{"text": prompt}]}]
	}
	if jsonMode:
		payload["generationConfig"] = {"response_mime_type": "application/json"}

	payloadBytes = json.dumps(payload).encode("utf-8")
	modelsToTry = _getGeminiCandidateModels(model)

	lastError = None
	for candidateModel in modelsToTry:
		url = f"https://generativelanguage.googleapis.com/v1beta/models/{candidateModel}:generateContent"
		req = urllib.request.Request(
			url,
			data=payloadBytes,
			headers={
				"Content-Type": "application/json",
				"x-goog-api-key": apiKey,
			},
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				resData = json.loads(resp.read().decode("utf-8"))
			candidates = resData.get("candidates", [])
			if not candidates:
				raise RuntimeError("Model AI nie zwrócił odpowiedzi.")
			parts = candidates[0].get("content", {}).get("parts", [])
			texts = [p.get("text", "") for p in parts if "text" in p]
			result = "\n".join(texts).strip()
			if result:
				return result
		except urllib.error.HTTPError as he:
			lastError = he
			if he.code in (404, 503):
				continue
			raise he
		except Exception as e:
			lastError = e
			continue
	if lastError:
		raise lastError
	raise RuntimeError("Nie udało się uzyskać odpowiedzi od modelu Gemini.")


def _parseJsonFromGemini(raw: str) -> dict:
	clean = raw.strip()
	if clean.startswith("```"):
		clean = re.sub(r"^```(?:json)?\s*", "", clean)
		clean = re.sub(r"\s*```$", "", clean)
	clean = clean.strip()
	if not clean.startswith("{") and "{" in clean:
		clean = clean[clean.find("{"):]
	if not clean.endswith("}") and "}" in clean:
		clean = clean[:clean.rfind("}") + 1]
	return json.loads(clean)


def _translateHtmlManualInChunks(
	htmlContent: str,
	targetName: str,
	targetCode: str,
	apiKey: str,
	model: str = _DEFAULT_GEMINI_MODEL,
	progressCallback=None
) -> tuple[bool, str]:
	"""
	Tłumaczy podręcznik użytkownika HTML na język docelowy, dzieląc treść na mniejsze sekcje (chunki),
	co całkowicie eliminuje przekroczenia limitów czasu (timeout) i gwarantuje płynny postęp.
	"""
	if not htmlContent:
		return False, "Brak treści HTML do przetłumaczenia."

	# 1. Rozdziel na preBody (nagłówek, style), treść body oraz postBody (</body></html>)
	bodyMatch = re.search(r"^(.*?<body[^>]*>)(.*?)(</body>.*)$", htmlContent, flags=re.DOTALL | re.IGNORECASE)
	if bodyMatch:
		preBody = bodyMatch.group(1)
		bodyContent = bodyMatch.group(2)
		postBody = bodyMatch.group(3)
	else:
		preBody = f'<!DOCTYPE html>\n<html lang="{targetCode}">\n<head>\n<meta charset="utf-8">\n<title>Outlook Image Helper</title>\n</head>\n<body>'
		bodyContent = htmlContent
		postBody = "</body>\n</html>"

	# Zaktualizuj atrybut lang w tagu <html>
	preBody = re.sub(r'<html\s+lang=["\'][^"\']*["\']', f'<html lang="{targetCode}"', preBody, flags=re.IGNORECASE)

	# 2. Podziel treść body na sekcje według nagłówków <h2>
	rawSections = re.split(r"(?=<h2[\s>])", bodyContent, flags=re.IGNORECASE)

	# Połącz sekcje w zbalansowane bloki (chunki) do ok. 6000-7000 znaków
	chunks = []
	currentChunk = []
	currentLen = 0
	maxChunkSize = 6500

	for sec in rawSections:
		secLen = len(sec)
		if currentLen + secLen > maxChunkSize and currentChunk:
			chunks.append("".join(currentChunk))
			currentChunk = [sec]
			currentLen = secLen
		else:
			currentChunk.append(sec)
			currentLen += secLen
	if currentChunk:
		chunks.append("".join(currentChunk))

	totalChunks = len(chunks)
	translatedChunks = []

	for idx, chunk in enumerate(chunks, 1):
		if progressCallback:
			progressCallback(f"Krok 4: Tłumaczenie podręcznika HTML (sekcja {idx}/{totalChunks})...")

		promptChunk = (
			f"You are a professional technical translator specializing in screen reader software documentation.\n"
			f"Translate the following HTML snippet from Polish into {targetName} (language code: {targetCode}).\n"
			f"Strict rules:\n"
			f"1. Preserve ALL HTML tags, attributes (id, href, class, role, aria-label, etc.), CSS classes, table structure, and code tags completely intact.\n"
			f"2. Translate only the human-readable text, headings, table contents, and explanations.\n"
			f"3. Keep all keyboard shortcuts in <kbd> tags (e.g. <kbd>Ctrl+NVDA+A</kbd>, <kbd>Enter</kbd>, <kbd>P</kbd>, <kbd>I</kbd>, <kbd>F1</kbd>) exactly as they are.\n"
			f"4. Do NOT wrap output in markdown code fences (no ```html). Return ONLY the translated raw HTML snippet.\n\n"
			f"HTML snippet to translate:\n"
			f"{chunk}"
		)

		chunkTranslated = ""
		for attempt in range(2):
			try:
				rawSnippet = callGeminiTextOnly(promptChunk, apiKey=apiKey, model=model, jsonMode=False, timeout=90)
				cleanSnippet = rawSnippet.strip()
				if cleanSnippet.startswith("```"):
					cleanSnippet = re.sub(r"^```(?:html)?\s*", "", cleanSnippet)
					cleanSnippet = re.sub(r"\s*```$", "", cleanSnippet)
				chunkTranslated = cleanSnippet.strip()
				if chunkTranslated:
					break
			except Exception as ex:
				log.warning(f"outlookImageHelper: Błąd tłumaczenia sekcji {idx}/{totalChunks} (próba {attempt+1}): {ex}")
				if attempt == 0:
					time.sleep(1)

		# Jeśli tłumaczenie sekcji się nie powiodło po próbach, zachowaj oryginalną sekcję, by nie uszkodzić reszty podręcznika
		if not chunkTranslated:
			log.warning(f"outlookImageHelper: Zachowano oryginalną treść sekcji {idx}/{totalChunks}")
			chunkTranslated = chunk

		translatedChunks.append(chunkTranslated)

	# 3. Złóż całość z powrotem
	fullDoc = f"{preBody}\n{''.join(translatedChunks)}\n{postBody}"

	# Zaktualizuj pasek wyboru języków w nagłówku dokumentu (lang-switch)
	langSwitchPattern = r'<div class="lang-switch"[^>]*>.*?</div>'
	newLangSwitch = (
		f'<div class="lang-switch" role="navigation" aria-label="Wybór języka / Language selection">\n'
		f'  <span>Język / Language:</span>\n'
		f'  <a href="doc_pl.html" lang="pl">Polski</a>\n'
		f'  <a href="doc_en.html" lang="en">English</a>\n'
		f'  <span class="active" lang="{targetCode}">{targetName}</span>\n'
		f'</div>'
	)
	fullDoc = re.sub(langSwitchPattern, newLangSwitch, fullDoc, count=1, flags=re.DOTALL)

	# 4. Zapisz plik doc_<code>.html w CUSTOM_LANGUAGES_DIR oraz w ADDON_DIR
	targetUserDoc = os.path.join(CUSTOM_LANGUAGES_DIR, f"doc_{targetCode}.html")
	try:
		with open(targetUserDoc, "w", encoding="utf-8") as f:
			f.write(fullDoc)
	except Exception as ex:
		return False, f"Błąd zapisu pliku doc_{targetCode}.html: {ex}"

	try:
		targetAddonDoc = os.path.join(ADDON_DIR, f"doc_{targetCode}.html")
		with open(targetAddonDoc, "w", encoding="utf-8") as f:
			f.write(fullDoc)
	except Exception:
		pass

	return True, f"Pomyślnie przetłumaczono podręcznik HTML ({totalChunks} sekcji) i zapisano jako doc_{targetCode}.html."


def generateLanguageWithGemini(
	targetLangName: str,
	targetLangCode: str,
	doManifest: bool = True,
	doStrings: bool = True,
	doHelp: bool = True,
	doDocs: bool = False,
	incremental: bool = True,
	apiKey: str = "",
	model: str = _DEFAULT_GEMINI_MODEL,
	progressCallback=None,
	translateDocs: bool | None = None
) -> tuple[bool, str]:
	"""
	Generuje wersję językową (metadane, napisy interfejsu, pomoc F1 i opcjonalnie dokumentację HTML)
	za pomocą Gemini AI w modułowych krokach z pełnym wsparciem tłumaczenia różnicowego/przyrostowego
	oraz sekcyjnego tłumaczenia podręcznika HTML (zero timeoutów).
	Zapisuje wyniki na stałe w CUSTOM_LANGUAGES_DIR.
	"""
	if translateDocs is not None:
		doDocs = translateDocs

	code = targetLangCode.strip().lower()
	name = targetLangName.strip()
	if not code or not name:
		return False, "Nie podano kodu lub nazwy języka docelowego."
	if not apiKey:
		sec = config.conf.get(CONFIG_SECTION, {})
		apiKey = (sec.get("geminiApiKey", "") or "").strip()
	if not apiKey:
		return False, "Brak klucza API Google Gemini. Skonfiguruj go w ustawieniach dodatku."

	sourceTemplate = getLanguageTemplate(targetCode=code, targetName=name)
	os.makedirs(CUSTOM_LANGUAGES_DIR, exist_ok=True)
	jsonPath = os.path.join(CUSTOM_LANGUAGES_DIR, f"{code}.json")

	existingData = {}
	if os.path.isfile(jsonPath):
		try:
			with open(jsonPath, "r", encoding="utf-8") as f:
				loaded = json.load(f)
				if isinstance(loaded, dict):
					existingData = loaded
		except Exception as ex:
			log.warning(f"outlookImageHelper: Nie można odczytać istniejącego pliku {code}.json: {ex}")

	translatedData = {
		"langCode": code,
		"langName": name,
		"strings": dict(existingData.get("strings", {})) if incremental else {},
		"help": dict(existingData.get("help", {})) if incremental else {},
		"manifest": dict(existingData.get("manifest", {})) if incremental else {},
	}

	selectedSteps = [s for s, en in [("manifest", doManifest), ("strings", doStrings), ("help", doHelp), ("docs", doDocs)] if en]
	totalSteps = len(selectedSteps)
	if totalSteps == 0:
		return False, "Nie wybrano żadnego elementu do tłumaczenia."
	stepNum = 0

	# KROK: Metadane dodatku (opis i podsumowanie w manifest.ini)
	if doManifest:
		stepNum += 1
		hasManifest = (
			incremental and
			isinstance(translatedData.get("manifest"), dict) and
			translatedData["manifest"].get("summary") and
			translatedData["manifest"].get("description")
		)
		if hasManifest:
			if progressCallback:
				progressCallback(f"Krok {stepNum}/{totalSteps}: Metadane dodatku (manifest.ini) są już przetłumaczone (pominięto).")
		else:
			if progressCallback:
				progressCallback(f"Krok {stepNum}/{totalSteps}: Tłumaczenie opisu i metadanych dodatku...")

			promptManifest = (
				f"You are an expert translator specializing in screen reader software accessibility.\n"
				f"Translate the following add-on summary and description from Polish into {name} (language code: {code}).\n"
				f"Strict guidelines:\n"
				f"1. Return ONLY a valid JSON object with exact keys: 'summary' and 'description'.\n"
				f"2. Translate faithfully into {name} for NVDA users.\n\n"
				f"{json.dumps(sourceTemplate['manifest'], ensure_ascii=False, indent=2)}"
			)
			try:
				rawManifest = callGeminiTextOnly(promptManifest, apiKey=apiKey, model=model, jsonMode=True, timeout=60)
				manObj = _parseJsonFromGemini(rawManifest)
				if not isinstance(translatedData.get("manifest"), dict):
					translatedData["manifest"] = {}
				translatedData["manifest"].update(manObj)
				if progressCallback:
					progressCallback(f"Krok {stepNum}/{totalSteps}: Pomyślnie przetłumaczono metadane dodatku.")
			except Exception as ex:
				log.warning(f"outlookImageHelper: Błąd tłumaczenia manifestu: {ex}")
				if not translatedData.get("manifest"):
					translatedData["manifest"] = sourceTemplate["manifest"]
				if progressCallback:
					progressCallback(f"Krok {stepNum}/{totalSteps}: Ostrzeżenie przy manifest.ini: {ex}")

	# KROK: Napisy interfejsu użytkownika (UI Strings - ~176 fraz) z obsługą różnicową
	if doStrings:
		stepNum += 1
		if not isinstance(translatedData.get("strings"), dict):
			translatedData["strings"] = {}

		plStrings = sourceTemplate["strings"]
		existingStrings = translatedData["strings"]

		if incremental:
			missingStrings = {k: v for k, v in plStrings.items() if k not in existingStrings or not str(existingStrings[k]).strip()}
		else:
			missingStrings = dict(plStrings)

		if not missingStrings:
			if progressCallback:
				progressCallback(f"Krok {stepNum}/{totalSteps}: Wszystkie napisy interfejsu ({len(plStrings)} fraz) są już przetłumaczone (pominięto).")
		else:
			if progressCallback:
				if incremental and len(missingStrings) < len(plStrings):
					progressCallback(f"Krok {stepNum}/{totalSteps}: Tłumaczenie przyrostowe {len(missingStrings)} nowych/brakujących fraz (z {len(plStrings)})...")
				else:
					progressCallback(f"Krok {stepNum}/{totalSteps}: Tłumaczenie napisów interfejsu ({len(missingStrings)} fraz)...")

			promptStrings = (
				f"You are an expert translator specializing in NVDA screen reader software accessibility.\n"
				f"Translate the following dictionary of user interface strings from Polish into {name} (language code: {code}).\n"
				f"Strict guidelines:\n"
				f"1. Return ONLY a valid JSON object with the exact same keys and translated string values.\n"
				f"2. Preserve all keyboard shortcuts (e.g. Alt+C, Enter, Esc, F1, F2), format specifiers (e.g. {{name}}, {{count}}, {{idx}}), and accelerator ampersands (e.g. &Kategoria, &Actions).\n"
				f"3. Use standard {name} terminology for NVDA and screen readers.\n\n"
				f"{json.dumps(missingStrings, ensure_ascii=False, indent=2)}"
			)
			try:
				rawStrings = callGeminiTextOnly(promptStrings, apiKey=apiKey, model=model, jsonMode=True, timeout=90)
				transStrings = _parseJsonFromGemini(rawStrings)
				translatedData["strings"].update(transStrings)
				if progressCallback:
					progressCallback(f"Krok {stepNum}/{totalSteps}: Pomyślnie przetłumaczono {len(transStrings)} fraz interfejsu.")
			except Exception as ex:
				log.error(f"outlookImageHelper: Błąd tłumaczenia napisów UI: {ex}", exc_info=True)
				return False, f"Błąd tłumaczenia napisów interfejsu: {ex}"

	# KROK: Pomoc kontekstowa F1 (okno listy, czat, podgląd, ustawienia)
	if doHelp:
		stepNum += 1
		if not isinstance(translatedData.get("help"), dict):
			translatedData["help"] = {}

		plHelp = sourceTemplate["help"]
		existingHelp = translatedData["help"]

		if incremental:
			missingHelp = {k: v for k, v in plHelp.items() if k not in existingHelp or not str(existingHelp[k]).strip()}
		else:
			missingHelp = dict(plHelp)

		if not missingHelp:
			if progressCallback:
				progressCallback(f"Krok {stepNum}/{totalSteps}: Pomoc kontekstowa F1 jest już w pełni przetłumaczona (pominięto).")
		else:
			if progressCallback:
				progressCallback(f"Krok {stepNum}/{totalSteps}: Tłumaczenie pomocy kontekstowej F1 ({len(missingHelp)} sekcji)...")

			promptHelp = (
				f"You are an expert translator specializing in accessibility for blind computer users.\n"
				f"Translate the following contextual help texts from Polish into {name} (language code: {code}).\n"
				f"Strict guidelines:\n"
				f"1. Return ONLY a valid JSON object with identical keys from input.\n"
				f"2. Preserve all keyboard shortcuts (H, 2, 3, 4, 5, B, E, C, S, F1, Esc), dividers ('===', '---'), and formatting.\n\n"
				f"{json.dumps(missingHelp, ensure_ascii=False, indent=2)}"
			)
			try:
				rawHelp = callGeminiTextOnly(promptHelp, apiKey=apiKey, model=model, jsonMode=True, timeout=90)
				transHelp = _parseJsonFromGemini(rawHelp)
				translatedData["help"].update(transHelp)
				if progressCallback:
					progressCallback(f"Krok {stepNum}/{totalSteps}: Pomyślnie przetłumaczono pomoc kontekstową F1.")
			except Exception as ex:
				log.warning(f"outlookImageHelper: Błąd tłumaczenia pomocy kontekstowej: {ex}")
				if not translatedData.get("help"):
					translatedData["help"] = sourceTemplate["help"]
				if progressCallback:
					progressCallback(f"Krok {stepNum}/{totalSteps}: Ostrzeżenie przy pomocy F1: {ex}")

	# Zapisz natychmiast stan pliku JSON języka w CUSTOM_LANGUAGES_DIR
	if doManifest or doStrings or doHelp:
		try:
			with open(jsonPath, "w", encoding="utf-8") as f:
				json.dump(translatedData, f, ensure_ascii=False, indent=2)
		except Exception as ex:
			return False, f"Błąd zapisu pliku {code}.json: {ex}"

		# Zapisz manifest.ini w ADDON_DIR/locale/<code>/manifest.ini
		try:
			locDir = os.path.join(ADDON_DIR, "locale", code)
			os.makedirs(locDir, exist_ok=True)
			manData = translatedData.get("manifest", {})
			sumText = manData.get("summary", "")
			descText = manData.get("description", "")
			if sumText or descText:
				with open(os.path.join(locDir, "manifest.ini"), "w", encoding="utf-8") as mf:
					mf.write(f'summary = "{sumText}"\n')
					mf.write(f'description = """{descText}"""\n')
		except Exception as ex:
			log.warning(f"outlookImageHelper: Nie udało się zapisać manifest.ini w locale/{code}: {ex}")

		loadCustomLanguages()

	# KROK: Tłumaczenie podręcznika użytkownika HTML w zbalansowanych sekcjach
	if doDocs:
		stepNum += 1
		docSrcPath = os.path.join(ADDON_DIR, "doc_pl.html")
		if not os.path.isfile(docSrcPath):
			docSrcPath = os.path.join(ADDON_DIR, "doc.html")
		if os.path.isfile(docSrcPath):
			try:
				with open(docSrcPath, "r", encoding="utf-8") as df:
					htmlContent = df.read()
				okDoc, msgDoc = _translateHtmlManualInChunks(
					htmlContent=htmlContent,
					targetName=name,
					targetCode=code,
					apiKey=apiKey,
					model=model,
					progressCallback=progressCallback
				)
				if progressCallback:
					progressCallback(f"Krok {stepNum}/{totalSteps}: {msgDoc}")
			except Exception as ex:
				warnMsg = f"Ostrzeżenie: Błąd tłumaczenia podręcznika HTML: {ex}."
				log.warning(f"outlookImageHelper: {warnMsg}")
				if progressCallback:
					progressCallback(f"Krok {stepNum}/{totalSteps}: {warnMsg}")

	return True, f"Pomyślnie zaktualizowano wersję językową '{name}' ({code})."


class TranslationProgressTicker:
	"""Zapewnia dyskretne dźwięki postępu i okresowe przypomnienia głosowe podczas tłumaczenia AI."""

	def __init__(self, onTickVoiceCallback=None):
		self.running = True
		self.step = 0
		self.onTickVoiceCallback = onTickVoiceCallback
		self.currentStatus = "Trwa łączenie z serwerem Google AI..."
		try:
			tones.beep(880, 40)
		except Exception:
			pass
		self._scheduleNext()

	def setStatus(self, statusText: str):
		self.currentStatus = statusText
		self.step = 0
		try:
			tones.beep(1046, 35)
		except Exception:
			pass

	def _scheduleNext(self):
		if not self.running:
			return
		wx.CallLater(1500, self._onTick)

	def _onTick(self):
		if not self.running:
			return
		self.step += 1
		# Subtelny, krótki ton (25 ms) co 1.5 sekundy
		try:
			freq = 660 + (self.step % 3) * 110
			tones.beep(freq, 25)
		except Exception:
			pass
		# Co ok. 6 sekund (4 kroki x 1.5s) przypomnienie głosowe
		if self.step % 4 == 0 and self.onTickVoiceCallback:
			try:
				self.onTickVoiceCallback(self.currentStatus)
			except Exception:
				pass
		self._scheduleNext()

	def stop(self, success: bool = True):
		self.running = False
		try:
			if success:
				tones.beep(880, 50)
				tones.beep(1175, 70)
			else:
				tones.beep(330, 150)
		except Exception:
			pass


class GenerateLanguageDialog(wx.Dialog):
	"""
	Dostępny dialog umożliwiający wygenerowanie lub przyrostową aktualizację wersji językowej
	przez Google Gemini z wyborem kroków (checkboxy), trybem sekwencyjnym i dziennikiem postępu.
	"""

	def __init__(self, parent, onGeneratedCallback=None):
		super().__init__(
			parent,
			title=_T("dlg_trans_title"),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
			size=(720, 640),
		)
		self.onGeneratedCallback = onGeneratedCallback
		self.ticker = None
		self.pendingSteps = []
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		lblTargetName = wx.StaticText(self, label=_T("dlg_trans_target_name"))
		mainSizer.Add(lblTargetName, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
		self.targetNameCtrl = wx.TextCtrl(self, value="Deutsch")
		self.targetNameCtrl.SetName(_T("dlg_trans_target_name"))
		mainSizer.Add(self.targetNameCtrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		lblTargetCode = wx.StaticText(self, label=_T("dlg_trans_target_code"))
		mainSizer.Add(lblTargetCode, 0, wx.LEFT | wx.RIGHT | wx.TOP, 5)
		self.targetCodeCtrl = wx.TextCtrl(self, value="de")
		self.targetCodeCtrl.SetName(_T("dlg_trans_target_code"))
		mainSizer.Add(self.targetCodeCtrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# Zakres i opcje tłumaczenia (Checkboxy)
		scopeBox = wx.StaticBox(self, label=_T("dlg_trans_scope_box"))
		scopeSizer = wx.StaticBoxSizer(scopeBox, wx.VERTICAL)

		self.chkIncremental = wx.CheckBox(self, label=_T("dlg_trans_chk_incremental"))
		self.chkIncremental.SetValue(True)
		self.chkIncremental.SetName(self.chkIncremental.GetLabel())
		scopeSizer.Add(self.chkIncremental, 0, wx.ALL, 4)

		self.chkManifest = wx.CheckBox(self, label=_T("dlg_trans_chk_manifest"))
		self.chkManifest.SetValue(True)
		self.chkManifest.SetName(self.chkManifest.GetLabel())
		scopeSizer.Add(self.chkManifest, 0, wx.ALL, 4)

		self.chkStrings = wx.CheckBox(self, label=_T("dlg_trans_chk_strings"))
		self.chkStrings.SetValue(True)
		self.chkStrings.SetName(self.chkStrings.GetLabel())
		scopeSizer.Add(self.chkStrings, 0, wx.ALL, 4)

		self.chkHelp = wx.CheckBox(self, label=_T("dlg_trans_chk_help"))
		self.chkHelp.SetValue(True)
		self.chkHelp.SetName(self.chkHelp.GetLabel())
		scopeSizer.Add(self.chkHelp, 0, wx.ALL, 4)

		self.chkTranslateDoc = wx.CheckBox(self, label=_T("dlg_trans_chk_doc"))
		self.chkTranslateDoc.SetValue(False)
		self.chkTranslateDoc.SetName(self.chkTranslateDoc.GetLabel())
		scopeSizer.Add(self.chkTranslateDoc, 0, wx.ALL, 4)

		self.chkStepByStep = wx.CheckBox(self, label=_T("dlg_trans_chk_step_by_step"))
		self.chkStepByStep.SetValue(False)
		self.chkStepByStep.SetName(self.chkStepByStep.GetLabel())
		scopeSizer.Add(self.chkStepByStep, 0, wx.ALL, 4)

		mainSizer.Add(scopeSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# Dziennik postępu i historia komunikatów (dostępne pole tekstowe do czytania strzałkami i kopiowania)
		lblLog = wx.StaticText(self, label=_T("dlg_trans_log_lbl"))
		mainSizer.Add(lblLog, 0, wx.LEFT | wx.RIGHT | wx.TOP, 5)

		self.logCtrl = wx.TextCtrl(
			self,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL,
			size=(-1, 160),
		)
		self.logCtrl.SetName(_T("dlg_trans_log_lbl"))
		self.logCtrl.SetValue(
			"Wybierz zakres tłumaczenia powyżej, a następnie kliknij 'Rozpocznij tłumaczenie AI'.\n"
			"Dzięki tłumaczeniu przyrostowemu istniejące pliki są aktualizowane błyskawicznie o brakujące frazy."
		)
		mainSizer.Add(self.logCtrl, 1, wx.EXPAND | wx.ALL, 10)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnStart = wx.Button(self, label=_T("dlg_trans_btn_start"))
		self.btnStart.SetDefault()
		self.btnStart.Bind(wx.EVT_BUTTON, self.onStart)
		btnSizer.Add(self.btnStart, 0, wx.RIGHT, 10)

		self.btnCopyLog = wx.Button(self, label=_T("dlg_trans_btn_copy_log"))
		self.btnCopyLog.Bind(wx.EVT_BUTTON, self.onCopyLog)
		btnSizer.Add(self.btnCopyLog, 0, wx.RIGHT, 10)

		self.btnCancel = wx.Button(self, id=wx.ID_CANCEL, label=_T("btn_close"))
		self.btnCancel.Bind(wx.EVT_BUTTON, self.onCancel)
		btnSizer.Add(self.btnCancel, 0)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.Bind(wx.EVT_CLOSE, self.onCancel)

		self.SetSizer(mainSizer)
		self.Fit()
		self.CenterOnScreen()
		self.targetNameCtrl.SetFocus()

	def _setCheckboxesEnabled(self, enabled: bool):
		self.chkIncremental.Enable(enabled)
		self.chkManifest.Enable(enabled)
		self.chkStrings.Enable(enabled)
		self.chkHelp.Enable(enabled)
		self.chkTranslateDoc.Enable(enabled)
		self.chkStepByStep.Enable(enabled)

	def onCopyLog(self, evt):
		txt = self.logCtrl.GetValue().strip()
		if txt and copyHtmlAndTextToClipboard(txt):
			ui.message("Skopiowano dziennik postępu do schowka.")

	def appendLog(self, text: str):
		timestamp = time.strftime("%H:%M:%S")
		line = f"[{timestamp}] {text}\n"
		self.logCtrl.AppendText(line)
		self.logCtrl.ShowPosition(self.logCtrl.GetLastPosition())

	def onCancel(self, evt):
		self.pendingSteps = []
		if self.ticker:
			self.ticker.stop(success=False)
			self.ticker = None
		self.Destroy()

	def onStart(self, evt):
		name = self.targetNameCtrl.GetValue().strip()
		code = self.targetCodeCtrl.GetValue().strip().lower()
		if not name or not code:
			ui.message("Wpisz poprawną nazwę oraz kod języka.")
			return

		sec = config.conf.get(CONFIG_SECTION, {})
		apiKey = (sec.get("geminiApiKey", "") or "").strip()
		if not apiKey:
			ui.message(_T("wizard_key_empty"))
			self.appendLog("BŁĄD: Brak skonfigurowanego klucza API Google Gemini.")
			return

		model = (sec.get("geminiModel", _DEFAULT_GEMINI_MODEL) or _DEFAULT_GEMINI_MODEL).strip()

		# Kontynuacja w trybie sekwencyjnym
		if self.pendingSteps:
			self._runNextPendingStep(name, code, apiKey, model)
			return

		# Rozpoczęcie nowej sesji tłumaczenia: zbierz wybrane kroki
		steps = []
		if self.chkManifest.GetValue():
			steps.append(("manifest", "1. Metadane dodatku"))
		if self.chkStrings.GetValue():
			steps.append(("strings", "2. Napisy interfejsu"))
		if self.chkHelp.GetValue():
			steps.append(("help", "3. Pomoc kontekstowa F1"))
		if self.chkTranslateDoc.GetValue():
			steps.append(("docs", "4. Podręcznik użytkownika HTML"))

		if not steps:
			ui.message("Zaznacz przynajmniej jeden krok do przetłumaczenia.")
			return

		self.logCtrl.SetValue("")
		self.appendLog(f"Rozpoczęcie tłumaczenia na język: {name} (kod: {code}, model: {model})...")
		if self.chkIncremental.GetValue():
			self.appendLog("Tryb: Tłumaczenie przyrostowe (różnicowe) aktywne.")

		if self.chkStepByStep.GetValue():
			self.pendingSteps = list(steps)
			self._runNextPendingStep(name, code, apiKey, model)
		else:
			self._runAllStepsBatch(name, code, apiKey, model)

	def _runAllStepsBatch(self, name: str, code: str, apiKey: str, model: str):
		self.btnStart.Enable(False)
		self.targetNameCtrl.Enable(False)
		self.targetCodeCtrl.Enable(False)
		self._setCheckboxesEnabled(False)

		self.ticker = TranslationProgressTicker(onTickVoiceCallback=lambda msg: ui.message(msg))

		def updateProgress(statusText):
			def _uiUpdate():
				if self.ticker:
					self.ticker.setStatus(statusText)
				self.appendLog(statusText)
				ui.message(statusText)
			wx.CallAfter(_uiUpdate)

		def worker():
			ok, res = generateLanguageWithGemini(
				targetLangName=name,
				targetLangCode=code,
				doManifest=self.chkManifest.GetValue(),
				doStrings=self.chkStrings.GetValue(),
				doHelp=self.chkHelp.GetValue(),
				doDocs=self.chkTranslateDoc.GetValue(),
				incremental=self.chkIncremental.GetValue(),
				apiKey=apiKey,
				model=model,
				progressCallback=updateProgress
			)
			def finish():
				try:
					if self.ticker:
						self.ticker.stop(success=ok)
						self.ticker = None

					self.btnStart.SetLabel(_T("dlg_trans_btn_start"))
					self.btnStart.Enable(True)
					self.targetNameCtrl.Enable(True)
					self.targetCodeCtrl.Enable(True)
					self._setCheckboxesEnabled(True)

					if ok:
						self.appendLog(f"SUKCES: {res}")
						ui.message(_T("dlg_trans_success", name=name, code=code))
						if self.onGeneratedCallback:
							self.onGeneratedCallback(code)
						self.logCtrl.SetFocus()
					else:
						self.appendLog(f"BŁĄD: {res}")
						ui.message(_T("dlg_trans_err", err=res))
						self.logCtrl.SetFocus()
				except Exception as ex:
					log.error(f"outlookImageHelper: Błąd w finish GenerateLanguageDialog: {ex}")
			wx.CallAfter(finish)

		threading.Thread(target=worker, daemon=True).start()

	def _runNextPendingStep(self, name: str, code: str, apiKey: str, model: str):
		if not self.pendingSteps:
			return
		stepKey, stepLabel = self.pendingSteps.pop(0)

		self.btnStart.Enable(False)
		self.targetNameCtrl.Enable(False)
		self.targetCodeCtrl.Enable(False)
		self._setCheckboxesEnabled(False)

		self.ticker = TranslationProgressTicker(onTickVoiceCallback=lambda msg: ui.message(msg))

		def updateProgress(statusText):
			def _uiUpdate():
				if self.ticker:
					self.ticker.setStatus(statusText)
				self.appendLog(statusText)
				ui.message(statusText)
			wx.CallAfter(_uiUpdate)

		def worker():
			ok, res = generateLanguageWithGemini(
				targetLangName=name,
				targetLangCode=code,
				doManifest=(stepKey == "manifest"),
				doStrings=(stepKey == "strings"),
				doHelp=(stepKey == "help"),
				doDocs=(stepKey == "docs"),
				incremental=self.chkIncremental.GetValue(),
				apiKey=apiKey,
				model=model,
				progressCallback=updateProgress
			)
			def finish():
				try:
					if self.ticker:
						self.ticker.stop(success=ok)
						self.ticker = None

					if not ok:
						self.pendingSteps = []
						self.btnStart.SetLabel(_T("dlg_trans_btn_start"))
						self.btnStart.Enable(True)
						self.targetNameCtrl.Enable(True)
						self.targetCodeCtrl.Enable(True)
						self._setCheckboxesEnabled(True)
						self.appendLog(f"BŁĄD: {res}")
						ui.message(f"Błąd kroku {stepLabel}: {res}")
						self.logCtrl.SetFocus()
						return

					if self.pendingSteps:
						nextKey, nextLabel = self.pendingSteps[0]
						self.btnStart.SetLabel(_T("dlg_trans_btn_next", step=nextLabel))
						self.btnStart.Enable(True)
						self.btnStart.SetFocus()
						stepDoneMsg = f"{stepLabel} zakończony pomyślnie. Naciśnij 'Dalej' lub Enter, aby rozpocząć: {nextLabel}."
						self.appendLog(stepDoneMsg)
						ui.message(stepDoneMsg)
					else:
						self.pendingSteps = []
						self.btnStart.SetLabel(_T("dlg_trans_btn_start"))
						self.btnStart.Enable(True)
						self.targetNameCtrl.Enable(True)
						self.targetCodeCtrl.Enable(True)
						self._setCheckboxesEnabled(True)
						successMsg = f"Wszystkie etapy tłumaczenia na język '{name}' ({code}) zostały pomyślnie zakończone!"
						self.appendLog(f"SUKCES: {successMsg}")
						ui.message(successMsg)
						if self.onGeneratedCallback:
							self.onGeneratedCallback(code)
						self.logCtrl.SetFocus()
				except Exception as ex:
					log.error(f"outlookImageHelper: Błąd w finish _runNextPendingStep: {ex}")
			wx.CallAfter(finish)

		threading.Thread(target=worker, daemon=True).start()



_pending_item_sound_call = None


def _cancel_item_sound():
	global _pending_item_sound_call
	if _pending_item_sound_call:
		try:
			_pending_item_sound_call.Stop()
		except Exception:
			pass
		_pending_item_sound_call = None


def playItemTypeSound(item: GraphicItem | AttachmentItem | None, debounce_ms: int = 100):
	"""
	Odtwarza subtelny dźwięk przypisany do typu elementu (ikona mediów, telefon, e-mail, link, linia podziału, zrzut).
	Stosuje opóźnienie debounce i przerywanie poprzednich timerów, aby szybkie przewijanie listy
	nie kolejkowało odtwarzania dźwięków na wolniejszych komputerach.
	"""
	global _pending_item_sound_call
	_cancel_item_sound()
	if not item:
		return

	def do_play():
		global _pending_item_sound_call
		_pending_item_sound_call = None
		try:
			sec = config.conf.get(CONFIG_SECTION, {})
			if not sec.get("playSound", True) or not sec.get("navSounds", True):
				return
			theme = sec.get("soundTheme", "associative")
			themeDir = os.path.join(SOUNDS_DIR, theme)
			soundName = ""
			if isinstance(item, GraphicItem):
				if getattr(item, "linkCategory", "") == "social":
					soundName = "social_media"
				elif getattr(item, "linkCategory", "") == "phone":
					soundName = "phone"
				elif getattr(item, "linkCategory", "") == "email":
					soundName = "email"
				elif getattr(item, "hyperlinkUrl", ""):
					soundName = "link"
				elif getattr(item, "isDivider", False):
					soundName = "divider"
				else:
					soundName = "image_alert"
			elif isinstance(item, AttachmentItem):
				soundName = "link"

			if soundName:
				vol = sec.get("soundVolume", 80)
				scaledPath = _get_scaled_sound_path(soundName, volume=vol)
				if scaledPath and os.path.isfile(scaledPath):
					nvwave.playWaveFile(scaledPath, asynchronous=True)
					_log_sound_event(soundName, vol, "Lista dialogowa elementów", "Odtworzono pomyślnie")
		except Exception as e:
			log.debug(f"outlookImageHelper: playItemTypeSound error: {e}")

	if debounce_ms > 0:
		_pending_item_sound_call = wx.CallLater(debounce_ms, do_play)
	else:
		do_play()


# ==============================================================================
# OKNO DIALOGOWE: LISTA GRAFIK W WIADOMOŚCI
# ==============================================================================
# OKNO DIALOGOWE: POWIĘKSZONY PODGLĄD GRAFIKI (DLA SŁABOWIDZĄCYCH)
# ==============================================================================


class ImagePreviewDialog(wx.Dialog):
	"""
	Dostępne okno powiększonego podglądu obrazu dla osób słabowidzących.
	Umożliwia płynne skalowanie obrazu (+/-), dopasowanie do okna (0), 100% (1),
	otwarcie w domyślnej przeglądarce Windows (O) oraz nawigację suwakami.
	"""

	def __init__(self, parent, item: GraphicItem | AttachmentItem, tempImagePath: str, startInverted: bool | None = None):
		self.item = item
		self.imagePath = tempImagePath
		self.zoomFactor = 1.0
		self.normalImage = None
		self.invertedImage = None

		if startInverted is not None:
			self.isInverted = bool(startInverted)
		else:
			self.isInverted = bool(getattr(item, "isInverted", False)) or isImageInverted(getattr(item, "msgKey", ""), getattr(item, "index", 0))

		try:
			if self.imagePath and os.path.isfile(self.imagePath):
				self.normalImage = wx.Image(self.imagePath, wx.BITMAP_TYPE_ANY)
				if self.normalImage.IsOk():
					self.invertedImage = invertWxImage(self.normalImage)
		except Exception as e:
			log.error(f"outlookImageHelper: Błąd ładowania obrazu do podglądu: {e}")

		self.originalImage = self.normalImage

		name = getattr(item, "customName", "") or getattr(item, "fileName", "") or f"Grafika {item.index}"
		title = _T("preview_title", name=name)

		super().__init__(
			parent,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX,
			size=(920, 700),
		)
		self.SetMinSize((520, 420))

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# Informacja o powiększeniu i skrótach
		self.lblInfo = wx.StaticText(
			self,
			label=self._getInfoText(),
		)
		mainSizer.Add(self.lblInfo, 0, wx.ALL | wx.EXPAND, 8)

		# Obszar przewijany na obraz z ciemnym tłem o wysokim kontraście
		self.scroller = wx.ScrolledWindow(self, style=wx.HSCROLL | wx.VSCROLL | wx.BORDER_SUNKEN)
		self.scroller.SetBackgroundColour(wx.Colour(30, 30, 30))
		self.scroller.SetScrollRate(20, 20)

		self.imageStatic = wx.StaticBitmap(self.scroller)
		scrollerSizer = wx.BoxSizer(wx.VERTICAL)
		scrollerSizer.Add(self.imageStatic, 0, wx.ALIGN_CENTER | wx.ALL, 10)
		self.scroller.SetSizer(scrollerSizer)

		mainSizer.Add(self.scroller, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

		# Pasek przycisków
		btnSizer = wx.WrapSizer(wx.HORIZONTAL)

		self.btnZoomIn = wx.Button(self, label=_T("preview_btn_in"))
		btnSizer.Add(self.btnZoomIn, 0, wx.ALL, 4)

		self.btnZoomOut = wx.Button(self, label=_T("preview_btn_out"))
		btnSizer.Add(self.btnZoomOut, 0, wx.ALL, 4)

		self.btnFit = wx.Button(self, label=_T("preview_btn_fit"))
		btnSizer.Add(self.btnFit, 0, wx.ALL, 4)

		self.btn100 = wx.Button(self, label=_T("preview_btn_100"))
		btnSizer.Add(self.btn100, 0, wx.ALL, 4)

		self.btnInvert = wx.Button(self, label=self._getInvertBtnLabel())
		btnSizer.Add(self.btnInvert, 0, wx.ALL, 4)

		self.btnOpenOS = wx.Button(self, label=_T("preview_btn_open_os"))
		btnSizer.Add(self.btnOpenOS, 0, wx.ALL, 4)

		self.btnHelp = wx.Button(self, label=_T("preview_btn_help"))
		btnSizer.Add(self.btnHelp, 0, wx.ALL, 4)

		self.btnClose = wx.Button(self, wx.ID_CANCEL, label=_T("preview_btn_close"))
		btnSizer.Add(self.btnClose, 0, wx.ALL, 4)

		mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

		self.SetSizer(mainSizer)
		self.CenterOnScreen()
		self.Maximize(True)

		# Zdarzenia
		self.btnZoomIn.Bind(wx.EVT_BUTTON, lambda evt: self.zoomStep(1.3))
		self.btnZoomOut.Bind(wx.EVT_BUTTON, lambda evt: self.zoomStep(1.0 / 1.3))
		self.btnFit.Bind(wx.EVT_BUTTON, lambda evt: self.fitToWindow())
		self.btn100.Bind(wx.EVT_BUTTON, lambda evt: self.setZoom(1.0))
		self.btnInvert.Bind(wx.EVT_BUTTON, lambda evt: self.toggleInvert())
		self.btnOpenOS.Bind(wx.EVT_BUTTON, lambda evt: self.openInSystemViewer())
		self.btnHelp.Bind(wx.EVT_BUTTON, lambda evt: showContextHelp("preview", parent=self))
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

		wx.CallAfter(self.fitToWindow)

	def _getInvertBtnLabel(self) -> str:
		return _T("preview_btn_invert_on") if self.isInverted else _T("preview_btn_invert_off")

	def _getInfoText(self) -> str:
		pct = int(round(self.zoomFactor * 100))
		activeImg = self.invertedImage if self.isInverted else self.normalImage
		origW = activeImg.GetWidth() if (activeImg and activeImg.IsOk()) else getattr(self.item, "widthPx", 0)
		origH = activeImg.GetHeight() if (activeImg and activeImg.IsOk()) else getattr(self.item, "heightPx", 0)
		invTag = f" [{_T('tag_inverted').strip('[]')}]" if self.isInverted else ""
		return _T("preview_help_lbl", zoom=pct, w=origW, h=origH, inv=invTag)

	def updateBitmap(self, announce: bool = True):
		activeImg = self.invertedImage if self.isInverted else self.normalImage
		if not activeImg or not activeImg.IsOk():
			return
		w = int(max(1, round(activeImg.GetWidth() * self.zoomFactor)))
		h = int(max(1, round(activeImg.GetHeight() * self.zoomFactor)))
		scaled = activeImg.Scale(w, h, wx.IMAGE_QUALITY_HIGH)
		bmp = wx.Bitmap(scaled)
		self.imageStatic.SetBitmap(bmp)
		self.scroller.FitInside()
		self.scroller.Layout()
		self.lblInfo.SetLabel(self._getInfoText())
		if announce:
			pct = int(round(self.zoomFactor * 100))
			ui.message(_T("preview_zoom_announced", zoom=pct))

	def setZoom(self, factor: float, announce: bool = True):
		self.zoomFactor = max(0.1, min(20.0, factor))
		self.updateBitmap(announce=announce)


	def zoomStep(self, multiplier: float):
		self.setZoom(self.zoomFactor * multiplier, announce=True)

	def fitToWindow(self):
		activeImg = self.invertedImage if self.isInverted else self.normalImage
		if not activeImg or not activeImg.IsOk():
			return
		cw, ch = self.scroller.GetClientSize()
		cw = max(cw - 30, 100)
		ch = max(ch - 30, 100)
		iw = activeImg.GetWidth()
		ih = activeImg.GetHeight()
		if iw > 0 and ih > 0:
			scaleW = cw / iw
			scaleH = ch / ih
			self.setZoom(min(scaleW, scaleH, 3.0), announce=True)

	def toggleInvert(self):
		"""Przełącza tryb inwersji kolorów (negatyw) w podglądzie oraz zapisuje stan dla tej grafiki."""
		self.isInverted = not self.isInverted
		msgKey = getattr(self.item, "msgKey", "")
		idx = getattr(self.item, "index", 0)
		if msgKey and idx:
			setImageInverted(msgKey, idx, self.isInverted)
		setattr(self.item, "isInverted", self.isInverted)
		self.btnInvert.SetLabel(self._getInvertBtnLabel())
		self.updateBitmap(announce=False)
		if self.isInverted:
			ui.message(_T("preview_invert_on"))
		else:
			ui.message(_T("preview_invert_off"))

	def openInSystemViewer(self):
		targetPath = self.imagePath
		if self.isInverted and self.imagePath and os.path.isfile(self.imagePath):
			tempInv = createInvertedPngFile(self.imagePath)
			if tempInv and os.path.isfile(tempInv):
				targetPath = tempInv
		if targetPath and os.path.isfile(targetPath):
			try:
				os.startfile(targetPath)
				ui.message(_T("preview_opened_os"))
			except Exception as e:
				ui.message(f"Błąd otwierania obrazu: {e}")

	def onCharHook(self, event):
		key = event.GetKeyCode()
		isAltGr = (event.ControlDown() and event.AltDown()) or bool(winUser.getKeyState(winUser.VK_RMENU) & 0x8000)
		if isAltGr:
			event.Skip()
			return

		if key == wx.WXK_ESCAPE:
			self.Close()
			return
		elif key == wx.WXK_F1:
			showContextHelp("preview", parent=self)
			return
		elif key in (ord('+'), ord('='), wx.WXK_NUMPAD_ADD):
			self.zoomStep(1.3)
			return
		elif key in (ord('-'), wx.WXK_NUMPAD_SUBTRACT):
			self.zoomStep(1.0 / 1.3)
			return
		elif key in (ord('0'), ord('D'), ord('d')):
			self.fitToWindow()
			return
		elif key in (ord('1'),):
			self.setZoom(1.0)
			return
		elif key in (ord('I'), ord('i')) and not event.ControlDown() and not event.AltDown():
			self.toggleInvert()
			return
		elif key in (ord('O'), ord('o')) and not event.ControlDown() and not event.AltDown():
			self.openInSystemViewer()
			return

		event.Skip()



def openGraphicListDialog():
	"""Otwiera okno dialogowe listy wszystkich grafik i załączników w bieżącej wiadomości."""
	selCount, sel, isMulti = getSelectedMessagesInfo()
	if isMulti:
		handleMultipleMessagesSelected(selCount, sel, action="list")
		return

	ticker = BusySoundTicker(startTone=True, interval=0.25, baseFreq=560, endTone=False).start()
	try:
		mailItem = getCurrentMailItem()
		doc = getActiveWordDocument(mailItem=mailItem)
		if not mailItem and doc:
			mailItem = getCurrentMailItem(doc=doc)

		items = []
		totalLen = 0
		if doc:
			minDim = config.conf[CONFIG_SECTION].get("minDimension", 15)
			items = getOrCollectGraphics(doc, minDimension=minDim, fast=False, forceRefresh=False)
			try:
				totalLen = len(doc.Content.Text)
			except Exception:
				pass

		attachments = []
		if mailItem:
			attachments = extractAttachmentsFromCurrentMessage(mailItem=mailItem, includeInline=False)
	finally:
		ticker.stop(success=True)

	if not items and not attachments:
		if isMessageListFocused() or not isMessageBodyFocused():
			handleReadingPaneOff(mailItem, action="list")
			return
		if mailItem:
			try:
				subj = str(getattr(mailItem, "Subject", "") or "").strip()
				if subj:
					ui.message(_T("msg_no_elements_subj", subj=subj))
					return
			except Exception:
				pass
		ui.message(_T("msg_no_elements"))
		return

	d = GraphicElementsDialog(gui.mainFrame, items, attachments=attachments, totalDocLength=totalLen)
	gui.mainFrame.prePopup()
	d.Show()
	gui.mainFrame.postPopup()
	wx.CallLater(70, lambda: d.notebook.SetFocus())
	if not items and attachments:
		wx.CallLater(100, lambda: d._focusActiveListBox())



def showGraphicActionsMenu(item: GraphicItem | AttachmentItem, parent=None, callbackRefresh=None):
	"""
	Wyświetla w pełni dostępne menu czynności dla wskazanej grafiki lub załącznika,
	zarówno z poziomu okna listy, jak i bezpośrednio z treści otwartej wiadomości e-mail.
	"""
	if not item:
		return

	parentWin = parent or gui.mainFrame
	menu = wx.Menu()

	idOcr = wx.NewIdRef()
	idPreview = wx.NewIdRef()
	idInvertToggle = wx.NewIdRef()
	idAiQuick = wx.NewIdRef()
	idAiPrompt = wx.NewIdRef()
	idAiChat = wx.NewIdRef()
	idRename = wx.NewIdRef()
	idClearName = wx.NewIdRef()
	idOpen = wx.NewIdRef()
	idSelect = wx.NewIdRef()
	idCopy = wx.NewIdRef()
	idSave = wx.NewIdRef()
	idOpenList = wx.NewIdRef()
	idInstructions = wx.NewIdRef()
	idSettings = wx.NewIdRef()
	idHelp = wx.NewIdRef()

	isAtt = isinstance(item, AttachmentItem)

	if not isAtt:
		# Menu dla grafiki w treści maila
		menu.Append(idOcr, _T("menu_ocr"))
		menu.Append(idAiQuick, _T("menu_ai_quick"))
		menu.Append(idAiPrompt, _T("menu_ai_prompt"))
		menu.Append(idAiChat, _T("menu_ai_chat"))
		menu.AppendSeparator()

		menu.Append(idPreview, _T("menu_preview"))
		invLabel = "Wyłącz inwersję kolorów (standardowe barwy)\tI" if getattr(item, "isInverted", False) else _T("menu_invert_toggle")
		menu.Append(idInvertToggle, invLabel)
		menu.AppendSeparator()

		menu.Append(idRename, _T("menu_rename"))
		if getattr(item, "customName", ""):
			menu.Append(idClearName, _T("menu_clear_name"))
		menu.AppendSeparator()

		if getattr(item, "hyperlinkUrl", ""):
			svc = getattr(item, "linkServiceName", "")
			dom = svc if svc else getCleanDomain(item.hyperlinkUrl)
			menu.Append(idOpen, _T("menu_open_link", domain=dom) if dom else _T("menu_open_link_simple"))

		menu.Append(idSelect, _T("menu_select_doc"))
		menu.Append(idCopy, _T("menu_copy_img"))
		menu.Append(idSave, _T("menu_save_file"))
		menu.AppendSeparator()
		menu.Append(idInstructions, "&Instrukcja: jak odblokować i pobrać obrazy w programie Outlook (Ctrl+Shift+W)\tAlt+P")
		if not isinstance(parent, wx.Dialog):
			menu.AppendSeparator()
			menu.Append(idOpenList, "&Otwórz pełną listę wszystkich grafik w oknie...\tCtrl+NVDA+A")
	else:
		# Menu dla załącznika
		isPdf = getattr(item, "isPdf", False)
		isDocx = getattr(item, "isDocx", False)
		isImage = getattr(item, "isImage", False)

		if isPdf:
			menu.Append(idAiQuick, "&Analiza / Streszczenie dokumentu PDF przez AI\tAlt+O")
			menu.Append(idAiChat, "&Rozmowa z AI o dokumencie PDF (Q&A)...")
			menu.Append(idAiPrompt, "Analiza PDF z &wyborem prompta...")
			menu.Append(idOcr, _T("btn_ocr_pdf") + "\tAlt+R")
		elif isDocx:
			menu.Append(idAiQuick, "&Analiza / Streszczenie dokumentu Word przez AI\tAlt+O")
			menu.Append(idAiChat, "&Rozmowa z AI o dokumencie Word (Q&A)...")
			menu.Append(idAiPrompt, "Analiza Word z &wyborem prompta...")
			menu.Append(idOcr, _T("btn_read_doc") + "\tAlt+R")
		elif isImage:
			menu.Append(idOcr, _T("menu_ocr"))
			menu.Append(idPreview, _T("menu_preview"))
			menu.Append(idInvertToggle, _T("menu_invert_toggle"))
			menu.Append(idAiQuick, _T("menu_ai_quick"))
			menu.Append(idAiChat, _T("menu_ai_chat"))
			menu.Append(idAiPrompt, _T("menu_ai_prompt"))
		else:
			menu.Append(idAiQuick, "&Analiza pliku przez AI\tAlt+O")
			menu.Append(idAiChat, "&Rozmowa z AI o pliku...")
			menu.Append(idOcr, "&Odczytaj zawartość tekstową\tAlt+R")

		menu.AppendSeparator()
		menu.Append(idOpen, "&Otwórz plik załącznika w aplikacji domyślnej\tCtrl+Enter")
		menu.Append(idSave, _T("btn_save_one_att") + "\tAlt+Z")

	menu.AppendSeparator()
	menu.Append(idSettings, "&Przejdź do ustawień dodatku...")
	menu.Append(idHelp, _T("menu_help"))

	# Powiązania zdarzeń
	def onOcr(evt):
		if not ensureGraphicReady(item, parent=parentWin, actionName="rozpoznania tekstu (OCR)"):
			return
		if GlobalPlugin.instance:
			GlobalPlugin.instance.triggerOcrForItem(item, parentDialog=parent if isinstance(parent, wx.Dialog) else None)
	parentWin.Bind(wx.EVT_MENU, onOcr, id=idOcr)

	def onAiQuick(evt):
		if not ensureGraphicReady(item, parent=parentWin, actionName="analizy i opisu przez AI"):
			return
		def doAi():
			if GlobalPlugin.instance:
				GlobalPlugin.instance.triggerAiDescriptionForItem(item, parentDialog=parent if isinstance(parent, wx.Dialog) else None, customPrompt="")
		if not getActiveAiApiKey():
			checkAndPromptAiApiKey(parentWin=parentWin, onKeyReadyCallback=doAi)
			return
		doAi()
	parentWin.Bind(wx.EVT_MENU, onAiQuick, id=idAiQuick)

	def onAiPrompt(evt):
		if not ensureGraphicReady(item, parent=parentWin, actionName="analizy i opisu przez AI"):
			return
		def doPrompt():
			with AiPromptSelectionDialog(parentWin, item) as dlg:
				if dlg.ShowModal() == wx.ID_OK:
					if GlobalPlugin.instance:
						GlobalPlugin.instance.triggerAiDescriptionForItem(item, parentDialog=parent if isinstance(parent, wx.Dialog) else None, customPrompt=dlg.selectedPrompt)
		if not getActiveAiApiKey():
			checkAndPromptAiApiKey(parentWin=parentWin, onKeyReadyCallback=doPrompt)
			return
		doPrompt()
	parentWin.Bind(wx.EVT_MENU, onAiPrompt, id=idAiPrompt)

	def onAiChat(evt):
		if not ensureGraphicReady(item, parent=parentWin, actionName="rozmowy z AI"):
			return
		if isinstance(item, AttachmentItem) and item.isPdf:
			if not ensurePdfUnlocked(item, parentWin=parentWin):
				return
		def doChat():
			d = AiConversationDialog(parentWin, item)
			d.ShowModal()
			d.Destroy()
			if callbackRefresh:
				callbackRefresh()
		if not getActiveAiApiKey():
			checkAndPromptAiApiKey(parentWin=parentWin, onKeyReadyCallback=doChat)
			return
		doChat()
	parentWin.Bind(wx.EVT_MENU, onAiChat, id=idAiChat)

	def onPreview(evt):
		if not ensureGraphicReady(item, parent=parentWin, actionName="wyświetlenia powiększonego podglądu"):
			return
		if isAtt:
			tempPath = saveAttachmentToTempFile(item)
		else:
			tempPath = extractGraphicToPngFile(item)
		if not tempPath or not os.path.isfile(tempPath):
			showAccessibleMessageBox(_T("err_extract_img"), title="Błąd podglądu", parent=parentWin)
			return
		gui.mainFrame.prePopup()
		try:
			d = ImagePreviewDialog(parentWin, item, tempPath)
			d.ShowModal()
			d.Destroy()
		finally:
			gui.mainFrame.postPopup()
		if callbackRefresh:
			callbackRefresh()
	parentWin.Bind(wx.EVT_MENU, onPreview, id=idPreview)

	def onInvertToggle(evt):
		if not ensureGraphicReady(item, parent=parentWin, actionName="inwersji kolorów"):
			return
		newState = not getattr(item, "isInverted", False)
		msgKey = getattr(item, "msgKey", "") or getMessageKey()
		idx = getattr(item, "index", 0)
		if msgKey and idx:
			setImageInverted(msgKey, idx, newState)
		msg = _T("list_invert_toggled_on", idx=idx) if newState else _T("list_invert_toggled_off", idx=idx)
		ui.message(msg)
		if callbackRefresh:
			callbackRefresh()
	parentWin.Bind(wx.EVT_MENU, onInvertToggle, id=idInvertToggle)

	def onInstructions(evt):
		showBlockedImageInstructionsDialog(parentWin, item)
	parentWin.Bind(wx.EVT_MENU, onInstructions, id=idInstructions)

	def onRename(evt):
		if not isinstance(item, GraphicItem):
			ui.message("Zmień nazwę jest dostępne tylko dla grafik w treści wiadomości.")
			return
		curName = getattr(item, "customName", "")
		with wx.TextEntryDialog(
			parentWin,
			"Wprowadź własną, czytelną nazwę lub opis dla tej grafiki:\n(Zostanie ona zapamiętana przy każdym kolejnym otwarciu tej wiadomości)",
			caption="Zmień nazwę / nadaj własny opis grafice",
			value=curName,
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				val = dlg.GetValue().strip()
				if val:
					item.customName = val
					try:
						if hasattr(item.rawObject, "AlternativeText"):
							item.rawObject.AlternativeText = val
					except Exception:
						pass
					msgKey = getattr(item, "msgKey", "") or getMessageKey()
					setCustomNameForGraphic(msgKey, item.index, val)
					ui.message(f"Zapisano nową nazwę grafiki: {val}")
				else:
					item.customName = ""
					try:
						if hasattr(item.rawObject, "AlternativeText"):
							item.rawObject.AlternativeText = ""
					except Exception:
						pass
					msgKey = getattr(item, "msgKey", "") or getMessageKey()
					setCustomNameForGraphic(msgKey, item.index, "")
					ui.message("Przywrócono domyślną nazwę grafiki.")
				if callbackRefresh:
					callbackRefresh()
	parentWin.Bind(wx.EVT_MENU, onRename, id=idRename)

	def onClearName(evt):
		item.customName = ""
		try:
			if hasattr(item.rawObject, "AlternativeText"):
				item.rawObject.AlternativeText = ""
		except Exception:
			pass
		msgKey = getattr(item, "msgKey", "") or getMessageKey()
		setCustomNameForGraphic(msgKey, item.index, "")
		ui.message("Przywrócono domyślną nazwę grafiki.")
		if callbackRefresh:
			callbackRefresh()
	parentWin.Bind(wx.EVT_MENU, onClearName, id=idClearName)

	def onOpen(evt):
		safeOpenItem(item, parent=parentWin)
	parentWin.Bind(wx.EVT_MENU, onOpen, id=idOpen)

	def onSelect(evt):
		if isinstance(item, GraphicItem):
			try:
				item.rawObject.Select()
				ui.message(f"Zaznaczono grafikę {item.index}. Naciśnij klawisz Aplikacje (Shift+F10) dla menu czynności.")
			except Exception as e:
				ui.message(f"Nie udało się zaznaczyć grafiki: {e}")
	parentWin.Bind(wx.EVT_MENU, onSelect, id=idSelect)

	def onCopy(evt):
		if isinstance(item, GraphicItem):
			if not ensureGraphicReady(item, parent=parentWin, actionName="skopiowania do schowka"):
				return
			try:
				if item.isInline:
					item.rawObject.Range.Copy()
				else:
					item.rawObject.Copy()
				ui.message(f"Skopiowano grafikę {item.index} do schowka systemowego.")
			except Exception as e:
				showAccessibleMessageBox(f"Błąd kopiowania do schowka:\n{e}", title="Błąd schowka", parent=parentWin)
	parentWin.Bind(wx.EVT_MENU, onCopy, id=idCopy)

	def onSave(evt):
		if isAtt:
			with wx.FileDialog(
				parentWin,
				message="Zapisz załącznik jako plik",
				defaultFile=item.fileName,
				style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
			) as fileDialog:
				if fileDialog.ShowModal() == wx.ID_CANCEL:
					return
				destPath = fileDialog.GetPath()
			try:
				item.rawAttachment.SaveAsFile(destPath)
				ui.message(f"Zapisano załącznik jako: {os.path.basename(destPath)}")
			except Exception as e:
				showAccessibleMessageBox(f"Błąd podczas zapisywania załącznika:\n{e}", title="Błąd zapisu", parent=parentWin)
		else:
			if not ensureGraphicReady(item, parent=parentWin, actionName="zapisania do pliku"):
				return
			tempPng = extractGraphicToPngFile(item)
			if not tempPng or not os.path.isfile(tempPng):
				showAccessibleMessageBox("Nie udało się pobrać grafiki z edytora wiadomości.", title="Błąd zapisu", parent=parentWin)
				return

			useAi = config.conf[CONFIG_SECTION].get("aiNaming", False)
			apiKey = config.conf[CONFIG_SECTION].get("geminiApiKey", "")

			def _doSaveWithPrompt(aiCandidate=""):
				suggestedName = generateDescriptiveFilename(item, aiCandidate=aiCandidate)
				with wx.FileDialog(
					parentWin,
					message="Zapisz grafikę jako plik PNG",
					defaultFile=suggestedName,
					wildcard="Pliki obrazów PNG (*.png)|*.png",
					style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
				) as fileDialog:
					if fileDialog.ShowModal() == wx.ID_CANCEL:
						try:
							os.remove(tempPng)
						except Exception:
							pass
						return
					destPath = fileDialog.GetPath()

				try:
					import shutil
					shutil.copyfile(tempPng, destPath)
					ui.message(f"Zapisano grafikę jako: {os.path.basename(destPath)}")
				except Exception as e:
					showAccessibleMessageBox(f"Błąd podczas zapisywania pliku:\n{e}", title="Błąd zapisu", parent=parentWin)
				finally:
					try:
						os.remove(tempPng)
					except Exception:
						pass

			if useAi and apiKey and item.isScreenshotOrUnlabeled and not item.ocrText:
				def _aiNamingWorker():
					aiCand = ""
					try:
						aiCand = describeWithGemini(
							tempPng,
							apiKey,
							customPrompt="Podaj krótką nazwę pliku (maksymalnie 3-4 słowa po polsku, bez rozszerzenia .png), która najlepiej opisuje ten obraz. Zwróć wyłącznie samą nazwę."
						)
					except Exception as ex:
						log.debug(f"outlookImageHelper: błąd zapytania AI o nazwę pliku: {ex}")
					wx.CallAfter(_doSaveWithPrompt, aiCand)
				threading.Thread(target=_aiNamingWorker, daemon=True).start()
			else:
				_doSaveWithPrompt()
				showAccessibleMessageBox(f"Błąd podczas zapisywania pliku:\n{e}", title="Błąd zapisu", parent=parentWin)
			finally:
				try:
					os.remove(tempPng)
				except Exception:
					pass
	parentWin.Bind(wx.EVT_MENU, onSave, id=idSave)

	def onOpenList(evt):
		wx.CallAfter(openGraphicListDialog)
	parentWin.Bind(wx.EVT_MENU, onOpenList, id=idOpenList)

	def onSettings(evt):
		wx.CallAfter(gui.mainFrame.popupSettingsDialog, gui.settingsDialogs.NVDASettingsDialog, OutlookImageHelperSettingsPanel)
	parentWin.Bind(wx.EVT_MENU, onSettings, id=idSettings)

	def onHelp(evt):
		showContextHelp("list" if isinstance(parent, wx.Dialog) else "message", parent=parentWin)
	parentWin.Bind(wx.EVT_MENU, onHelp, id=idHelp)

	if not isinstance(parent, wx.Dialog):
		idOutlookMenu = wx.NewIdRef()
		menu.AppendSeparator()
		menu.Append(idOutlookMenu, "&Oryginalne menu podręczne programu Outlook\tShift+F10")

		def onOutlookMenu(evt):
			if GlobalPlugin.instance:
				GlobalPlugin.instance._bypassNextContextMenu = True
			import core
			def _sendApps():
				import winUser
				winUser.keybd_event(winUser.VK_SHIFT, 0, 0, 0)
				winUser.keybd_event(winUser.VK_F10, 0, 0, 0)
				winUser.keybd_event(winUser.VK_F10, 0, 2, 0)
				winUser.keybd_event(winUser.VK_SHIFT, 0, 2, 0)
			core.callLater(120, _sendApps)
		parentWin.Bind(wx.EVT_MENU, onOutlookMenu, id=idOutlookMenu)

	gui.mainFrame.prePopup()
	parentWin.PopupMenu(menu)
	gui.mainFrame.postPopup()
	menu.Destroy()


def _setRadioBoxFocus(radioBox, selection=0):
	"""
	Ustawia bezpośredni fokus klawiatury Win32 na konkretnym przycisku opcji (dziecku wxRadioBox),
	dzięki czemu NVDA natychmiast oznajmia np. '1. Grafiki i zrzuty ekranu w treści (22) przycisk opcji oznaczone Alt+ 1'
	zamiast 'wybór grup' / nagłówka ramki.
	"""
	try:
		import ctypes
		hwnd = radioBox.GetHandle()
		user32 = ctypes.windll.user32
		GW_CHILD = 5
		GW_HWNDNEXT = 2
		child = user32.GetWindow(hwnd, GW_CHILD)
		idx = 0
		target = child
		while child:
			if idx == selection:
				target = child
				break
			child = user32.GetWindow(child, GW_HWNDNEXT)
			idx += 1
		if target:
			user32.SetFocus(target)
	except Exception as ex:
		log.debug(f"outlookImageHelper: _setRadioBoxFocus error: {ex}")


def getOutlookApp():
	"""Zwraca obiekt Outlook.Application niezależnie od tego, jakie okno ma obecnie fokus w NVDA."""
	try:
		import appModuleHandler
		table = getattr(appModuleHandler, "runningTable", {})
		for mod in table.values():
			if getattr(mod, "appName", "").lower() == "outlook":
				nativeOm = getattr(mod, "nativeOm", None)
				if nativeOm:
					return nativeOm
	except Exception as ex:
		log.debug(f"outlookImageHelper.getOutlookApp runningTable: {ex}")

	try:
		focus = api.getFocusObject()
		if focus:
			appMod = getattr(focus, "appModule", None)
			if appMod and getattr(appMod, "appName", "").lower() == "outlook":
				nativeOm = getattr(appMod, "nativeOm", None)
				if nativeOm:
					return nativeOm
	except Exception:
		pass

	try:
		import comtypes.client
		ol = comtypes.client.GetActiveObject("Outlook.Application")
		if ol:
			return ol
	except Exception as ex:
		log.debug(f"outlookImageHelper.getOutlookApp comtypes: {ex}")

	return None


def downloadWebImage(item: GraphicItem) -> str | None:
	"""Wyświetla instrukcję odblokowania obrazów w programie Outlook."""
	showBlockedImageInstructionsDialog(None, item)
	return None


def getShapeTextContent(item: GraphicItem) -> str:
	"""
	Odczytuje tekst zawarty wewnątrz kształtu WordMail (TextFrame / TextRange),
	obiektu Shape, altText lub title.
	"""
	txt = ""
	raw = getattr(item, "rawObject", None)
	if raw:
		try:
			tf = getattr(raw, "TextFrame", None)
			if tf and getattr(tf, "HasText", 0):
				t = tf.TextRange.Text
				if t:
					txt = str(t).replace("\r\n", "\n").replace("\r", "\n").strip()
		except Exception as ex:
			log.debug(f"outlookImageHelper.getShapeTextContent TextFrame: {ex}")

		if not txt:
			try:
				tf2 = getattr(raw, "TextFrame2", None)
				if tf2 and getattr(tf2, "HasText", 0):
					t = tf2.TextRange.Text
					if t:
						txt = str(t).replace("\r\n", "\n").replace("\r", "\n").strip()
			except Exception as ex:
				log.debug(f"outlookImageHelper.getShapeTextContent TextFrame2: {ex}")

		if not txt:
			try:
				alt = getattr(raw, "AlternativeText", "")
				if alt:
					txt = str(alt).strip()
			except Exception:
				pass

	if not txt and getattr(item, "altText", ""):
		txt = item.altText.strip()
	if not txt and getattr(item, "title", ""):
		txt = item.title.strip()
	return txt


def ensureGraphicReady(item, parent=None, actionName: str = "wykonania czynności") -> bool:
	"""
	Upewnia się, że grafika jest gotowa do dalszego przetwarzania (podgląd, OCR, AI, inwersja).
	1. Sprawdza pamięć podręczną (obraz wyodrębniony z edytora WordMail).
	2. Jeśli grafika posiada obiekt WordMail (rawObject), podejmuje próbę wyodrębnienia jej
	   bezpośrednio z edytora WordMail programu Outlook.
	   Jeśli obraz zawiera rzeczywiste piksele (nie jest pustym placeholderem), zapisuje go w pamięci i zwraca True.
	3. Jeśli element to linia podziału lub separator (isDivider), wyświetla dostępne okno dialogowe i zwraca False.
	4. Jeśli obraz w edytorze WordMail okazał się pusty/zablokowany przez zabezpieczenia programu Outlook (isBlankOrSolid == True)
	   lub jest zewnętrzną grafiką ze strony WWW niepobraną przez program pocztowy:
	   wyświetla w pełni dostępne okno z jasną instrukcją krok po kroku (Ctrl+Shift+W -> Enter w programie Outlook).
	5. Jeśli to przycisk / kształt wektorowy programu Outlook (AutoShape, TextBox, Group):
	   - W przypadku OCR / odczytu tekstu: odczytuje tekst przycisku i prezentuje go w dostępnym oknie
	     dialogowym z pełną nawigacją kursorem NVDA, a także zapisuje w item.ocrText.
	   - W przypadku podglądu, inwersji, analizy AI, zapisu: wyświetla dostępne okno dialogowe ze szczegółami,
	     treścią elementu, linkiem i wskazówkami.
	6. Jeśli nie udało się wyodrębnić grafiki z innego powodu, wyświetla dostępne okno dialogowe z wyjaśnieniem przyczyn.
	"""
	if not isinstance(item, GraphicItem):
		return True

	# 1. Sprawdź, czy mamy już sprawny plik w pamięci podręcznej z rzeczywistymi pikselami
	if getattr(item, "_downloadedTempPath", None) and os.path.isfile(item._downloadedTempPath) and not getattr(item, "isBlankOrSolid", False):
		return True
	if getattr(item, "_extractedTempPath", None) and os.path.isfile(item._extractedTempPath) and not getattr(item, "isBlankOrSolid", False):
		return True

	# 2. Próba bezpośredniego wyodrębnienia lokalnego z edytora WordMail
	if getattr(item, "rawObject", None):
		try:
			tempPath = extractGraphicToPngFile(item)
			if tempPath and os.path.isfile(tempPath) and not getattr(item, "isBlankOrSolid", False):
				item._hasLocalRenderedPixels = True
				item._extractedTempPath = tempPath
				return True
		except Exception as ex:
			log.debug(f"outlookImageHelper.ensureGraphicReady: Próba wyodrębnienia lokalnego: {ex}")

	# 3. Sprawdź, czy to kształt wektorowy / przycisk WordMail (AutoShape, TextBox, Group)
	isVector = bool(getattr(item, "isVectorShape", False))
	if not isVector:
		try:
			rawObj = getattr(item, "rawObject", None)
			if rawObj is not None and getattr(rawObj, "Type", 0) in (1, 6, 17):
				isVector = True
		except Exception:
			pass

	# 4. Jeśli to standardowy separator lub linia podziału (bez kształtów wektorowych/przycisków)
	if getattr(item, "isDivider", False) and not isVector:
		title = "Linia podziału lub separator"
		msg = (
			f"Element {item.index}: {item.getDisplayLabel()}\n\n"
			f"Wybrany element to linia podziału, separator lub ozdobnik formatujący bez treści graficznej ani tekstu do {actionName}.\n\n"
			f"Wymiary separatora: {item.widthPx} × {item.heightPx} px."
		)
		showAccessibleMessageBox(msg, title=title, parent=parent)
		return False

	# 5. Jeśli obraz jest pusty / zablokowany przez zabezpieczenia programu Outlook
	isBlocked = (
		getattr(item, "isBlankOrSolid", False) or
		(getattr(item, "isExternalWebImage", False) and not getattr(item, "isDownloaded", False))
	)

	if isBlocked:
		showBlockedImageInstructionsDialog(parent, item, actionName=actionName)
		return False

	# 6. Obsługa przycisków i kształtów wektorowych WordMail
	if isVector:
		btnText = getShapeTextContent(item)
		realUrl = decodeSafeLinks(item.hyperlinkUrl) if getattr(item, "hyperlinkUrl", "") else ""
		actionLower = actionName.lower()

		if any(w in actionLower for w in ("ocr", "rozpoznan", "tekst")):
			if btnText:
				item.ocrText = btnText
				title = f"Tekst elementu — Element {item.index}"
				msgLines = [
					f"Element {item.index}: {item.getDisplayLabel()}",
					"",
					"Ten element jest przyciskiem lub kształtem wektorowym programu Outlook (AutoShape) z bezpośrednio zawartą treścią tekstową.",
					"",
					"Odczytany tekst elementu:",
					"--------------------------------------------------",
					btnText,
					"--------------------------------------------------",
				]
				if realUrl:
					msgLines.extend([
						"",
						"Powiązany adres WWW (link docelowy):",
						realUrl,
						"",
						"Wskazówka: Aby otworzyć ten link w przeglądarce, naciśnij Ctrl+Enter na liście elementów lub użyj przycisku 'Otwórz link'.",
					])
				if parent and hasattr(parent, "refreshListLabels"):
					try:
						parent.refreshListLabels()
					except Exception:
						pass
				showAccessibleMessageBox("\n".join(msgLines), title=title, parent=parent)
				return False
			else:
				title = f"Tekst elementu — Element {item.index}"
				msgLines = [
					f"Element {item.index}: {item.getDisplayLabel()}",
					"",
					"Ten element jest kształtem wektorowym programu Outlook (AutoShape), ale nie zawiera żadnego tekstu wewnątrz ani tekstu alternatywnego.",
					"Nie jest to również plik graficzny (raster), z którego można rozpoznać tekst za pomocą silnika OCR.",
				]
				if realUrl:
					msgLines.extend([
						"",
						"Powiązany adres WWW (link docelowy):",
						realUrl,
						"",
						"Wskazówka: Aby otworzyć ten link w przeglądarce, naciśnij Ctrl+Enter na liście elementów lub użyj przycisku 'Otwórz link'.",
					])
				showAccessibleMessageBox("\n".join(msgLines), title=title, parent=parent)
				return False

		elif any(w in actionLower for w in ("podgląd", "preview", "powiększ")):
			title = f"Podgląd niedostępny — Element {item.index}"
			msgLines = [
				f"Element {item.index}: {item.getDisplayLabel()}",
				"",
				"Podgląd graficzny jest niedostępny dla tego elementu.",
				"",
				"Wybrany element to przycisk lub kształt wektorowy programu WordMail/Outlook (AutoShape), a nie rastrowy plik graficzny (taki jak PNG, JPG czy GIF).",
				"Kształty tego typu są renderowane bezpośrednio przez mechanizm wektorowy programu Word i nie posiadają pliku ze zdjęciem, który można wyświetlić lub powiększyć w oknie podglądu.",
				"",
				"Szczegóły elementu:",
				f"- Wymiary: {item.widthPx} × {item.heightPx} px ({item.widthPt:.1f} × {item.heightPt:.1f} pt)",
				f"- Tekst na przycisku: {btnText or '(brak tekstu)'}",
			]
			if realUrl:
				msgLines.append(f"- Adres docelowy (link): {realUrl}")
			msgLines.extend([
				"",
				"Dostępne czynności dla tego elementu:",
				"- Naciśnij Enter lub wybierz 'Rozpoznaj tekst (OCR)', aby odczytać treść tekstową tego elementu w oknie tekstowym.",
			])
			if realUrl:
				msgLines.append("- Naciśnij Ctrl+Enter lub wybierz 'Otwórz link', aby przejść do docelowej strony internetowej.")
			showAccessibleMessageBox("\n".join(msgLines), title=title, parent=parent)
			return False

		elif any(w in actionLower for w in ("inwersj", "kolor")):
			title = f"Inwersja niedostępna — Element {item.index}"
			msg = (
				f"Element {item.index}: {item.getDisplayLabel()}\n\n"
				f"Inwersja kolorów jest niedostępna dla tego elementu.\n\n"
				f"Wybrany element to kształt wektorowy / przycisk programu Outlook, a nie plik graficzny. "
				f"Inwersja kolorów jest przeznaczona wyłącznie dla rastrowych zdjęć i grafik wklejonych w treść wiadomości."
			)
			showAccessibleMessageBox(msg, title=title, parent=parent)
			return False

		elif any(w in actionLower for w in ("ai", "rozmow", "sztuczn")):
			title = f"Analiza AI niedostępna — Element {item.index}"
			msg = (
				f"Element {item.index}: {item.getDisplayLabel()}\n\n"
				f"Wizualna analiza przez sztuczną inteligencję (AI) jest niedostępna dla tego elementu.\n\n"
				f"Wybrany element to przycisk lub kształt wektorowy programu Outlook (AutoShape) bez rastrowego pliku graficznego z pikselami. "
				f"Narzędzie Gemini Vision analizuje treść zdjęć i rysunków.\n\n"
				f"Tekst zawarty na tym przycisku: \"{btnText or '(brak)'}\".\n"
				f"Powiązany adres docelowy: {realUrl or '(brak linku)'}."
			)
			showAccessibleMessageBox(msg, title=title, parent=parent)
			return False

		elif any(w in actionLower for w in ("zapis", "plik")):
			title = f"Zapisywanie niedostępne — Element {item.index}"
			msg = (
				f"Element {item.index}: {item.getDisplayLabel()}\n\n"
				f"Nie można zapisać tego elementu jako pliku graficznego PNG.\n\n"
				f"Wybrany element to przycisk lub kształt wektorowy programu Outlook (AutoShape), "
				f"a program pocztowy nie udostępnia dla niego rastrowego strumienia danych graficznych."
			)
			showAccessibleMessageBox(msg, title=title, parent=parent)
			return False

		else:
			title = f"Element wektorowy — Element {item.index}"
			msg = (
				f"Element {item.index}: {item.getDisplayLabel()}\n\n"
				f"Czynność '{actionName}' nie może zostać wykonana dla tego elementu.\n\n"
				f"Wybrany element to przycisk lub kształt wektorowy programu WordMail/Outlook (AutoShape) bez rastrowego pliku graficznego.\n\n"
				f"Treść tekstowa elementu: \"{btnText or '(brak)'}\".\n"
				f"Link docelowy: {realUrl or '(brak)'}."
			)
			showAccessibleMessageBox(msg, title=title, parent=parent)
			return False

	# 6. Sprawdź ostatecznie, czy dysponujemy poprawnym plikiem graficznym z pikselami
	hasExtracted = (
		(getattr(item, "_downloadedTempPath", None) and os.path.isfile(item._downloadedTempPath) and not getattr(item, "isBlankOrSolid", False)) or
		(getattr(item, "_extractedTempPath", None) and os.path.isfile(item._extractedTempPath) and not getattr(item, "isBlankOrSolid", False))
	)
	if not hasExtracted:
		title = f"Błąd wyodrębnienia grafiki — Element {item.index}"
		msg = (
			f"Element {item.index}: {item.getDisplayLabel()}\n\n"
			f"Nie udało się wyodrębnić zawartości graficznej wybranego elementu z edytora wiadomości programu Outlook do celu {actionName}.\n\n"
			f"Możliwe przyczyny:\n"
			f"1. Element może być zablokowany przez zasady prywatności programu Outlook (naciśnij Ctrl+Shift+W w programie Outlook, aby pobrać całą zawartość).\n"
			f"2. Program Outlook nie udostępnił danych graficznych dla tego obiektu za pośrednictwem interfejsu WordMail COM.\n"
			f"3. Obiekt jest elementem dynamicznym lub skryptowym bez rastrowej reprezentacji w pamięci podręcznej."
		)
		showAccessibleMessageBox(msg, title=title, parent=parent)
		return False

	return True


# ==============================================================================
# OKNO DIALOGOWE: LISTA GRAFIK W WIADOMOŚCI
# ==============================================================================


class GraphicElementsDialog(wx.Dialog):
	"""
	Dostępne okno dialogowe prezentujące spis wszystkich grafik
	wykrytych w bieżącej wiadomości z podziałem na 3 odrębne kategorie:
	1. Rzeczywiste grafiki i zrzuty ekranu w treści
	2. Linie podziału i separatory
	3. Załączniki do wiadomości
	Z pełnym menu czynności, podglądem z powiększeniem, OCR, AI, Q&A oraz pomocą F1.
	"""

	def __init__(
		self,
		parent,
		graphics: list[GraphicItem],
		attachments: list[AttachmentItem] | None = None,
		totalDocLength: int = 0,
	):
		super().__init__(
			parent,
			title=_T("dlg_elements_title"),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.allGraphics = graphics or []
		# Rozdzielenie grafik na właściwe oraz linie podziału/separatory
		self.realGraphics = [g for g in self.allGraphics if not g.isDivider]
		self.dividerGraphics = [g for g in self.allGraphics if g.isDivider]
		self.attachments = attachments or []
		self.totalDocLength = totalDocLength

		# Domyślny wybór kategorii:
		# 0 = Rzeczywiste grafiki w treści (jeśli istnieją)
		# 1 = Załączniki do wiadomości (jeśli brak grafik, a są pliki)
		# 2 = Linie podziału, separatory i ozdobniki (jeśli są tylko separatory)
		self.hideDuplicates = True
		if self.realGraphics:
			self.currentMode = 0
		elif self.attachments:
			self.currentMode = 1
		elif self.dividerGraphics:
			self.currentMode = 2
		else:
			self.currentMode = 0

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# 0. Zakładki kategorii elementów (wx.Notebook: 1. Grafiki, 2. Załączniki, 3. Linie podziału i ozdobniki)
		self.notebook = wx.Notebook(self, style=wx.NB_TOP)

		# Strona 0: Grafiki w treści
		self.pageReal = wx.Panel(self.notebook)
		sizerReal = wx.BoxSizer(wx.VERTICAL)
		self.lblListReal = wx.StaticText(self.pageReal, label=self._getListLabelForMode(0))
		sizerReal.Add(self.lblListReal, 0, wx.LEFT | wx.RIGHT | wx.TOP, 4)
		self.listBoxReal = wx.ListBox(self.pageReal, choices=self._getActiveChoicesForMode(0), style=wx.LB_EXTENDED)
		sizerReal.Add(self.listBoxReal, 1, wx.EXPAND | wx.ALL, 4)
		self.pageReal.SetSizer(sizerReal)
		self.notebook.AddPage(self.pageReal, _T("cat_real_graphics", count=len(self.realGraphics)))

		# Strona 1: Załączniki
		self.pageAtt = wx.Panel(self.notebook)
		sizerAtt = wx.BoxSizer(wx.VERTICAL)
		self.lblListAtt = wx.StaticText(self.pageAtt, label=self._getListLabelForMode(1))
		sizerAtt.Add(self.lblListAtt, 0, wx.LEFT | wx.RIGHT | wx.TOP, 4)
		self.listBoxAtt = wx.ListBox(self.pageAtt, choices=self._getActiveChoicesForMode(1), style=wx.LB_EXTENDED)
		sizerAtt.Add(self.listBoxAtt, 1, wx.EXPAND | wx.ALL, 4)
		self.pageAtt.SetSizer(sizerAtt)
		self.notebook.AddPage(self.pageAtt, _T("cat_attachments", count=len(self.attachments)))

		# Strona 2: Linie podziału i ozdobniki
		self.pageDiv = wx.Panel(self.notebook)
		sizerDiv = wx.BoxSizer(wx.VERTICAL)
		self.lblListDiv = wx.StaticText(self.pageDiv, label=self._getListLabelForMode(2))
		sizerDiv.Add(self.lblListDiv, 0, wx.LEFT | wx.RIGHT | wx.TOP, 4)
		self.listBoxDiv = wx.ListBox(self.pageDiv, choices=self._getActiveChoicesForMode(2), style=wx.LB_EXTENDED)
		sizerDiv.Add(self.listBoxDiv, 1, wx.EXPAND | wx.ALL, 4)
		self.pageDiv.SetSizer(sizerDiv)
		self.notebook.AddPage(self.pageDiv, _T("cat_divider_graphics", count=len(self.dividerGraphics)))

		self.listBoxes = [self.listBoxReal, self.listBoxAtt, self.listBoxDiv]
		self.listLabels = [self.lblListReal, self.lblListAtt, self.lblListDiv]

		self.notebook.SetSelection(self.currentMode)
		mainSizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 8)

		# 2. Pole szczegółów wybranego elementu
		self.lblDetails = wx.StaticText(self, label=_T("lbl_details"))
		mainSizer.Add(self.lblDetails, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)

		self.detailsText = wx.TextCtrl(
			self,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_BESTWRAP,
			size=(-1, 100),
		)
		mainSizer.Add(self.detailsText, 0, wx.EXPAND | wx.ALL, 8)

		# Checkbox grupowania powtarzających się grafik (umieszczony pod oknem ze szczegółami)
		self.chkDeduplicate = wx.CheckBox(
			self,
			label="&Grupuj powtarzające się grafiki (Alt+D / klawisz A)",
		)
		self.chkDeduplicate.SetValue(True)
		self.chkDeduplicate.Bind(wx.EVT_CHECKBOX, self.onDeduplicateToggled)
		mainSizer.Add(self.chkDeduplicate, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

		# 3. Pasek przycisków głównych czynności
		btnSizer = wx.WrapSizer(wx.HORIZONTAL)

		self.btnActions = wx.Button(self, label=_T("btn_actions"))
		btnSizer.Add(self.btnActions, 0, wx.ALL, 4)

		self.btnPreview = wx.Button(self, label=_T("btn_preview"))
		btnSizer.Add(self.btnPreview, 0, wx.ALL, 4)

		self.btnInvert = wx.Button(self, label=_T("btn_invert"))
		btnSizer.Add(self.btnInvert, 0, wx.ALL, 4)

		self.btnOcr = wx.Button(self, label=_T("btn_ocr"))
		self.btnOcr.SetDefault()
		btnSizer.Add(self.btnOcr, 0, wx.ALL, 4)

		self.btnAi = wx.Button(self, label=_T("btn_ai_describe"))
		btnSizer.Add(self.btnAi, 0, wx.ALL, 4)

		self.btnChat = wx.Button(self, label=_T("btn_ai_chat"))
		btnSizer.Add(self.btnChat, 0, wx.ALL, 4)

		self.btnRename = wx.Button(self, label=_T("btn_rename"))
		btnSizer.Add(self.btnRename, 0, wx.ALL, 4)

		self.btnOpenLink = wx.Button(self, label=_T("btn_open_link"))
		btnSizer.Add(self.btnOpenLink, 0, wx.ALL, 4)

		self.btnSave = wx.Button(self, label=_T("btn_save"))
		btnSizer.Add(self.btnSave, 0, wx.ALL, 4)

		self.btnHelp = wx.Button(self, label=_T("btn_help"))
		btnSizer.Add(self.btnHelp, 0, wx.ALL, 4)

		self.btnClose = wx.Button(self, wx.ID_CANCEL, label=_T("btn_close"))
		btnSizer.Add(self.btnClose, 0, wx.ALL, 4)

		mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.ALL, 8)
		self.SetSizer(mainSizer)
		self.Fit()
		self.SetSize((780, 580))
		self.CenterOnScreen()

		# Zdarzenia
		self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.onPageChanged)
		self.notebook.Bind(wx.EVT_KEY_DOWN, self.onNotebookKeyDown)

		for lb in self.listBoxes:
			lb.Bind(wx.EVT_LISTBOX, self.onSelectionChanged)
			lb.Bind(wx.EVT_LISTBOX_DCLICK, self.onOcrClicked)
			lb.Bind(wx.EVT_CONTEXT_MENU, self.onContextMenuEvent)
			lb.Bind(wx.EVT_KEY_DOWN, self.onListKeyDown)

		self.btnActions.Bind(wx.EVT_BUTTON, lambda evt: self.showContextMenu())
		self.btnPreview.Bind(wx.EVT_BUTTON, lambda evt: self.onPreviewClicked(self.getSelectedItem()))
		self.btnInvert.Bind(wx.EVT_BUTTON, lambda evt: self.onInvertClicked())
		self.btnOcr.Bind(wx.EVT_BUTTON, self.onOcrClicked)
		self.btnAi.Bind(wx.EVT_BUTTON, self.onAiPromptClicked)
		self.btnChat.Bind(wx.EVT_BUTTON, self.onAiChatClicked)
		self.btnRename.Bind(wx.EVT_BUTTON, self.onRenameItemClicked)
		self.btnOpenLink.Bind(wx.EVT_BUTTON, self.onOpenLinkButtonClicked)
		self.btnSave.Bind(wx.EVT_BUTTON, self.onSaveClicked)
		self.btnHelp.Bind(wx.EVT_BUTTON, lambda evt: showContextHelp("list", parent=self))
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		self.Bind(wx.EVT_SHOW, self.onShowDialog)

		for mode, lb in enumerate(self.listBoxes):
			if self._getItemsForMode(mode):
				lb.SetSelection(0)
		self.updateDetails()
		playItemTypeSound(self.getSelectedItem())

	@property
	def listBox(self) -> wx.ListBox:
		return self.listBoxes[self.currentMode]

	@property
	def lblList(self) -> wx.StaticText:
		return self.listLabels[self.currentMode]

	def onShowDialog(self, evt):
		if evt.IsShown():
			wx.CallAfter(lambda: self.notebook.SetFocus())
		evt.Skip()

	def onNotebookKeyDown(self, event):
		key = event.GetKeyCode()
		if key == wx.WXK_DOWN:
			self._focusActiveListBox()
			return
		elif key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
			self._focusActiveListBox()
			return
		event.Skip()

	def ensureGraphicReady(self, item, actionName: str = "wykonania czynności") -> bool:
		return ensureGraphicReady(item, parent=self, actionName=actionName)

	def onDeduplicateToggled(self, event=None):
		"""Przełącza ukrywanie lub pokazywanie powtarzających się grafik."""
		self.hideDuplicates = self.chkDeduplicate.GetValue()
		for mode in (0, 2):
			self.listLabels[mode].SetLabel(self._getListLabelForMode(mode))
			self.listBoxes[mode].Set(self._getChoicesForMode(mode))
			if self._getItemsForMode(mode):
				self.listBoxes[mode].SetSelection(0)
			items = self._getItemsForMode(mode)
			total = len(self.realGraphics) if mode == 0 else len(self.dividerGraphics)
			if self.hideDuplicates and len(items) < total:
				cntStr = f"{len(items)} z {total}"
			else:
				cntStr = str(total)
			tag = "cat_real_graphics" if mode == 0 else "cat_divider_graphics"
			self.notebook.SetPageText(mode, _T(tag, count=cntStr))
		self.updateDetails()

	def _getDeduplicated(self, items: list[GraphicItem | AttachmentItem]) -> list[GraphicItem | AttachmentItem]:
		"""Grupuje identyczne grafiki w jeden reprezentatywny element z licznikiem powtórzeń."""
		if not items:
			return []
		seen = {}
		result = []
		for item in items:
			if isinstance(item, GraphicItem):
				hasCid = bool(item.cid and item.cid.strip())
				isDivider = bool(getattr(item, "isDivider", False))
				hasIdenticalSourceUrl = bool(getattr(item, "sourceUrl", "") and item.sourceUrl.strip())
				if hasCid:
					key = ("cid", item.cid.strip().lower(), round(item.widthPt, 1), round(item.heightPt, 1))
				elif hasIdenticalSourceUrl:
					key = ("sourceUrl", item.sourceUrl.strip().lower(), round(item.widthPt, 1), round(item.heightPt, 1))
				elif isDivider:
					key = ("divider", round(item.widthPt, 1), round(item.heightPt, 1))
				else:
					# Brak silnego identyfikatora tożsamości źródła - nie uznawaj różnych grafik za identyczne
					key = ("unique", id(item), item.index)

				if key in seen:
					seen[key]._repeatCount = getattr(seen[key], "_repeatCount", 1) + 1
				else:
					item._repeatCount = 1
					seen[key] = item
					result.append(item)
			else:
				result.append(item)
		return result

	def switchCategory(self, mode: int, focusList: bool = True):
		"""Przełącza aktywną zakładkę (0=Grafiki, 1=Załączniki, 2=Linie podziału) i opcjonalnie przenosi fokus na listę."""
		if mode < 0 or mode >= self.notebook.GetPageCount():
			return

		self.currentMode = mode
		if self.notebook.GetSelection() != mode:
			self.notebook.SetSelection(mode)

		# Ogłoś nagłówek wybranej zakładki (np. "1. Grafiki i zrzuty ekranu w treści (2)")
		tabText = self.notebook.GetPageText(mode).replace("&", "")
		ui.message(tabText)

		self.updateDetails()
		playItemTypeSound(self.getSelectedItem())

		if focusList:
			wx.CallAfter(self._focusActiveListBox)

	def _focusActiveListBox(self):
		"""Ustawia bezpośredni fokus na liście aktywnej zakładki i zaznacza pierwszy element (jeśli lista nie jest pusta)."""
		try:
			lb = self.listBox
			items = self._getActiveItems()
			if items:
				lb.SetSelection(0)
			lb.SetFocus()
			self.updateDetails()
			if not items:
				catNames = {
					0: "Brak grafik w tej wiadomości.",
					1: "Brak załączników w tej wiadomości.",
					2: "Brak linii podziału ani elementów ozdobnych w tej wiadomości.",
				}
				ui.message(catNames.get(self.currentMode, "Brak elementów w wybranej kategorii."))
		except Exception as ex:
			log.debug(f"outlookImageHelper._focusActiveListBox: {ex}")

	def onPageChanged(self, event):
		"""Obsługuje przełączenie zakładki (np. klawiszami strzałek w lewo/prawo lub Ctrl+Tab)."""
		newMode = event.GetSelection()
		if newMode != -1 and newMode != self.currentMode:
			self.currentMode = newMode
			self.updateDetails()
			playItemTypeSound(self.getSelectedItem())
		event.Skip()

	def _getItemsForMode(self, mode: int) -> list[GraphicItem | AttachmentItem]:
		if mode == 0:
			items = self.realGraphics
		elif mode == 1:
			items = self.attachments
		else:
			items = self.dividerGraphics

		if getattr(self, "hideDuplicates", True) and mode in (0, 2):
			return self._getDeduplicated(items)
		else:
			for it in items:
				if hasattr(it, "_repeatCount"):
					it._repeatCount = 1
			return items

	def _getChoicesForMode(self, mode: int) -> list[str]:
		items = self._getItemsForMode(mode)
		return [it.getDisplayLabel(self.totalDocLength) for it in items]

	def _getActiveChoicesForMode(self, mode: int) -> list[str]:
		return self._getChoicesForMode(mode)

	def _getListLabelForMode(self, mode: int) -> str:
		if mode == 0:
			if getattr(self, "hideDuplicates", True):
				act = self._getItemsForMode(0)
				if len(act) < len(self.realGraphics):
					return f"Lista grafik w treści ({len(act)} unikalnych z {len(self.realGraphics)}):"
			return _T("lbl_real_list")
		elif mode == 1:
			return _T("lbl_att_list")
		else:
			if getattr(self, "hideDuplicates", True):
				act = self._getItemsForMode(2)
				if len(act) < len(self.dividerGraphics):
					return f"Linie podziału, separatory i ozdobniki ({len(act)} unikalnych z {len(self.dividerGraphics)}):"
			return _T("lbl_divider_list")

	def _getActiveItems(self) -> list[GraphicItem | AttachmentItem]:
		return self._getItemsForMode(self.currentMode)

	def _getActiveChoices(self) -> list[str]:
		return self._getChoicesForMode(self.currentMode)

	def _getListLabel(self) -> str:
		return self._getListLabelForMode(self.currentMode)

	def onCharHook(self, event):
		key = event.GetKeyCode()
		isAltGr = (event.ControlDown() and event.AltDown()) or bool(winUser.getKeyState(winUser.VK_RMENU) & 0x8000)
		if isAltGr:
			event.Skip()
			return

		if event.AltDown() and not event.ControlDown() and not event.ShiftDown():
			targetIdx = None
			if key in (ord('1'), wx.WXK_NUMPAD1):
				targetIdx = 0
			elif key in (ord('2'), wx.WXK_NUMPAD2):
				targetIdx = 1
			elif key in (ord('3'), wx.WXK_NUMPAD3):
				targetIdx = 2

			if targetIdx is not None and 0 <= targetIdx < self.notebook.GetPageCount():
				self.switchCategory(targetIdx, focusList=True)
				return
			elif key in (ord('D'), ord('d')):
				self.hideDuplicates = not self.hideDuplicates
				self.chkDeduplicate.SetValue(self.hideDuplicates)
				self.onDeduplicateToggled()
				msg = "Grupowanie powtarzających się grafik włączone" if self.hideDuplicates else "Wyświetlanie wszystkich powtórzeń (bez grupowania)"
				ui.message(msg)
				return

		if key == wx.WXK_ESCAPE:
			self.Close()
		elif key == wx.WXK_F1:
			showContextHelp("list", parent=self)
		elif key == ord('C') and event.AltDown():
			self.showContextMenu()
		else:
			event.Skip()


	def onListKeyDown(self, event):
		key = event.GetKeyCode()
		if key == wx.WXK_WINDOWS_MENU or (key == wx.WXK_F10 and event.ShiftDown()):
			self.showContextMenu()
		elif key == wx.WXK_F1:
			showContextHelp("list", parent=self)
		elif key in (ord('P'), ord('p')) and not event.ControlDown() and not event.AltDown():
			item = self.getSelectedItem()
			if item:
				self.onPreviewClicked(item)
		elif key in (ord('I'), ord('i')) and not event.ControlDown() and not event.AltDown():
			self.onInvertClicked()
		elif key == wx.WXK_F2:
			self.onRenameItemClicked(None)
		elif key in (ord('A'), ord('a')) and not event.ControlDown() and not event.AltDown():
			self.hideDuplicates = not self.hideDuplicates
			self.chkDeduplicate.SetValue(self.hideDuplicates)
			self.onDeduplicateToggled()
			msg = "Grupowanie powtarzających się grafik włączone" if self.hideDuplicates else "Wyświetlanie wszystkich powtórzeń (bez grupowania)"
			ui.message(msg)
		elif key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
			if event.ControlDown():
				item = self.getSelectedItem()
				if item:
					self.onOpenLinkClicked(item)
					return
			self.onOcrClicked(None)
		else:
			event.Skip()

	def onContextMenuEvent(self, event):
		self.showContextMenu()

	def onInvertClicked(self):
		item = self.getSelectedItem()
		if not item:
			return
		if not self.ensureGraphicReady(item, actionName="inwersji kolorów"):
			return
		self.toggleInvertItem(item)

	def toggleInvertItem(self, item: GraphicItem | AttachmentItem):
		"""Przełącza tryb inwersji kolorów dla wybranego elementu i trwale go zapamiętuje."""
		if isinstance(item, AttachmentItem) and not getattr(item, "isImage", False):
			showAccessibleMessageBox(_T("invert_msg_not_graphic"), title="Inwersja niedostępna", parent=self)
			return
		newInv = not getattr(item, "isInverted", False)
		item.isInverted = newInv
		msgKey = getattr(item, "msgKey", "") or getMessageKey()
		idx = getattr(item, "index", 0)
		if msgKey and idx:
			setImageInverted(msgKey, idx, newInv)
		self.refreshListLabels()
		if newInv:
			ui.message(_T("list_invert_toggled_on", idx=idx))
		else:
			ui.message(_T("list_invert_toggled_off", idx=idx))

	def onPreviewClicked(self, item: GraphicItem | AttachmentItem | None):
		"""Otwiera okno powiększonego podglądu dla wybranej grafiki lub załączonego obrazu."""
		if not item:
			return
		if not self.ensureGraphicReady(item, actionName="wyświetlenia powiększonego podglądu"):
			return
		if isinstance(item, AttachmentItem):
			if not getattr(item, "isImage", False):
				showAccessibleMessageBox(_T("preview_only_images"), title="Podgląd niedostępny", parent=self)
				return
			tempPath = saveAttachmentToTempFile(item)
			if not tempPath or not os.path.isfile(tempPath):
				showAccessibleMessageBox(_T("err_extract_att"), title="Błąd załącznika", parent=self)
				return
		else:
			tempPath = getattr(item, "_extractedTempPath", None) or getattr(item, "_downloadedTempPath", None)
			if not tempPath or not os.path.isfile(tempPath):
				ui.message(_T("msg_extracting_preview"))
				tempPath = extractGraphicToPngFile(item)
				if not tempPath or not os.path.isfile(tempPath):
					showAccessibleMessageBox(_T("err_extract_img"), title="Błąd podglądu", parent=self)
					return

		gui.mainFrame.prePopup()
		try:
			d = ImagePreviewDialog(self, item, tempPath)
			d.ShowModal()
			d.Destroy()
		finally:
			gui.mainFrame.postPopup()
		self.refreshListLabels()

	def getSelectedItem(self) -> GraphicItem | AttachmentItem | None:
		items = self._getActiveItems()
		selections = self.listBox.GetSelections()
		if selections:
			idx = selections[0]
			if 0 <= idx < len(items):
				return items[idx]
		return None

	def getSelectedItemsList(self) -> list[GraphicItem | AttachmentItem]:
		items = self._getActiveItems()
		selections = self.listBox.GetSelections()
		return [items[i] for i in selections if 0 <= i < len(items)]

	# Aliasy dla kompatybilności
	getSelectedGraphic = getSelectedItem
	getSelectedGraphicsList = getSelectedItemsList

	def onSelectionChanged(self, event):
		self.updateDetails()
		playItemTypeSound(self.getSelectedItem())

	def updateDetails(self):
		selected = self.getSelectedItemsList()
		if len(selected) == 1:
			item = selected[0]
			self.detailsText.SetValue(item.getDetailedInfo())
			self.btnChat.Enable()

			if self.currentMode in (0, 2):
				# Grafiki w treści (właściwe lub separatory)
				self.btnPreview.Enable()
				self.btnInvert.Enable()
				self.btnRename.Enable()
				self.btnSave.SetLabel(_T("btn_save_one_img"))
				self.btnOcr.SetLabel(_T("btn_ocr"))
				self.btnAi.SetLabel(_T("btn_ai_describe"))
				self.btnChat.SetLabel(_T("btn_ai_chat"))
				if getattr(item, "hyperlinkUrl", ""):
					svc = getattr(item, "linkServiceName", "")
					dom = svc if svc else getCleanDomain(item.hyperlinkUrl)
					self.btnOpenLink.Enable()
					self.btnOpenLink.SetLabel(f"&Otwórz link ({dom})..." if dom else _T("btn_open_link"))
				else:
					self.btnOpenLink.Disable()
					self.btnOpenLink.SetLabel(_T("btn_open_link"))
			else:
				# Załączniki do wiadomości
				self.btnRename.Disable()
				self.btnSave.SetLabel(_T("btn_save_one_att"))
				self.btnOpenLink.Enable()
				self.btnOpenLink.SetLabel(_T("btn_open_file"))
				if getattr(item, "isImage", False):
					self.btnPreview.Enable()
					self.btnInvert.Enable()
				else:
					self.btnPreview.Disable()
					self.btnInvert.Disable()

				if getattr(item, "isPdf", False):
					self.btnOcr.SetLabel(_T("btn_ocr_pdf"))
					self.btnAi.SetLabel("&Analiza PDF przez AI...")
					self.btnChat.SetLabel("&Rozmowa z AI o PDF...")
				elif getattr(item, "isDocx", False) or getattr(item, "isText", False):
					self.btnOcr.SetLabel(_T("btn_read_doc"))
					self.btnAi.SetLabel("&Streszczenie przez AI...")
					self.btnChat.SetLabel("&Rozmowa z AI o dokumencie...")
				else:
					self.btnOcr.SetLabel(_T("btn_ocr"))
					self.btnAi.SetLabel(_T("btn_ai_describe"))
					self.btnChat.SetLabel(_T("btn_ai_chat"))

		elif len(selected) > 1:
			cnt = len(selected)
			self.detailsText.SetValue(f"Zaznaczono {cnt} elementów. Naciśnij klawisz Aplikacje lub 'Czynności' dla operacji zbiorczych.")
			self.btnSave.SetLabel(f"&Zapisz zaznaczone ({cnt})...")
			self.btnRename.Disable()
			self.btnPreview.Disable()
			self.btnInvert.Disable()
			self.btnChat.Disable()
			self.btnOpenLink.Disable()
		else:
			self.detailsText.SetValue("")
			self.btnSave.SetLabel(_T("btn_save"))
			self.btnRename.Disable()
			self.btnPreview.Disable()
			self.btnInvert.Disable()
			self.btnChat.Disable()
			self.btnOpenLink.Disable()

	def refreshListLabels(self):
		"""Aktualizuje etykiety w liście po zakończeniu OCR lub AI."""
		try:
			selections = self.listBox.GetSelections()
			choices = self._getActiveChoices()
			for idx, text in enumerate(choices):
				if idx < self.listBox.GetCount():
					self.listBox.SetString(idx, text)
			for idx in selections:
				if idx < self.listBox.GetCount():
					self.listBox.SetSelection(idx)
			self.updateDetails()
		except Exception as e:
			log.debug(f"outlookImageHelper: refreshListLabels error: {e}")

	# ==========================================================================
	# MENU CZYNNOŚCI (KONTEKSTOWE)
	# ==========================================================================

	def showContextMenu(self):
		selected = self.getSelectedItemsList()
		if not selected:
			ui.message("Nie zaznaczono żadnego elementu.")
			return

		if len(selected) == 1:
			showGraphicActionsMenu(selected[0], parent=self, callbackRefresh=self.refreshListLabels)
			return

		menu = wx.Menu()
		count = len(selected)
		idBatchOcr = wx.NewIdRef()
		idBatchAi = wx.NewIdRef()
		idBatchSave = wx.NewIdRef()
		idSaveAll = wx.NewIdRef()
		idDeselect = wx.NewIdRef()

		if self.currentMode in (0, 2):
			menu.Append(idBatchOcr, f"&Rozpoznaj tekst (OCR) dla zaznaczonych ({count})")
			menu.Append(idBatchAi, f"&Opisz przez AI dla zaznaczonych ({count})")
			menu.Append(idBatchSave, f"&Zapisz zaznaczone grafiki do folderu ({count})...")
			activeCount = len(self._getActiveItems())
			menu.Append(idSaveAll, f"Zapisz &wszystkie z tej kategorii do folderu ({activeCount})...")
		else:
			menu.Append(idBatchAi, f"&Analiza AI dla zaznaczonych załączników ({count})")
			menu.Append(idBatchSave, f"&Zapisz zaznaczone załączniki do folderu ({count})...")
			menu.Append(idSaveAll, f"Zapisz &wszystkie załączniki do folderu ({len(self.attachments)})...")

		menu.AppendSeparator()
		menu.Append(idDeselect, "&Odznacz wszystkie")

		self.Bind(wx.EVT_MENU, lambda evt: self.onBatchOcrClicked(selected), id=idBatchOcr)
		self.Bind(wx.EVT_MENU, lambda evt: self.onBatchAiClicked(selected), id=idBatchAi)
		self.Bind(wx.EVT_MENU, lambda evt: self.onSaveClicked(None), id=idBatchSave)
		self.Bind(wx.EVT_MENU, self.onSaveAllClicked, id=idSaveAll)
		self.Bind(wx.EVT_MENU, self.onDeselectAllClicked, id=idDeselect)

		self.PopupMenu(menu)
		menu.Destroy()

	# ==========================================================================
	# OBSŁUGA AKCJI
	# ==========================================================================

	def onRenameItemClicked(self, event=None):
		"""Umożliwia użytkownikowi nadanie własnej czytelnej nazwy dla grafiki (F2)."""
		item = self.getSelectedItem()
		if not item or not isinstance(item, GraphicItem):
			ui.message("Zmień nazwę jest dostępne tylko dla grafik w treści wiadomości.")
			return
		curName = getattr(item, "customName", "")
		with wx.TextEntryDialog(
			self,
			"Wprowadź własną, czytelną nazwę lub opis dla tej grafiki:\n(Zostanie ona zapamiętana przy każdym kolejnym otwarciu tej wiadomości)",
			caption="Zmień nazwę / nadaj własny opis grafice",
			value=curName,
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				val = dlg.GetValue().strip()
				if val:
					item.customName = val
					try:
						if hasattr(item.rawObject, "AlternativeText"):
							item.rawObject.AlternativeText = val
					except Exception:
						pass
					msgKey = getattr(item, "msgKey", "") or getMessageKey()
					setCustomNameForGraphic(msgKey, item.index, val)
					self.refreshListLabels()
					ui.message(f"Zapisano nową nazwę grafiki: {val}")
				else:
					self.onClearCustomNameClicked(item)

	def onClearCustomNameClicked(self, item: GraphicItem):
		"""Usuwa własną nazwę grafiki i przywraca opis domyślny."""
		item.customName = ""
		try:
			if hasattr(item.rawObject, "AlternativeText"):
				item.rawObject.AlternativeText = ""
		except Exception:
			pass
		msgKey = getattr(item, "msgKey", "") or getMessageKey()
		setCustomNameForGraphic(msgKey, item.index, "")
		self.refreshListLabels()
		ui.message("Przywrócono domyślną nazwę grafiki.")

	def onOpenLinkButtonClicked(self, event):
		item = self.getSelectedItem()
		if item:
			self.onOpenLinkClicked(item)

	def onOpenLinkClicked(self, item: GraphicItem | AttachmentItem):
		safeOpenItem(item, parent=self)

	def onOcrClicked(self, event):
		selected = self.getSelectedItemsList()
		if not selected:
			return
		if len(selected) > 1:
			self.onBatchOcrClicked(selected)
		else:
			if not self.ensureGraphicReady(selected[0], actionName="rozpoznania tekstu (OCR)"):
				return
			GlobalPlugin.instance.triggerOcrForItem(selected[0], parentDialog=self)

	def onAiQuickClicked(self, event):
		item = self.getSelectedItem()
		if not item:
			return
		if not self.ensureGraphicReady(item, actionName="analizy i opisu przez AI"):
			return
		GlobalPlugin.instance.triggerAiDescriptionForItem(item, parentDialog=self, customPrompt="")

	def onAiPromptClicked(self, event):
		item = self.getSelectedItem()
		if not item:
			return
		if not self.ensureGraphicReady(item, actionName="analizy i opisu przez AI"):
			return
		def doPrompt():
			with AiPromptSelectionDialog(self, item) as dlg:
				if dlg.ShowModal() == wx.ID_OK:
					prompt = dlg.selectedPrompt
					GlobalPlugin.instance.triggerAiDescriptionForItem(item, parentDialog=self, customPrompt=prompt)
		if not getActiveAiApiKey():
			checkAndPromptAiApiKey(parentWin=self, onKeyReadyCallback=doPrompt)
			return
		doPrompt()

	def onAiChatClicked(self, event):
		item = self.getSelectedItem()
		if not item:
			return
		if not self.ensureGraphicReady(item, actionName="rozmowy z AI"):
			return
		if isinstance(item, AttachmentItem) and item.isPdf:
			if not ensurePdfUnlocked(item, parentWin=self):
				return
		def doChat():
			gui.mainFrame.prePopup()
			try:
				d = AiConversationDialog(self, item)
				d.ShowModal()
				d.Destroy()
			finally:
				gui.mainFrame.postPopup()
			self.refreshListLabels()
		if not getActiveAiApiKey():
			checkAndPromptAiApiKey(parentWin=self, onKeyReadyCallback=doChat)
			return
		doChat()

	def onBatchOcrClicked(self, items: list[GraphicItem | AttachmentItem]):
		GlobalPlugin.instance.triggerBatchOcr(items, parentDialog=self)

	def onBatchAiClicked(self, items: list[GraphicItem | AttachmentItem]):
		def doBatch():
			GlobalPlugin.instance.triggerBatchAi(items, parentDialog=self)
		if not getActiveAiApiKey():
			checkAndPromptAiApiKey(parentWin=self, onKeyReadyCallback=doBatch)
			return
		doBatch()

	def onDeselectAllClicked(self, event):
		items = self._getActiveItems()
		for i in range(len(items)):
			self.listBox.Deselect(i)
		if items:
			self.listBox.SetSelection(0)
		self.updateDetails()

	def onSelectClicked(self, event):
		item = self.getSelectedItem()
		if not item or not isinstance(item, GraphicItem):
			return
		self.Close()
		try:
			item.rawObject.Select()
			ui.message(f"Zaznaczono grafikę {item.index}. Naciśnij klawisz Aplikacje (Shift+F10) dla menu obrazu.")
		except Exception as e:
			ui.message(f"Nie udało się zaznaczyć grafiki: {e}")

	def onCopyClicked(self, event):
		item = self.getSelectedItem()
		if not item or not isinstance(item, GraphicItem):
			return
		if not self.ensureGraphicReady(item, actionName="skopiowania do schowka"):
			return
		try:
			if item.isInline:
				item.rawObject.Range.Copy()
			else:
				item.rawObject.Copy()
			ui.message(f"Skopiowano grafikę {item.index} do schowka systemowego.")
		except Exception as e:
			showAccessibleMessageBox(f"Błąd kopiowania do schowka:\n{e}", title="Błąd schowka", parent=self)

	def onSaveClicked(self, event):
		selected = self.getSelectedItemsList()
		if not selected:
			showAccessibleMessageBox("Nie wybrano żadnego elementu do zapisania.", title="Zapisywanie", parent=self)
			return

		if self.currentMode in (0, 2):
			# Zapisywanie grafik
			if len(selected) > 1:
				self._batchSaveGraphics(selected, f"Zapisano {len(selected)} wybranych grafik")
				return

			item = selected[0]
			if not self.ensureGraphicReady(item, actionName="zapisania do pliku"):
				return

			tempPng = extractGraphicToPngFile(item)
			if not tempPng or not os.path.isfile(tempPng):
				showAccessibleMessageBox("Nie udało się pobrać grafiki z edytora wiadomości.", title="Błąd zapisu", parent=self)
				return

			useAi = config.conf[CONFIG_SECTION].get("aiNaming", False)
			apiKey = config.conf[CONFIG_SECTION].get("geminiApiKey", "")

			def _doSaveWithPrompt(aiCandidate=""):
				suggestedName = generateDescriptiveFilename(item, aiCandidate=aiCandidate)
				with wx.FileDialog(
					self,
					message="Zapisz grafikę jako plik PNG",
					defaultFile=suggestedName,
					wildcard="Pliki obrazów PNG (*.png)|*.png",
					style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
				) as fileDialog:
					if fileDialog.ShowModal() == wx.ID_CANCEL:
						try:
							os.remove(tempPng)
						except Exception:
							pass
						return
					destPath = fileDialog.GetPath()

				try:
					import shutil
					shutil.copyfile(tempPng, destPath)
					ui.message(f"Zapisano grafikę jako: {os.path.basename(destPath)}")
				except Exception as e:
					showAccessibleMessageBox(f"Błąd podczas zapisywania pliku:\n{e}", title="Błąd zapisu", parent=self)
				finally:
					try:
						os.remove(tempPng)
					except Exception:
						pass

			if useAi and apiKey and item.isScreenshotOrUnlabeled and not item.ocrText:
				def _aiNamingWorker():
					aiCand = ""
					try:
						aiCand = describeWithGemini(
							tempPng,
							apiKey,
							customPrompt="Podaj krótką nazwę pliku (maksymalnie 3-4 słowa po polsku, bez rozszerzenia .png), która najlepiej opisuje ten obraz. Zwróć wyłącznie samą nazwę."
						)
					except Exception as ex:
						log.debug(f"outlookImageHelper: błąd zapytania AI o nazwę pliku: {ex}")
					wx.CallAfter(_doSaveWithPrompt, aiCand)
				threading.Thread(target=_aiNamingWorker, daemon=True).start()
			else:
				_doSaveWithPrompt()
		else:
			# Zapisywanie załączników
			if len(selected) > 1:
				self._batchSaveAttachments(selected, f"Zapisano {len(selected)} wybranych załączników")
				return

			attItem = selected[0]
			with wx.FileDialog(
				self,
				message=f"Zapisz załącznik: {attItem.fileName}",
				defaultFile=attItem.fileName,
				wildcard="Wszystkie pliki (*.*)|*.*",
				style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
			) as fileDialog:
				if fileDialog.ShowModal() == wx.ID_CANCEL:
					return
				destPath = fileDialog.GetPath()

			try:
				attItem.rawAttachment.SaveAsFile(destPath)
				ui.message(f"Zapisano załącznik jako: {os.path.basename(destPath)}")
			except Exception as e:
				showAccessibleMessageBox(f"Błąd podczas zapisywania załącznika:\n{e}", title="Błąd zapisu", parent=self)

	def onSaveAllClicked(self, event):
		if self.currentMode in (0, 2):
			targetGraphics = self.realGraphics if self.currentMode == 0 else self.dividerGraphics
			if not targetGraphics:
				ui.message("Brak grafik do zapisania.")
				return
			self._batchSaveGraphics(targetGraphics, f"Pomyślnie zapisano wszystkie {len(targetGraphics)} grafik")
		else:
			if not self.attachments:
				ui.message("Brak załączników do zapisania.")
				return
			self._batchSaveAttachments(self.attachments, f"Pomyślnie zapisano wszystkie {len(self.attachments)} załączników")

	def _batchSaveGraphics(self, items: list[GraphicItem], successMessagePrefix: str):
		"""Seryjny zapis listy grafik do wybranego folderu z asynchronicznym nazywaniem AI bez blokowania NVDA."""
		with wx.DirDialog(
			self,
			message="Wybierz folder docelowy dla zapisywanych grafik",
			style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST,
		) as dirDialog:
			if dirDialog.ShowModal() == wx.ID_CANCEL:
				return
			targetDir = dirDialog.GetPath()

		if not os.path.isdir(targetDir):
			ui.message("Wybrany folder nie istnieje.")
			return

		ui.message(f"Trwa przygotowywanie {len(items)} grafik do zapisu...")

		useAi = config.conf[CONFIG_SECTION].get("aiNaming", False)
		apiKey = config.conf[CONFIG_SECTION].get("geminiApiKey", "")

		# Ekstrakcja grafik na wątku STA przed uruchomieniem wątku w tle
		prepared = []
		for item in items:
			try:
				p = extractGraphicToPngFile(item)
				prepared.append((item, p))
			except Exception as ex:
				log.debug(f"outlookImageHelper: błąd ekstrakcji grafiki {item.index}: {ex}")
				prepared.append((item, None))

		def _worker():
			savedCount = 0
			errors = 0
			usedNames = set()
			with BusySoundTicker(startTone=True, interval=0.22, baseFreq=580, endTone=True):
				for item, tempPng in prepared:
					if not tempPng or not os.path.isfile(tempPng):
						errors += 1
						continue
					try:
						aiCand = ""
						if useAi and apiKey and item.isScreenshotOrUnlabeled and not item.ocrText:
							try:
								aiCand = describeWithGemini(
									tempPng,
									apiKey,
									customPrompt="Podaj krótką nazwę pliku (maksymalnie 3-4 słowa po polsku, bez rozszerzenia .png), która najlepiej opisuje ten obraz. Zwróć wyłącznie samą nazwę."
								)
							except Exception:
								pass
						baseFilename = generateDescriptiveFilename(item, aiCandidate=aiCand)
						candidatePath = os.path.join(targetDir, baseFilename)
						root, ext = os.path.splitext(baseFilename)
						counter = 1
						while candidatePath.lower() in usedNames or os.path.exists(candidatePath):
							candidatePath = os.path.join(targetDir, f"{root}_{counter}{ext}")
							counter += 1

						usedNames.add(candidatePath.lower())
						import shutil
						shutil.copyfile(tempPng, candidatePath)
						savedCount += 1
						try:
							os.remove(tempPng)
						except Exception:
							pass
					except Exception as ex:
						log.error(f"outlookImageHelper: Błąd seryjnego zapisu grafiki {item.index}: {ex}", exc_info=True)
						errors += 1

			def _onDone():
				if errors == 0:
					ui.message(f"{successMessagePrefix} w folderze: {os.path.basename(targetDir)}.")
				else:
					ui.message(f"Zapisano {savedCount} grafik (wystąpiły błędy przy {errors} obiektach).")
			wx.CallAfter(_onDone)

		threading.Thread(target=_worker, daemon=True).start()

	def _batchSaveAttachments(self, items: list[AttachmentItem], successMessagePrefix: str):
		"""
		Seryjny zapis załączników do wybranego folderu z pełną ochroną przed path traversal
		oraz automatycznym rozwiązywaniem kolizji nazw.
		"""
		with wx.DirDialog(
			self,
			message="Wybierz folder docelowy dla zapisywanych załączników",
			style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST,
		) as dirDialog:
			if dirDialog.ShowModal() == wx.ID_CANCEL:
				return
			targetDir = dirDialog.GetPath()

		if not os.path.isdir(targetDir):
			ui.message("Wybrany folder nie istnieje.")
			return

		absTargetDir = os.path.abspath(targetDir)
		ui.message(f"Trwa zapisywanie {len(items)} załączników do folderu...")

		savedCount = 0
		errors = 0
		usedNames = set()

		with BusySoundTicker(startTone=True, interval=0.22, baseFreq=580, endTone=True):
			for att in items:
				try:
					rawName = att.fileName or f"zalacznik_{att.index}"
					cleanName = "".join(c for c in rawName if c not in '<>:"/\\|?*').strip(" .")
					if not cleanName:
						cleanName = f"zalacznik_{att.index}"

					root, ext = os.path.splitext(cleanName)
					candidate = os.path.abspath(os.path.join(absTargetDir, f"{root}{ext}"))

					# Weryfikacja czy ścieżka mieści się w wybranym folderze
					if not candidate.startswith(absTargetDir + os.sep) and candidate != absTargetDir:
						raise ValueError("Niebezpieczna nazwa pliku próbująca wyjść poza folder docelowy.")

					# Rozwiązywanie kolizji nazw plików
					counter = 1
					while candidate.lower() in usedNames or os.path.exists(candidate):
						candidate = os.path.abspath(os.path.join(absTargetDir, f"{root}_{counter}{ext}"))
						if not candidate.startswith(absTargetDir + os.sep):
							raise ValueError("Niebezpieczna ścieżka pliku.")
						counter += 1

					usedNames.add(candidate.lower())
					att.rawAttachment.SaveAsFile(candidate)
					savedCount += 1
				except Exception as ex:
					log.error(f"outlookImageHelper: Błąd zapisu załącznika {att.fileName}: {ex}", exc_info=True)
					errors += 1

		if errors == 0:
			ui.message(f"{successMessagePrefix} w folderze: {os.path.basename(targetDir)}.")
		else:
			ui.message(f"Zapisano {savedCount} załączników (błędy przy {errors} plikach).")

	def onKeyClicked(self, event):
		gui.mainFrame.prePopup()
		try:
			d = GeminiApiKeyDialog(self)
			d.ShowModal()
			d.Destroy()
		finally:
			gui.mainFrame.postPopup()


# ==============================================================================
# PANEL USTAWIEŃ W PREFERENCJACH NVDA
# ==============================================================================


class SettingsPanelTitleDescriptor:
	"""Dynamiczny deskryptor tytułu panelu i kategorii gestów zależny od języka."""
	def __get__(self, obj, objtype=None):
		return _T("addon_title")


# ==============================================================================
# MODUŁ AKTUALIZACJI DODATKU (WCAG 2.1 / 2.2 AA)
# ==============================================================================

def _parse_version(v_str: str) -> tuple:
	parts = re.findall(r'\d+', str(v_str or ""))
	return tuple(int(p) for p in parts) if parts else (0,)

def _is_newer(remote_v: str, local_v: str) -> bool:
	return _parse_version(remote_v) > _parse_version(local_v)


class UpdateNotificationDialog(wx.Dialog):
	"""
	W pełni dostępne okno powiadomienia o aktualizacji (WCAG 2.1 / 2.2 AA).
	Zapewnia pełną obsługę klawiatury (Enter, Escape, Tab), wysoki kontrast
	oraz czytelny odczyt listy zmian.
	"""
	def __init__(self, parent, addon_title, local_ver, remote_ver, changelog, on_install_callback):
		super().__init__(
			parent,
			title=f"Aktualizacja: {addon_title}",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
		)
		self.on_install_callback = on_install_callback

		main_sizer = wx.BoxSizer(wx.VERTICAL)

		# Nagłówek
		lbl_header = wx.StaticText(self, label=f"Dostępna jest nowa wersja dodatku: {addon_title}!")
		font = lbl_header.GetFont()
		font.SetWeight(wx.FONTWEIGHT_BOLD)
		lbl_header.SetFont(font)
		main_sizer.Add(lbl_header, 0, wx.ALL, 12)

		# Informacja o wersjach
		lbl_ver = wx.StaticText(self, label=f"Zainstalowana wersja: {local_ver}   --->   Nowa wersja: {remote_ver}")
		main_sizer.Add(lbl_ver, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		# Etykieta zmian
		lbl_changes = wx.StaticText(self, label="&Wykaz nowości i ulepszeń w tej wersji:")
		main_sizer.Add(lbl_changes, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

		# Pole tekstowe z opisem zmian (dostępne, wieloliniowe, tylko do odczytu)
		self.txt_changelog = wx.TextCtrl(
			self,
			value=changelog or "Brak dodatkowego opisu zmian.",
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_BESTWRAP,
			size=(540, 160)
		)
		main_sizer.Add(self.txt_changelog, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

		# Komunikat pytający
		lbl_prompt = wx.StaticText(self, label="Czy chcesz teraz pobrać i zainstalować tę aktualizację?")
		main_sizer.Add(lbl_prompt, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

		# Przyciski akcji
		btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btn_install = wx.Button(self, wx.ID_YES, label="&Zaktualizuj teraz")
		self.btn_cancel = wx.Button(self, wx.ID_CANCEL, label="&Przypomnij później")

		self.btn_install.Bind(wx.EVT_BUTTON, self.on_install)
		self.btn_cancel.Bind(wx.EVT_BUTTON, self.on_cancel)

		btn_sizer.Add(self.btn_install, 0, wx.RIGHT, 10)
		btn_sizer.Add(self.btn_cancel, 0)
		main_sizer.Add(btn_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 12)

		self.SetSizer(main_sizer)
		main_sizer.Fit(self)
		self.CenterOnScreen()
		self.btn_install.SetDefault()
		self.btn_install.SetFocus()

	def on_install(self, event):
		self.EndModal(wx.ID_YES)
		if callable(self.on_install_callback):
			wx.CallAfter(self.on_install_callback)

	def on_cancel(self, event):
		self.EndModal(wx.ID_CANCEL)


class AddonUpdaterEngine:
	def __init__(self, addon_name: str, addon_title: str, current_version: str, addon_dir: str, config_section: str):
		self.addon_name = addon_name
		self.addon_title = addon_title
		self.current_version = current_version
		self.addon_dir = addon_dir
		self.config_section = config_section

	def get_configured_source(self) -> str:
		conf = config.conf.get(self.config_section, {})
		return conf.get("updateSourceUrl", "https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/update_info.json")

	def fetch_manifest(self) -> tuple[dict | None, str | None]:
		configured = self.get_configured_source()
		candidates = []
		if configured:
			candidates.append(configured.strip())
		candidates.extend([
			"https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/update_info.json",
			"https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/updates.json",
			r"G:\Mój dysk\Antigravity\Nvda dodatki\updates.json",
			os.path.join(self.addon_dir, "update_info.json"),
			os.path.join(self.addon_dir, "updates.json"),
			os.path.join(os.path.dirname(self.addon_dir), "updates.json"),
			r"C:\Antigravity\Nvda dodatki\updates.json",
		])

		raw_content = None
		resolved_source = None
		last_err = None

		for s in candidates:
			if not s:
				continue
			try:
				if s.startswith("http://") or s.startswith("https://"):
					req = urllib.request.Request(s, headers={"User-Agent": "NVDA-Addon-Updater/1.0"})
					with urllib.request.urlopen(req, timeout=10) as resp:
						raw_content = resp.read().decode("utf-8-sig")
						resolved_source = s
						break
				elif os.path.isfile(s):
					with open(s, "r", encoding="utf-8-sig") as f:
						raw_content = f.read()
						resolved_source = s
						break
			except Exception as ex:
				last_err = ex
				continue

		if not raw_content:
			return None, str(last_err or "Nie znaleziono pliku ogłoszeniowego na dysku ani pod wskazanym adresem.")

		try:
			try:
				data = json.loads(raw_content)
			except Exception:
				fixed = re.sub(r'\\(?![/"\\bfnrtu]|u[0-9a-fA-F]{4})', r'\\\\', raw_content)
				data = json.loads(fixed)
			if self.addon_name in data and isinstance(data[self.addon_name], dict):
				item = data[self.addon_name]
			elif isinstance(data, dict) and "version" in data:
				item = data
			else:
				return None, f"Brak wpisu dla dodatku '{self.addon_name}' w pliku ogłoszeniowym."

			item["_source"] = resolved_source
			return item, None
		except Exception as ex:
			return None, f"Błąd przetwarzania formatu JSON: {ex}"

	def check_for_updates(self, manual: bool = False):
		t = threading.Thread(target=self._worker_check, args=(manual,), daemon=True)
		t.start()

	def _worker_check(self, manual: bool):
		try:
			info, err = self.fetch_manifest()
			if not info:
				if manual:
					msg = f"Nie udało się sprawdzić aktualizacji dla dodatku {self.addon_title}.\n\nSzczegóły: {err}"
					wx.CallAfter(gui.messageBox, msg, f"Aktualizacje: {self.addon_title}", wx.OK | wx.ICON_WARNING)
				return

			remote_ver = str(info.get("version", "")).strip()
			if _is_newer(remote_ver, self.current_version):
				try:
					config.conf[self.config_section]["lastUpdateCheckTime"] = time.time()
				except Exception:
					pass
				download_url = info.get("downloadUrl", "")
				changelog = info.get("changelog", "")

				def _show_dialog():
					def _on_confirm():
						self.perform_update(download_url, remote_ver)
					dlg = UpdateNotificationDialog(
						parent=gui.mainFrame,
						addon_title=self.addon_title,
						local_ver=self.current_version,
						remote_ver=remote_ver,
						changelog=changelog,
						on_install_callback=_on_confirm
					)
					gui.mainFrame.prePopup()
					dlg.ShowModal()
					dlg.Destroy()
					gui.mainFrame.postPopup()

				wx.CallAfter(_show_dialog)
			else:
				if manual:
					msg = f"Posiadasz najnowszą wersję dodatku {self.addon_title} ({self.current_version}).\nBrak nowszych aktualizacji na Twoim dysku."
					wx.CallAfter(gui.messageBox, msg, f"Aktualizacje: {self.addon_title}", wx.OK | wx.ICON_INFORMATION)
		except Exception as ex:
			if manual:
				wx.CallAfter(gui.messageBox, f"Wystąpił nieoczekiwany błąd podczas sprawdzania aktualizacji:\n{ex}", f"Błąd: {self.addon_title}", wx.OK | wx.ICON_ERROR)

	def perform_update(self, download_url: str, remote_ver: str = ""):
		t = threading.Thread(target=self._worker_download_install, args=(download_url, remote_ver), daemon=True)
		t.start()

	def _worker_download_install(self, download_url: str, remote_ver: str = ""):
		import addonHandler

		ui.message(f"Pobieranie aktualizacji dla {self.addon_title}...")
		target_pkg = None

		try:
			if download_url.startswith("http://") or download_url.startswith("https://"):
				req = urllib.request.Request(download_url, headers={"User-Agent": "NVDA-Addon-Updater/1.0"})
				tmp_file = os.path.join(tempfile.gettempdir(), f"{self.addon_name}_update.nvda-addon")
				with urllib.request.urlopen(req, timeout=30) as resp, open(tmp_file, "wb") as out:
					shutil.copyfileobj(resp, out)
				target_pkg = tmp_file
			else:
				candidates = [download_url]
				candidates.extend([
					os.path.join(r"G:\Mój dysk\Antigravity\Nvda dodatki\Pakiety_Instalacyjne", f"{self.addon_name}.nvda-addon"),
					os.path.join(r"C:\Antigravity\Nvda dodatki\Pakiety_Instalacyjne", f"{self.addon_name}.nvda-addon"),
				])
				for cand in candidates:
					if cand and os.path.isfile(cand):
						target_pkg = cand
						break

			if not target_pkg or not os.path.isfile(target_pkg):
				raise FileNotFoundError(f"Nie znaleziono pliku instalacyjnego dodatku:\n{download_url}")

			def _do_install():
				installed_ok = False
				try:
					if hasattr(addonHandler, "installAddonBundle"):
						try:
							bundle = addonHandler.AddonBundle(target_pkg)
							addonHandler.installAddonBundle(bundle)
							installed_ok = True
						except TypeError:
							addonHandler.installAddonBundle(target_pkg)
							installed_ok = True
					elif hasattr(addonHandler, "installAddonPackage"):
						try:
							bundle = addonHandler.AddonBundle(target_pkg)
							addonHandler.installAddonPackage(bundle)
							installed_ok = True
						except TypeError:
							addonHandler.installAddonPackage(target_pkg)
							installed_ok = True
					elif hasattr(addonHandler, "installAddon"):
						bundle = addonHandler.AddonBundle(target_pkg)
						addonHandler.installAddon(bundle)
						installed_ok = True
				except Exception as ex:
					log.debug(f"{self.addon_name} updater internal install error: {ex}")

				if installed_ok:
					ver_text = f" do wersji {remote_ver}" if remote_ver else ""
					if gui.messageBox(
						f"Pomyślnie zaktualizowano dodatek {self.addon_title}{ver_text}.\n\nCzy chcesz teraz zrestartować NVDA, aby zastosować zmiany?",
						"Aktualizacja zakończona sukcesem",
						wx.YES_NO | wx.ICON_QUESTION
					) == wx.YES:
						import core
						core.restart()
					return

				try:
					os.startfile(target_pkg)
				except Exception as ex2:
					gui.messageBox(f"Nie udało się uruchomić instalatora:\n{ex2}", "Błąd instalacji", wx.OK | wx.ICON_ERROR)

			wx.CallAfter(_do_install)
		except Exception as ex:
			wx.CallAfter(gui.messageBox, f"Nie udało się pobrać aktualizacji:\n{ex}", f"Błąd aktualizacji: {self.addon_title}", wx.OK | wx.ICON_ERROR)


# ==============================================================================
# NARZĘDZIA TWÓRCY I ADMINISTRATORA DODATKU (BEZPIECZEŃSTWO, TŁUMACZENIA, AKTUALIZACJE)
# ==============================================================================

_ENC_AUTHOR_RECOVERY = "PD06KjQEERAZFgsCFD5bXBxVICQ="
_AUTHOR_ENC_KEY = b"OIH_SecureKey_2026"
_ADMIN_SALT = "OIH_CreatorSalt_2026#"
# SHA-256 salted hash dla domyślnego hasła "admin"
_DEFAULT_ADMIN_HASH = "e8590188bd0a58811de48cdd23f90969703e27d782e0dc4d1d2cfb47df4f7cb9"


def get_active_settings_panel():
	"""Zwraca aktywny panel ustawień, jeśli jest otwarty w oknie preferencji NVDA."""
	global _activeSettingsPanelRef
	if _activeSettingsPanelRef is not None:
		try:
			return _activeSettingsPanelRef()
		except Exception:
			pass
	return None


def getAuthorRecoveryEmail() -> str:
	"""Zwraca odszyfrowany w locie adres e-mail twórcy do odzyskiwania dostępu."""
	try:
		import base64
		raw = base64.b64decode(_ENC_AUTHOR_RECOVERY)
		return bytes(b ^ _AUTHOR_ENC_KEY[i % len(_AUTHOR_ENC_KEY)] for i, b in enumerate(raw)).decode("utf-8")
	except Exception as ex:
		log.error(f"outlookImageHelper: Błąd deszyfrowania adresu e-mail twórcy: {ex}")
		return ""


def _hashPassword(pwd: str) -> str:
	"""Generuje bezpieczny, solony skrót kryptograficzny SHA-256 z podanego hasła."""
	import hashlib
	val = (_ADMIN_SALT + (pwd or "").strip()).encode("utf-8")
	return hashlib.sha256(val).hexdigest().lower()


def getAdminPasswordHash() -> str:
	"""Zwraca zapisany skrót hasła twórcy z konfiguracji NVDA lub domyślny skrót."""
	try:
		sec = config.conf.get(CONFIG_SECTION, {})
		stored = str(sec.get("adminPasswordHash", "") or "").strip().lower()
		if stored:
			return stored
		# Automatyczna migracja starego hasła jawnego, jeśli istniało
		legacy = str(sec.get("adminPassword", "") or "").strip()
		if legacy:
			newHash = _hashPassword(legacy)
			try:
				sec["adminPasswordHash"] = newHash
				if "adminPassword" in sec:
					del sec["adminPassword"]
				config.conf.save()
			except Exception:
				pass
			return newHash
		return _DEFAULT_ADMIN_HASH
	except Exception:
		return _DEFAULT_ADMIN_HASH


def setAdminPassword(newPwd: str):
	"""Zapisuje solony skrót SHA-256 nowego hasła twórcy w konfiguracji NVDA (bez tekstu jawnego)."""
	try:
		sec = config.conf[CONFIG_SECTION]
		sec["adminPasswordHash"] = _hashPassword(newPwd)
		if "adminPassword" in sec:
			del sec["adminPassword"]
		config.conf.save()
	except Exception as ex:
		log.error(f"outlookImageHelper: Błąd zapisu hasła twórcy: {ex}", exc_info=True)


def verifyAdminPassword(entered: str) -> bool:
	"""Weryfikuje podane hasło poprzez porównanie skrótów kryptograficznych."""
	val = (entered or "").strip()
	if not val:
		return False
	h = _hashPassword(val)
	return h == getAdminPasswordHash()


def sendPasswordResetEmail(targetEmail: str = "", resetCode: str = "") -> tuple[bool, str]:
	"""Wysyła jednorazowy kod weryfikacyjny na podany adres e-mail przez program Microsoft Outlook."""
	try:
		if not targetEmail:
			targetEmail = getAuthorRecoveryEmail()
		if not targetEmail:
			return False, "Brak zdefiniowanego adresu e-mail twórcy."
		olApp = getOutlookApp()
		if not olApp:
			return False, "Nie udało się połączyć z programem Microsoft Outlook. Upewnij się, że Outlook jest uruchomiony."
		mail = olApp.CreateItem(0)
		mail.To = targetEmail
		mail.Subject = f"Outlook Image Helper — Kod resetowania hasła: {resetCode}"
		body = (
			"Dzień dobry,\n\n"
			"Otrzymano żądanie zresetowania hasła twórcy w dodatku Outlook Image Helper.\n\n"
			f"Twój jednorazowy kod weryfikacyjny to: {resetCode}\n\n"
			"Wprowadź powyższy 6-cyfrowy kod w oknie odzyskiwania hasła w NVDA, aby zdefiniować nowe hasło.\n\n"
			"Jeśli nie inicjowałeś tej operacji, zignoruj tę wiadomość.\n\n"
			"— Outlook Image Helper & Antigravity"
		)
		mail.Body = body
		mail.Send()
		return True, "Kod weryfikacyjny został wysłany."
	except Exception as ex:
		log.error(f"outlookImageHelper: Błąd wysyłania e-maila resetującego hasło: {ex}", exc_info=True)
		return False, f"Błąd wysyłania wiadomości przez Outlook: {ex}"


class ChangeAdminPasswordDialog(wx.Dialog):
	"""Dostępne okno zmiany hasła twórcy dodatku."""
	def __init__(self, parent, isReset: bool = False):
		super().__init__(
			parent,
			title="Ustaw nowe hasło twórcy dodatku" if isReset else "Zmień hasło twórcy dodatku",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.isReset = isReset
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		if not isReset:
			lblOld = wx.StaticText(self, label="&Aktualne hasło twórcy:")
			mainSizer.Add(lblOld, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
			self.oldPwdCtrl = wx.TextCtrl(self, style=wx.TE_PASSWORD, size=(360, -1))
			self.oldPwdCtrl.SetName("Aktualne hasło twórcy")
			mainSizer.Add(self.oldPwdCtrl, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		lblNew1 = wx.StaticText(self, label="&Nowe hasło twórcy:")
		mainSizer.Add(lblNew1, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
		self.newPwd1Ctrl = wx.TextCtrl(self, style=wx.TE_PASSWORD, size=(360, -1))
		self.newPwd1Ctrl.SetName("Nowe hasło twórcy")
		mainSizer.Add(self.newPwd1Ctrl, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		lblNew2 = wx.StaticText(self, label="&Powtórz nowe hasło:")
		mainSizer.Add(lblNew2, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
		self.newPwd2Ctrl = wx.TextCtrl(self, style=wx.TE_PASSWORD, size=(360, -1))
		self.newPwd2Ctrl.SetName("Powtórz nowe hasło")
		mainSizer.Add(self.newPwd2Ctrl, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnSave = wx.Button(self, wx.ID_OK, label="&Zapisz nowe hasło")
		self.btnSave.SetDefault()
		btnSizer.Add(self.btnSave, 0, wx.ALL, 5)

		self.btnCancel = wx.Button(self, wx.ID_CANCEL, label="&Anuluj")
		btnSizer.Add(self.btnCancel, 0, wx.ALL, 5)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(mainSizer)
		self.Fit()
		self.CenterOnScreen()

		self.btnSave.Bind(wx.EVT_BUTTON, self.onSave)
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		else:
			event.Skip()

	def onSave(self, event):
		if not self.isReset:
			curOld = self.oldPwdCtrl.GetValue().strip()
			if not verifyAdminPassword(curOld):
				ui.message("Podane aktualne hasło jest nieprawidłowe.")
				self.oldPwdCtrl.SetFocus()
				return
		p1 = self.newPwd1Ctrl.GetValue().strip()
		p2 = self.newPwd2Ctrl.GetValue().strip()
		if not p1:
			ui.message("Nowe hasło nie może być puste.")
			self.newPwd1Ctrl.SetFocus()
			return
		if p1 != p2:
			ui.message("Wprowadzone nowe hasła nie są identyczne.")
			self.newPwd2Ctrl.SetFocus()
			return
		setAdminPassword(p1)
		ui.message("Hasło twórcy dodatku zostało pomyślnie zmienione.")
		self.EndModal(wx.ID_OK)


class PasswordRecoveryDialog(wx.Dialog):
	"""Dostępne okno odzyskiwania hasła twórcy dodatku przy użyciu autoryzowanego adresu e-mail."""
	def __init__(self, parent):
		super().__init__(
			parent,
			title="Odzyskiwanie hasła twórcy dodatku",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.recoveryEmail = getAuthorRecoveryEmail()
		import random
		self.expectedCode = f"{random.randint(100000, 999999)}"

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		self.infoLbl = wx.StaticText(
			self,
			label="Na powiązany adres e-mail twórcy dodatku został wysłany jednorazowy 6-cyfrowy kod weryfikacyjny.\nSprawdź swoją skrzynkę odbiorczą i wprowadź otrzymany kod poniżej:"
		)
		mainSizer.Add(self.infoLbl, 0, wx.ALL, 10)

		lblCode = wx.StaticText(self, label="&Wprowadź 6-cyfrowy kod weryfikacyjny:")
		mainSizer.Add(lblCode, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)

		self.codeCtrl = wx.TextCtrl(self, size=(240, -1))
		self.codeCtrl.SetName("Wprowadź 6-cyfrowy kod weryfikacyjny")
		mainSizer.Add(self.codeCtrl, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		actionSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnResend = wx.Button(self, label="&Wyślij kod ponownie")
		self.btnResend.Bind(wx.EVT_BUTTON, lambda evt: self.doSendCode())
		actionSizer.Add(self.btnResend, 0, wx.RIGHT, 10)

		self.btnVerify = wx.Button(self, wx.ID_OK, label="&Zatwierdź kod i ustaw nowe hasło")
		self.btnVerify.SetDefault()
		self.btnVerify.Bind(wx.EVT_BUTTON, self.onVerify)
		actionSizer.Add(self.btnVerify, 0, wx.RIGHT, 10)

		self.btnCancel = wx.Button(self, wx.ID_CANCEL, label="&Anuluj")
		actionSizer.Add(self.btnCancel, 0)
		mainSizer.Add(actionSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(mainSizer)
		self.Fit()
		self.CenterOnScreen()

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		wx.CallLater(150, self.doSendCode)

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		else:
			event.Skip()

	def doSendCode(self):
		if not self.recoveryEmail:
			ui.message("Błąd autoryzacji: brak powiązanego adresu twórcy.")
			return
		ok, msg = sendPasswordResetEmail(self.recoveryEmail, self.expectedCode)
		if ok:
			ui.message("Wysłano kod weryfikacyjny na autoryzowany adres e-mail twórcy dodatku.")
		else:
			ui.message(f"Błąd wysyłania: {msg}")

	def onVerify(self, event):
		val = self.codeCtrl.GetValue().strip()
		if not val:
			ui.message("Wpisz kod weryfikacyjny.")
			return
		if self.expectedCode and val == self.expectedCode:
			self.EndModal(wx.ID_OK)
			dlg = ChangeAdminPasswordDialog(self.GetParent(), isReset=True)
			dlg.ShowModal()
			dlg.Destroy()
		else:
			ui.message("Nieprawidłowy kod weryfikacyjny. Spróbuj ponownie lub wyślij nowy kod.")
			self.codeCtrl.SetFocus()


class AdminLoginDialog(wx.Dialog):
	"""Okno logowania twórcy dodatku z opcją zmiany i bezpiecznego odzyskiwania hasła."""
	def __init__(self, parent):
		super().__init__(
			parent,
			title="Autoryzacja — Narzędzia twórcy dodatku",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.authenticated = False
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		lblInfo = wx.StaticText(
			self,
			label="Dostęp do narzędzi twórcy wymaga podania hasła:"
		)
		mainSizer.Add(lblInfo, 0, wx.ALL, 10)

		lblPwd = wx.StaticText(self, label="&Hasło twórcy dodatku:")
		mainSizer.Add(lblPwd, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)

		self.pwdCtrl = wx.TextCtrl(self, style=wx.TE_PASSWORD, size=(340, -1))
		self.pwdCtrl.SetName("Hasło twórcy dodatku")
		mainSizer.Add(self.pwdCtrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnLogin = wx.Button(self, wx.ID_OK, label="&Zaloguj i otwórz panel twórcy")
		self.btnLogin.SetDefault()
		btnSizer.Add(self.btnLogin, 0, wx.RIGHT, 10)

		self.btnChangePwd = wx.Button(self, label="&Zmień hasło...")
		self.btnChangePwd.Bind(wx.EVT_BUTTON, self.onChangePwd)
		btnSizer.Add(self.btnChangePwd, 0, wx.RIGHT, 10)

		self.btnForgot = wx.Button(self, label="&Odzyskaj hasło...")
		self.btnForgot.Bind(wx.EVT_BUTTON, self.onForgot)
		btnSizer.Add(self.btnForgot, 0, wx.RIGHT, 10)

		self.btnCancel = wx.Button(self, wx.ID_CANCEL, label="&Anuluj")
		btnSizer.Add(self.btnCancel, 0)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(mainSizer)
		self.Fit()
		self.CenterOnScreen()

		self.btnLogin.Bind(wx.EVT_BUTTON, self.onLogin)
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		self.pwdCtrl.SetFocus()

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		else:
			event.Skip()

	def onLogin(self, event):
		val = self.pwdCtrl.GetValue().strip()
		if verifyAdminPassword(val):
			self.authenticated = True
			self.EndModal(wx.ID_OK)
		else:
			ui.message("Niepoprawne hasło twórcy dodatku.")
			self.pwdCtrl.SetFocus()

	def onChangePwd(self, event):
		dlg = ChangeAdminPasswordDialog(self, isReset=False)
		dlg.ShowModal()
		dlg.Destroy()

	def onForgot(self, event):
		dlg = PasswordRecoveryDialog(self)
		dlg.ShowModal()
		dlg.Destroy()


class CreatorAdminDialog(wx.Dialog):
	"""
	Dedykowane okno narzędzi twórcy lub administratora dodatku Outlook Image Helper.
	Gromadzi generator wersji językowych, eksport/import tłumaczeń,
	konfigurację ścieżki pliku ogłoszeniowego aktualizacji oraz zarządzanie hasłem.
	"""
	def __init__(self, parent):
		super().__init__(
			parent,
			title="Narzędzia twórcy dodatku — Outlook Image Helper",
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
			size=(640, 480)
		)
		sec = config.conf.get(CONFIG_SECTION, {})

		mainSizer = wx.BoxSizer(wx.VERTICAL)

		# 1. Narzędzia lokalizacji i tłumaczeń
		sbLang = wx.StaticBox(self, label="1. Wersje językowe, szablony i generator tłumaczeń AI")
		sboxLang = wx.StaticBoxSizer(sbLang, wx.VERTICAL)

		row1 = wx.BoxSizer(wx.HORIZONTAL)
		self.btnTranslateAi = wx.Button(self, label=_T("set_btn_translate_ai"))
		self.btnTranslateAi.Bind(wx.EVT_BUTTON, self.onTranslateAi)
		row1.Add(self.btnTranslateAi, 0, wx.RIGHT, 10)

		self.btnExportLang = wx.Button(self, label=_T("set_btn_export_lang"))
		self.btnExportLang.Bind(wx.EVT_BUTTON, self.onExportLang)
		row1.Add(self.btnExportLang, 0)
		sboxLang.Add(row1, 0, wx.ALL, 5)

		row2 = wx.BoxSizer(wx.HORIZONTAL)
		self.btnImportLang = wx.Button(self, label=_T("set_btn_import_lang"))
		self.btnImportLang.Bind(wx.EVT_BUTTON, self.onImportLang)
		row2.Add(self.btnImportLang, 0, wx.RIGHT, 10)

		self.btnOpenLangFolder = wx.Button(self, label=_T("set_btn_open_lang_dir"))
		self.btnOpenLangFolder.Bind(wx.EVT_BUTTON, lambda evt: os.startfile(CUSTOM_LANGUAGES_DIR))
		row2.Add(self.btnOpenLangFolder, 0)
		sboxLang.Add(row2, 0, wx.ALL, 5)

		mainSizer.Add(sboxLang, 0, wx.EXPAND | wx.ALL, 10)

		# 2. Konfiguracja pliku aktualizacji dodatku
		sbUpd = wx.StaticBox(self, label="2. Konfiguracja pliku ogłoszeniowego aktualizacji")
		sboxUpd = wx.StaticBoxSizer(sbUpd, wx.VERTICAL)

		lblUpd = wx.StaticText(self, label="&Lokalizacja pliku ogłoszeniowego (ścieżka na dysku lub URL):")
		sboxUpd.Add(lblUpd, 0, wx.LEFT | wx.RIGHT | wx.TOP, 5)

		defaultUpd = sec.get("updateSourceUrl", "https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/update_info.json")
		isOldDefault = bool(defaultUpd and ("mój dysk" in str(defaultUpd).lower() and "antigravity" in str(defaultUpd).lower()))
		if not defaultUpd or isOldDefault:
			defaultUpd = "https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/update_info.json"
		self.txtUpdateSource = wx.TextCtrl(self, value=str(defaultUpd))
		self.txtUpdateSource.SetName("Lokalizacja pliku ogłoszeniowego (ścieżka na dysku lub URL):")
		sboxUpd.Add(self.txtUpdateSource, 0, wx.EXPAND | wx.ALL, 5)

		updBtnBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnCheckUpdatesNow = wx.Button(self, label="&Sprawdź dostępność aktualizacji teraz")
		self.btnCheckUpdatesNow.Bind(wx.EVT_BUTTON, self.onCheckUpdates)
		updBtnBox.Add(self.btnCheckUpdatesNow, 0, wx.RIGHT, 10)

		self.btnResetGitHub = wx.Button(self, label="&Przywróć oficjalny adres GitHub")
		def onResetGh(evt):
			self.txtUpdateSource.SetValue("https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/update_info.json")
			ui.message("Wstawiono oficjalny adres aktualizacji GitHub.")
		self.btnResetGitHub.Bind(wx.EVT_BUTTON, onResetGh)
		updBtnBox.Add(self.btnResetGitHub, 0)
		sboxUpd.Add(updBtnBox, 0, wx.ALL, 5)

		mainSizer.Add(sboxUpd, 0, wx.EXPAND | wx.ALL, 10)

		# 3. Bezpieczeństwo i hasło
		sbSec = wx.StaticBox(self, label="3. Bezpieczeństwo i hasło twórcy dodatku")
		sboxSec = wx.StaticBoxSizer(sbSec, wx.HORIZONTAL)

		self.btnChangePwd = wx.Button(self, label="&Zmień hasło twórcy...")
		self.btnChangePwd.Bind(wx.EVT_BUTTON, self.onChangePwd)
		sboxSec.Add(self.btnChangePwd, 0, wx.RIGHT, 10)

		lblRecoverInfo = wx.StaticText(self, label="Odzyskiwanie hasła powiązane z autoryzowanym kontem e-mail twórcy.")
		sboxSec.Add(lblRecoverInfo, 0, wx.ALIGN_CENTER_VERTICAL)

		mainSizer.Add(sboxSec, 0, wx.EXPAND | wx.ALL, 10)

		# Przyciski dolne
		btnBottomSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnSaveClose = wx.Button(self, wx.ID_OK, label="&Zapisz i zamknij")
		self.btnSaveClose.SetDefault()
		self.btnSaveClose.Bind(wx.EVT_BUTTON, self.onSaveAndClose)
		btnBottomSizer.Add(self.btnSaveClose, 0, wx.RIGHT, 10)

		self.btnClose = wx.Button(self, wx.ID_CANCEL, label="&Zamknij")
		btnBottomSizer.Add(self.btnClose, 0)

		mainSizer.Add(btnBottomSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(mainSizer)
		self.Fit()
		self.CenterOnScreen()

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CANCEL)
		else:
			event.Skip()

	def onSaveAndClose(self, event):
		val = self.txtUpdateSource.GetValue().strip()
		if val:
			config.conf[CONFIG_SECTION]["updateSourceUrl"] = val
			try:
				config.conf.save()
			except Exception:
				pass
		self.EndModal(wx.ID_OK)
		ui.message("Zapisano ustawienia twórcy dodatku.")

	def onTranslateAi(self, event):
		def onGen(code):
			panel = get_active_settings_panel()
			if panel:
				panel.refreshLanguageChoices(selectCode=code)
				config.conf[CONFIG_SECTION]["language"] = code
				panel.updateUILabels()
		dlg = GenerateLanguageDialog(self, onGeneratedCallback=onGen)
		dlg.ShowModal()
		dlg.Destroy()

	def onExportLang(self, event):
		with wx.FileDialog(
			self,
			message=_T("dlg_export_title"),
			wildcard="JSON files (*.json)|*.json",
			style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
			defaultFile="outlookImageHelper_language_template.json",
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				path = dlg.GetPath()
				if exportLanguageTemplate(path):
					ui.message(_T("dlg_export_success", file=os.path.basename(path)))
				else:
					ui.message("Nie udało się wyeksportować szablonu.")

	def onImportLang(self, event):
		with wx.FileDialog(
			self,
			message=_T("dlg_import_title"),
			wildcard="JSON files (*.json)|*.json",
			style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				path = dlg.GetPath()
				ok, code, name = importLanguageFile(path)
				if ok:
					panel = get_active_settings_panel()
					if panel:
						panel.refreshLanguageChoices(selectCode=code)
						config.conf[CONFIG_SECTION]["language"] = code
						panel.updateUILabels()
					ui.message(_T("dlg_import_success", name=name, code=code))
				else:
					ui.message(_T("dlg_import_err_invalid", err=name))

	def onCheckUpdates(self, event):
		updater = AddonUpdaterEngine(
			addon_name=CONFIG_SECTION,
			addon_title=_T("addon_title"),
			current_version=ADDON_VERSION,
			addon_dir=ADDON_DIR,
			config_section=CONFIG_SECTION,
		)
		updater.check_for_updates(manual=True)

	def onChangePwd(self, event):
		dlg = ChangeAdminPasswordDialog(self, isReset=False)
		dlg.ShowModal()
		dlg.Destroy()


def openCreatorAdminTools(parent=None):
	"""Otwiera okno autoryzacji twórcy, a po poprawnym haśle uruchamia panel administratora."""
	dlgLogin = AdminLoginDialog(parent or gui.mainFrame)
	if dlgLogin.ShowModal() == wx.ID_OK and dlgLogin.authenticated:
		dlgLogin.Destroy()
		dlgAdmin = CreatorAdminDialog(parent or gui.mainFrame)
		dlgAdmin.ShowModal()
		dlgAdmin.Destroy()
	else:
		dlgLogin.Destroy()


class AddonAboutDialog(wx.Dialog):
	"""
	W pełni dostępne okno z informacjami o dodatku (WCAG 2.1 / 2.2 AA).
	Pozwala użytkownikowi czytnika ekranu wejść fokusem (Tab) w pole tekstowe,
	przeczytać wersję, datę kompilacji, autora, licencję oraz uruchomić sprawdzanie
	aktualizacji lub podręcznik użytkownika.
	"""
	def __init__(self, parent=None):
		effLang = getEffectiveLanguage()
		title = "O dodatku: " + ADDON_TITLE if effLang == "pl" else "About: " + ADDON_TITLE
		super().__init__(
			parent or gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		lblHeader = wx.StaticText(self, label=ADDON_TITLE)
		f = lblHeader.GetFont()
		f.SetWeight(wx.FONTWEIGHT_BOLD)
		lblHeader.SetFont(f)
		mainSizer.Add(lblHeader, 0, wx.ALL, 12)

		lblDetails = wx.StaticText(
			self,
			label="&Szczegółowe informacje o dodatku i wersji:" if effLang == "pl" else "&Add-on details and version info:"
		)
		mainSizer.Add(lblDetails, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

		infoText = (
			f"Wersja dodatku: {ADDON_VERSION} (Kompilacja {ADDON_BUILD_ID})\n"
			f"Data kompilacji: {ADDON_BUILD_DATE}\n"
			f"Twórca: {ADDON_AUTHOR}\n"
			f"Zgodność: {ADDON_COMPATIBILITY}\n"
			f"Licencja: {ADDON_LICENSE}\n\n"
			"Opis dodatku:\n"
			"Wykrywa, listuje, kopiuje oraz wykonuje OCR i interaktywne opisywanie/rozmowę "
			"przez sztuczną inteligencję o grafikach, zrzutach ekranu oraz załącznikach "
			"(pliki PDF, DOCX, archiwa, obrazy) w programie Microsoft Outlook Classic."
		) if effLang == "pl" else (
			f"Add-on Version: {ADDON_VERSION} (Build {ADDON_BUILD_ID})\n"
			f"Build Date: {ADDON_BUILD_DATE}\n"
			f"Author: {ADDON_AUTHOR}\n"
			f"Compatibility: {ADDON_COMPATIBILITY}\n"
			f"License: {ADDON_LICENSE}\n\n"
			"Add-on Description:\n"
			"Detects, lists, copies, and performs OCR and interactive AI descriptions "
			"for inline graphics, screenshots, and attachments (PDF, DOCX, archives, images) "
			"in Microsoft Outlook Classic."
		)

		self.txtInfo = wx.TextCtrl(
			self,
			value=infoText,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_BESTWRAP,
			size=(540, 190)
		)
		self.txtInfo.SetName("Informacje o dodatku" if effLang == "pl" else "Add-on Information")
		mainSizer.Add(self.txtInfo, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

		# Przyciski akcji
		btnSizer = wx.BoxSizer(wx.HORIZONTAL)

		self.btnCheckUpdates = wx.Button(
			self,
			label="&Sprawdź aktualizacje dodatku..." if effLang == "pl" else "&Check for updates..."
		)
		self.btnCheckUpdates.Bind(wx.EVT_BUTTON, self.onCheckUpdates)
		btnSizer.Add(self.btnCheckUpdates, 0, wx.RIGHT, 10)

		self.btnGuide = wx.Button(
			self,
			label="&Podręcznik użytkownika..." if effLang == "pl" else "&User Guide..."
		)
		self.btnGuide.Bind(wx.EVT_BUTTON, lambda evt: openUserGuide())
		btnSizer.Add(self.btnGuide, 0, wx.RIGHT, 10)

		self.btnShare = wx.Button(
			self,
			label="&Podziel się dodatkiem..." if effLang == "pl" else "&Share add-on..."
		)
		self.btnShare.Bind(wx.EVT_BUTTON, lambda evt: showShareAddonDialog(self))
		btnSizer.Add(self.btnShare, 0, wx.RIGHT, 10)

		self.btnClose = wx.Button(self, wx.ID_CLOSE, label="&Zamknij" if effLang == "pl" else "&Close")
		self.btnClose.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CLOSE))
		btnSizer.Add(self.btnClose, 0)

		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 12)

		self.SetSizer(mainSizer)
		mainSizer.Fit(self)
		self.CenterOnScreen()

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		self.txtInfo.SetFocus()

	def onCharHook(self, evt):
		if evt.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CLOSE)
		else:
			evt.Skip()

	def onCheckUpdates(self, evt):
		updater = AddonUpdaterEngine(
			addon_name=CONFIG_SECTION,
			addon_title=ADDON_TITLE,
			current_version=ADDON_VERSION,
			addon_dir=ADDON_DIR,
			config_section=CONFIG_SECTION
		)
		updater.check_for_updates(manual=True)


class ShareAddonDialog(wx.Dialog):
	"""
	W pełni dostępne okno umożliwiające podzielenie się dodatkiem ze znajomymi.
	Zawiera linki do pobrania z oficjalnego profilu GitHub, gotowy tekst polecający
	oraz przyciski kopiowania do schowka i bezpośredniej wysyłki wiadomości e-mail.
	"""

	def __init__(self, parent):
		effLang = getEffectiveLanguage()
		isPl = (effLang == "pl")
		title = "Podziel się dodatkiem ze znajomymi" if isPl else "Share Add-on with Others"
		super().__init__(
			parent or gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
		)
		self.SetMinSize((640, 460))
		mainSizer = wx.BoxSizer(wx.VERTICAL)

		lblIntro = wx.StaticText(
			self,
			label=(
				"Skorzystaj z poniższych przycisków, aby skopiować link do pobrania, "
				"skopiować gotową treść lub od razu otworzyć nowy e-mail w programie Outlook:"
				if isPl else
				"Use the buttons below to copy the download link, copy the recommendation text, "
				"or instantly open a new email message in Outlook:"
			)
		)
		mainSizer.Add(lblIntro, 0, wx.ALL, 10)

		downloadUrl = "https://github.com/strugo10/outlookImageHelper/releases/latest/download/outlookImageHelper.nvda-addon"
		repoUrl = "https://github.com/strugo10/outlookImageHelper"

		shareBody = (
			"Cześć,\n\n"
			"Polecam Ci bezpłatny i w pełni dostępny dodatek dla czytnika ekranu NVDA:\n"
			"Outlook - Pomocnik obrazów, załączników i OCR (dedykowany dla Microsoft Outlook Classic).\n\n"
			"Kluczowe możliwości dodatku:\n"
			"• Automatyczne wykrywanie i lista wszystkich grafik, zrzutów ekranu i linii podziału w mailu,\n"
			"• Szybkie rozpoznawanie tekstu (OCR) ze zdjęć oraz skanowanych dokumentów PDF,\n"
			"• Szczegółowe opisywanie obrazów i interaktywna rozmowa z AI (Google Gemini / ChatGPT),\n"
			"• Powiększony podgląd z trybem inwersji kolorów (wysoki kontrast / negatyw) dla osób słabowidzących,\n"
			"• Seryjny zapis i obsługa załączników plikowych (PDF, Word DOCX, archiwa, obrazy).\n\n"
			f"Bezpośredni link do pobrania najnowszej wersji instalatora (.nvda-addon):\n{downloadUrl}\n\n"
			f"Oficjalna strona projektu na GitHubie:\n{repoUrl}\n\n"
			"Instalacja:\n"
			"Pobierz plik 'outlookImageHelper.nvda-addon' i otwórz go klawiszem Enter w Eksploratorze Windows, "
			"a następnie potwierdź instalację w programie NVDA.\n\n"
			"Wymagania: klasyczny program pocztowy Microsoft Outlook Classic (wersje 2013-365) oraz NVDA 2024.1+."
		) if isPl else (
			"Hello,\n\n"
			"I would like to recommend a free and accessible add-on for the NVDA screen reader:\n"
			"Outlook Image, Attachment & OCR Helper (specifically for Microsoft Outlook Classic).\n\n"
			"Key Features:\n"
			"• Automatic detection and listing of all inline graphics, screenshots, and divider lines in emails,\n"
			"• Fast optical character recognition (OCR) for images and scanned PDF documents,\n"
			"• Rich image descriptions and interactive Q&A chat with AI (Google Gemini / ChatGPT),\n"
			"• Magnified preview with high-contrast color inversion (negative mode) for low-vision users,\n"
			"• Batch saving and file attachment processing (PDF, Word DOCX, archives, images).\n\n"
			f"Direct download link for latest installer (.nvda-addon):\n{downloadUrl}\n\n"
			f"Official GitHub Project Page:\n{repoUrl}\n\n"
			"Installation:\n"
			"Download 'outlookImageHelper.nvda-addon', press Enter in Windows Explorer to open it, "
			"and confirm installation in NVDA.\n\n"
			"Compatibility: Microsoft Outlook Classic (2013-365) and NVDA 2024.1+."
		)

		self.txtShare = wx.TextCtrl(
			self,
			value=shareBody,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_BESTWRAP,
			size=(-1, 230)
		)
		self.txtShare.SetName("Gotowy tekst wiadomości z łączami do pobrania" if isPl else "Share message text with download links")
		mainSizer.Add(self.txtShare, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		btnSizer = wx.WrapSizer(wx.HORIZONTAL)

		btnCopyLink = wx.Button(self, label="&Kopiuj link do pobrania" if isPl else "&Copy download link")
		def onCopyLink(evt):
			if api.copyToClip(downloadUrl):
				ui.message("Skopiowano bezpośredni link do pobrania dodatku do schowka." if isPl else "Download link copied to clipboard.")
			tones.beep(880, 50)
		btnCopyLink.Bind(wx.EVT_BUTTON, onCopyLink)
		btnSizer.Add(btnCopyLink, 0, wx.ALL, 4)

		btnCopyMsg = wx.Button(self, label="Kopiuj &całą treść" if isPl else "Copy &entire text")
		def onCopyMsg(evt):
			if api.copyToClip(shareBody):
				ui.message("Skopiowano kompletną treść polecającą do schowka." if isPl else "Recommendation text copied to clipboard.")
			tones.beep(880, 50)
		btnCopyMsg.Bind(wx.EVT_BUTTON, onCopyMsg)
		btnSizer.Add(btnCopyMsg, 0, wx.ALL, 4)

		btnEmail = wx.Button(self, label="&Wyślij e-mail z poleceniem..." if isPl else "&Send email recommendation...")
		def onEmail(evt):
			try:
				import urllib.parse
				subj = "Polecam bezpłatny dodatek: Outlook Image Helper dla NVDA" if isPl else "Recommended free NVDA add-on: Outlook Image Helper"
				mailto = f"mailto:?subject={urllib.parse.quote(subj)}&body={urllib.parse.quote(shareBody)}"
				os.startfile(mailto)
			except Exception as ex:
				log.debug(f"Share mailto error: {ex}")
		btnEmail.Bind(wx.EVT_BUTTON, onEmail)
		btnSizer.Add(btnEmail, 0, wx.ALL, 4)

		btnClose = wx.Button(self, wx.ID_CLOSE, label="&Zamknij" if isPl else "&Close")
		btnClose.Bind(wx.EVT_BUTTON, lambda evt: self.EndModal(wx.ID_CLOSE))
		btnSizer.Add(btnClose, 0, wx.ALL, 4)

		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 8)
		self.SetSizer(mainSizer)
		self.Layout()
		self.CenterOnScreen()

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		wx.CallLater(50, lambda: self.txtShare.SetFocus())

	def onCharHook(self, evt):
		if evt.GetKeyCode() == wx.WXK_ESCAPE:
			self.EndModal(wx.ID_CLOSE)
		else:
			evt.Skip()


def showShareAddonDialog(parent=None):
	dlg = ShareAddonDialog(parent)
	gui.mainFrame.prePopup()
	try:
		dlg.ShowModal()
	finally:
		dlg.Destroy()
		gui.mainFrame.postPopup()


class OutlookImageHelperSettingsPanel(gui.settingsDialogs.SettingsPanel):
	title = SettingsPanelTitleDescriptor()

	def makeSettings(self, settingsSizer):
		global _activeSettingsPanelRef
		_activeSettingsPanelRef = weakref.ref(self)
		try:
			self._makeSettingsInternal(settingsSizer)
		except Exception as e:
			log.error(f"outlookImageHelper: Błąd podczas budowania panelu ustawień: {e}", exc_info=True)
			errLabel = wx.StaticText(self, label=f"Wystąpił błąd podczas ładowania panelu ustawień Outlook Image Helper: {e}")
			settingsSizer.Add(errLabel, flag=wx.ALL, border=15)

	def _makeSettingsInternal(self, settingsSizer):
		self.sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		sHelper = self.sHelper
		try:
			sec = config.conf[CONFIG_SECTION]
		except Exception:
			sec = {}
		effLang = getEffectiveLanguage()

		# 1. Powiadomienia (PIERWSZY ELEMENT USTAWIEŃ DODATKU)
		self.autoAnnounceChk = wx.CheckBox(self, label=_T("set_auto_announce"))
		self.autoAnnounceChk.SetValue(_to_bool(sec.get("autoAnnounce", True), default=True))
		sHelper.addItem(self.autoAnnounceChk)

		self.playSoundChk = wx.CheckBox(self, label=_T("set_play_sound"))
		self.playSoundChk.SetValue(_to_bool(sec.get("playSound", True), default=True))
		sHelper.addItem(self.playSoundChk)

		self.contextMenuChk = wx.CheckBox(self, label=_T("set_context_menu"))
		self.contextMenuChk.SetValue(_to_bool(sec.get("contextMenuEnabled", True), default=True))
		sHelper.addItem(self.contextMenuChk)

		# Wybór miejsca odtwarzania dźwięków (otwarcie / lista / oba / wyłączone)
		trigBox = wx.BoxSizer(wx.HORIZONTAL)
		self.lblTriggerMode = wx.StaticText(self, label=_T("set_sound_trigger_mode"))
		trigBox.Add(self.lblTriggerMode, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.triggerKeys = ["open_only", "list_only", "both", "off"]
		self.triggerChoicesPl = [
			_T("sound_mode_open_only"),
			_T("sound_mode_list_only"),
			_T("sound_mode_both"),
			_T("sound_mode_off"),
		]
		self.triggerChoicesEn = [
			"Only when opening a message (default)",
			"Only on the message list",
			"Both on the list and when opening",
			"Disabled",
		]
		currentTrigger = str(sec.get("soundTriggerMode", "open_only"))
		if currentTrigger not in self.triggerKeys:
			currentTrigger = "open_only"
		selTrigIdx = self.triggerKeys.index(currentTrigger)

		trigChoices = self.triggerChoicesEn if effLang == "en" else self.triggerChoicesPl
		self.soundTriggerModeChoice = wx.Choice(self, choices=trigChoices)
		self.soundTriggerModeChoice.SetName(_T("set_sound_trigger_mode"))
		self.soundTriggerModeChoice.SetSelection(selTrigIdx)
		trigBox.Add(self.soundTriggerModeChoice, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(trigBox)

		# Informacja o koordynacji dźwięków z Outlook Speech Filter
		is_sf_handling = isSpeechFilterHandlingInlineImages()
		coord_lbl = _T("set_sound_coord_detected") if is_sf_handling else _T("set_sound_coord_standalone")
		self.soundCoordNote = wx.StaticText(self, label=coord_lbl)
		self.soundCoordNote.Wrap(560)
		sHelper.addItem(self.soundCoordNote)

		self.navSoundsChk = wx.CheckBox(self, label=_T("set_nav_sounds"))
		self.navSoundsChk.SetValue(_to_bool(sec.get("navSounds", True), default=True))
		sHelper.addItem(self.navSoundsChk)

		# Wybór motywu dźwiękowego powiadomienia
		themeBox = wx.BoxSizer(wx.HORIZONTAL)
		self.lblTheme = wx.StaticText(self, label=_T("set_sound_theme"))
		themeBox.Add(self.lblTheme, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.themeKeys = ["associative", "minimal", "instrumental"]
		self.themeChoicesPl = [
			"Skojarzeniowy (migawka aparatu) — Domyślny",
			"Minimalistyczny (podwójny ton elektroniczny)",
			"Instrumentalny (wibrafon / dzwonek)",
		]
		self.themeChoicesEn = [
			"Associative (camera shutter) — Default",
			"Minimalist (double electronic tone)",
			"Instrumental (vibraphone / bell)",
		]
		currentTheme = str(sec.get("soundTheme", "associative"))
		if currentTheme not in self.themeKeys:
			currentTheme = "associative"
		selThemeIdx = self.themeKeys.index(currentTheme)

		tChoices = self.themeChoicesEn if effLang == "en" else self.themeChoicesPl
		self.themeChoice = wx.Choice(self, choices=tChoices)
		self.themeChoice.SetName(_T("set_sound_theme"))
		self.themeChoice.SetSelection(selThemeIdx)
		themeBox.Add(self.themeChoice, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(themeBox)

		# Suwak głośności dźwięków (0-100%) z pełną dostępnością klawiaturową
		volBox = wx.BoxSizer(wx.HORIZONTAL)
		self.lblVolume = wx.StaticText(self, label=_T("set_sound_volume"))
		volBox.Add(self.lblVolume, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		currentVol = sec.get("soundVolume", 80)
		try:
			currentVol = int(currentVol)
		except (ValueError, TypeError):
			currentVol = 80
		currentVol = max(0, min(100, currentVol))

		self.soundVolumeSlider = wx.Slider(
			self,
			value=currentVol,
			minValue=0,
			maxValue=100,
			style=wx.SL_HORIZONTAL,
		)
		self.soundVolumeSlider.SetName(_T("set_sound_volume"))
		self.soundVolumeSlider.SetPageSize(10)
		self.soundVolumeSlider.SetLineSize(1)
		self.soundVolumeSlider.Bind(wx.EVT_SLIDER, self.onSoundVolumeChanged)
		self.soundVolumeSlider.Bind(wx.EVT_KEY_DOWN, self.onSliderKeyDown)
		volBox.Add(self.soundVolumeSlider, proportion=1, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.volValueText = wx.StaticText(self, label=f"{currentVol}%")
		volBox.Add(self.volValueText, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=15)

		self.testSoundBtn = wx.Button(self, label=_T("set_test_volume"))
		def onTestSound(evt):
			val = self.soundVolumeSlider.GetValue()
			sounds = [
				("image_alert", "Powiadomienie o obecności obrazu"),
				("link", "Obraz zawierający link"),
				("divider", "Linia podziału / separator"),
				("social_media", "Ikona mediów społecznościowych"),
				("email", "Ikona adresu e-mail"),
				("phone", "Ikona numeru telefonu"),
			]
			ui.message(f"Odtwarzanie kolejno wszystkich dźwięków dodatku ({len(sounds)} dźwięków, głośność {val}%)...")
			import threading
			def _worker():
				for sName, desc in sounds:
					playAlertSound(sName, volume=val, context=f"Test dźwięku: {desc}", force=True)
					time.sleep(0.48)
			t = threading.Thread(target=_worker, daemon=True)
			t.start()
		self.testSoundBtn.Bind(wx.EVT_BUTTON, onTestSound)
		volBox.Add(self.testSoundBtn, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(volBox)

		# 2. Skróty klawiszowe oraz Instrukcja obsługi
		linksBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnInputGestures = wx.Button(self, label=_T("set_btn_gestures"))
		def onOpenGestures(evt):
			try:
				gui.mainFrame.onInputGesturesCommand(None)
			except Exception as e:
				log.warning(f"outlookImageHelper: onInputGesturesCommand error: {e}")
		self.btnInputGestures.Bind(wx.EVT_BUTTON, onOpenGestures)
		linksBox.Add(self.btnInputGestures, flag=wx.RIGHT, border=10)

		self.btnUserGuide = wx.Button(self, label=_T("set_btn_guide"))
		self.btnUserGuide.Bind(wx.EVT_BUTTON, lambda evt: openUserGuide())
		linksBox.Add(self.btnUserGuide)
		sHelper.addItem(linksBox)

		# 3. Przełącznik trybu zaawansowanego (Basic vs Advanced)
		self.chkShowAdvanced = wx.CheckBox(self, label=_T("set_chk_show_advanced"))
		self.chkShowAdvanced.SetValue(_to_bool(sec.get("showAdvancedSettings", False), default=False))
		self.chkShowAdvanced.Bind(wx.EVT_CHECKBOX, self.onShowAdvancedToggled)
		sHelper.addItem(self.chkShowAdvanced)

		# ======================================================================
		# SEKCJA ZAAWANSOWANA (ukrywana/pokazywana przez chkShowAdvanced)
		# ======================================================================

		# A. Język interfejsu dodatku
		self.langBox = wx.BoxSizer(wx.HORIZONTAL)
		self.lblLang = wx.StaticText(self, label=_T("set_lang_lbl"))
		self.langBox.Add(self.lblLang, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		availLangs = getAvailableLanguages()
		self.langKeys = [item[0] for item in availLangs]
		self.langChoicesPl = [item[1] for item in availLangs]
		self.langChoicesEn = [item[2] for item in availLangs]

		currentLang = str(sec.get("language", "auto")).strip().lower()
		if currentLang not in self.langKeys:
			currentLang = "auto"
		self.initialLanguage = currentLang
		selectedLangIdx = self.langKeys.index(currentLang)

		effLang = getEffectiveLanguage()
		choices = self.langChoicesEn if effLang == "en" else self.langChoicesPl
		self.languageChoice = wx.Choice(self, choices=choices)
		self.languageChoice.SetName(_T("set_lang_lbl"))
		self.languageChoice.SetSelection(selectedLangIdx)
		self.languageChoice.Bind(wx.EVT_CHOICE, self.onLanguageChanged)
		self.langBox.Add(self.languageChoice, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(self.langBox)

		# B. Domyślny język OCR (rozwijana lista ze zlokalizowanymi nazwami!)
		self.ocrBox = wx.BoxSizer(wx.HORIZONTAL)
		self.lblOcr = wx.StaticText(self, label=_T("set_ocr_lbl"))
		self.ocrBox.Add(self.lblOcr, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		self.ocrCodes = [code for code, _ in OCR_LANG_OPTIONS_PL]
		ocrOpts = OCR_LANG_OPTIONS_EN if effLang == "en" else OCR_LANG_OPTIONS_PL
		self.ocrNames = [name for _, name in ocrOpts]

		defaultOcrByLang = "en-US" if effLang == "en" else "pl-PL"
		rawOcr = sec.get("defaultOcrLang", None)
		currentOcr = str(rawOcr).strip() if rawOcr else defaultOcrByLang
		self.initialOcrLang = currentOcr

		if currentOcr in self.ocrCodes and currentOcr != "custom":
			selectedOcrIdx = self.ocrCodes.index(currentOcr)
			customVal = ""
		else:
			selectedOcrIdx = self.ocrCodes.index("custom")
			customVal = currentOcr if currentOcr != "custom" else ""

		self.ocrChoice = wx.Choice(self, choices=self.ocrNames)
		self.ocrChoice.SetName(_T("set_ocr_lbl"))
		self.ocrChoice.SetSelection(selectedOcrIdx)
		self.ocrBox.Add(self.ocrChoice, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(self.ocrBox)

		self.lblCustomOcr = wx.StaticText(self, label=_T("set_custom_ocr_lbl"))
		sHelper.addItem(self.lblCustomOcr)
		self.customOcrEdit = wx.TextCtrl(self)
		self.customOcrEdit.SetName(_T("set_custom_ocr_lbl"))
		self.customOcrEdit.SetValue(str(customVal or ""))
		self.customOcrEdit.Enable(selectedOcrIdx == self.ocrCodes.index("custom"))
		sHelper.addItem(self.customOcrEdit)

		def onOcrChoiceChanged(evt):
			selIdx = self.ocrChoice.GetSelection()
			isCustom = (0 <= selIdx < len(self.ocrCodes) and self.ocrCodes[selIdx] == "custom")
			self.customOcrEdit.Enable(isCustom)
			if isCustom:
				self.customOcrEdit.SetFocus()

		self.ocrChoice.Bind(wx.EVT_CHOICE, onOcrChoiceChanged)

		# C. Raport diagnostyczny i stan dodatku
		self.btnDiagnostics = wx.Button(self, label=_T("set_btn_diagnostics"))
		def onOpenDiagnostics(evt):
			dlg = DiagnosticReportDialog(self)
			dlg.ShowModal()
			dlg.Destroy()
		self.btnDiagnostics.Bind(wx.EVT_BUTTON, onOpenDiagnostics)
		sHelper.addItem(self.btnDiagnostics)

		# D. Prywatność i RODO (AI naming)
		self.aiNamingChk = wx.CheckBox(self, label=_T("set_ai_naming"))
		self.aiNamingChk.SetValue(_to_bool(sec.get("aiNaming", False), default=False))
		sHelper.addItem(self.aiNamingChk)

		# E. Sztuczna inteligencja: Wybór dostawcy (Gemini / OpenAI)
		self.provBox = wx.BoxSizer(wx.HORIZONTAL)
		self.lblAiProvider = wx.StaticText(self, label=_T("set_ai_provider"))
		self.provBox.Add(self.lblAiProvider, flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)

		currentProv = str(sec.get("aiProvider", "gemini")).strip().lower()
		self.aiProviderKeys = ["gemini", "openai"]
		provChoicesPl = [
			"Google Gemini (zalecany, bezpłatny klucz osobisty)",
			"OpenAI / ChatGPT (GPT-4o, GPT-4o-mini)",
		]
		provChoicesEn = [
			"Google Gemini (recommended, free personal key)",
			"OpenAI / ChatGPT (GPT-4o, GPT-4o-mini)",
		]
		pChoices = provChoicesEn if effLang == "en" else provChoicesPl
		selProvIdx = 1 if currentProv == "openai" else 0
		self.aiProviderChoice = wx.Choice(self, choices=pChoices)
		self.aiProviderChoice.SetName(_T("set_ai_provider"))
		self.aiProviderChoice.SetSelection(selProvIdx)
		self.provBox.Add(self.aiProviderChoice, flag=wx.ALIGN_CENTER_VERTICAL)
		sHelper.addItem(self.provBox)

		# F. Sekcja Google Gemini
		self.lblApiKey = wx.StaticText(self, label=_T("set_gemini_key"))
		sHelper.addItem(self.lblApiKey)
		geminiKeyVal = str(sec.get("geminiApiKey", "") or "")
		self.apiKeyEdit = wx.TextCtrl(self, style=wx.TE_PASSWORD)
		self.apiKeyEdit.SetName(_T("set_gemini_key"))
		self.apiKeyEdit.SetValue(geminiKeyVal)
		sHelper.addItem(self.apiKeyEdit)

		self.apiKeyEditPlain = wx.TextCtrl(self)
		self.apiKeyEditPlain.SetName(_T("set_gemini_key"))
		self.apiKeyEditPlain.SetValue(geminiKeyVal)
		self.apiKeyEditPlain.Hide()
		sHelper.addItem(self.apiKeyEditPlain)

		self.chkShowGeminiKey = wx.CheckBox(self, label="Pokaż klucz API Google Gemini" if effLang == "pl" else "Show Google Gemini API key")
		def onToggleGeminiKey(evt):
			if self.chkShowGeminiKey.GetValue():
				self.apiKeyEditPlain.SetValue(self.apiKeyEdit.GetValue())
				self.apiKeyEdit.Hide()
				self.apiKeyEditPlain.Show()
				if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
					self.sHelper.sizer.Show(self.apiKeyEdit, False)
					self.sHelper.sizer.Show(self.apiKeyEditPlain, True)
				self.apiKeyEditPlain.SetFocus()
			else:
				self.apiKeyEdit.SetValue(self.apiKeyEditPlain.GetValue())
				self.apiKeyEditPlain.Hide()
				self.apiKeyEdit.Show()
				if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
					self.sHelper.sizer.Show(self.apiKeyEditPlain, False)
					self.sHelper.sizer.Show(self.apiKeyEdit, True)
				self.apiKeyEdit.SetFocus()
			self.Layout()
		self.chkShowGeminiKey.Bind(wx.EVT_CHECKBOX, onToggleGeminiKey)
		sHelper.addItem(self.chkShowGeminiKey)

		self.apiBtnsBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnGetApiKey = wx.Button(self, label=_T("set_btn_get_api_key"))
		self.btnGetApiKey.Bind(wx.EVT_BUTTON, lambda evt: showGeminiKeyInstructions(self))
		self.apiBtnsBox.Add(self.btnGetApiKey, flag=wx.RIGHT, border=10)

		self.btnTestApiKey = wx.Button(self, label=_T("set_btn_test_api_key"))
		self.btnTestApiKey.Bind(wx.EVT_BUTTON, self.onTestApiKey)
		self.apiBtnsBox.Add(self.btnTestApiKey)
		sHelper.addItem(self.apiBtnsBox)

		self.lblModel = wx.StaticText(self, label=_T("set_gemini_model"))
		sHelper.addItem(self.lblModel)
		self.modelEdit = wx.TextCtrl(self)
		self.modelEdit.SetName(_T("set_gemini_model"))
		curGemModel = str(sec.get("geminiModel", _DEFAULT_GEMINI_MODEL) or _DEFAULT_GEMINI_MODEL).strip()
		if curGemModel in _RETIRED_GEMINI_MODELS:
			curGemModel = _DEFAULT_GEMINI_MODEL
		self.modelEdit.SetValue(curGemModel)
		sHelper.addItem(self.modelEdit)

		# G. Sekcja OpenAI (ChatGPT)
		self.lblOpenAiKey = wx.StaticText(self, label=_T("set_openai_key"))
		sHelper.addItem(self.lblOpenAiKey)
		openaiKeyVal = str(sec.get("openaiApiKey", "") or "")
		self.openAiKeyEdit = wx.TextCtrl(self, style=wx.TE_PASSWORD)
		self.openAiKeyEdit.SetName(_T("set_openai_key"))
		self.openAiKeyEdit.SetValue(openaiKeyVal)
		sHelper.addItem(self.openAiKeyEdit)

		self.openAiKeyEditPlain = wx.TextCtrl(self)
		self.openAiKeyEditPlain.SetName(_T("set_openai_key"))
		self.openAiKeyEditPlain.SetValue(openaiKeyVal)
		self.openAiKeyEditPlain.Hide()
		sHelper.addItem(self.openAiKeyEditPlain)

		self.chkShowOpenAiKey = wx.CheckBox(self, label="Pokaż klucz API OpenAI" if effLang == "pl" else "Show OpenAI API key")
		def onToggleOpenAiKey(evt):
			if self.chkShowOpenAiKey.GetValue():
				self.openAiKeyEditPlain.SetValue(self.openAiKeyEdit.GetValue())
				self.openAiKeyEdit.Hide()
				self.openAiKeyEditPlain.Show()
				if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
					self.sHelper.sizer.Show(self.openAiKeyEdit, False)
					self.sHelper.sizer.Show(self.openAiKeyEditPlain, True)
				self.openAiKeyEditPlain.SetFocus()
			else:
				self.openAiKeyEdit.SetValue(self.openAiKeyEditPlain.GetValue())
				self.openAiKeyEditPlain.Hide()
				self.openAiKeyEdit.Show()
				if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
					self.sHelper.sizer.Show(self.openAiKeyEditPlain, False)
					self.sHelper.sizer.Show(self.openAiKeyEdit, True)
				self.openAiKeyEdit.SetFocus()
			self.Layout()
		self.chkShowOpenAiKey.Bind(wx.EVT_CHECKBOX, onToggleOpenAiKey)
		sHelper.addItem(self.chkShowOpenAiKey)

		self.openAiBtnsBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnGetOpenAiKey = wx.Button(self, label=_T("set_btn_get_openai_key"))
		self.btnGetOpenAiKey.Bind(wx.EVT_BUTTON, lambda evt: showOpenAiKeyInstructions(self))
		self.openAiBtnsBox.Add(self.btnGetOpenAiKey, flag=wx.RIGHT, border=10)

		self.btnTestOpenAiKey = wx.Button(self, label=_T("set_btn_test_openai_key"))
		self.btnTestOpenAiKey.Bind(wx.EVT_BUTTON, self.onTestOpenAiKey)
		self.openAiBtnsBox.Add(self.btnTestOpenAiKey)
		sHelper.addItem(self.openAiBtnsBox)

		self.lblOpenAiModel = wx.StaticText(self, label=_T("set_openai_model"))
		sHelper.addItem(self.lblOpenAiModel)
		self.openAiModelEdit = wx.TextCtrl(self)
		self.openAiModelEdit.SetName(_T("set_openai_model"))
		self.openAiModelEdit.SetValue(str(sec.get("openaiModel", "gpt-4o-mini") or "gpt-4o-mini"))
		sHelper.addItem(self.openAiModelEdit)

		# H. Wspólny prompt
		self.lblPrompt = wx.StaticText(self, label=_T("set_gemini_prompt"))
		sHelper.addItem(self.lblPrompt)
		self.promptEdit = wx.TextCtrl(self, style=wx.TE_MULTILINE, size=(-1, 60))
		self.promptEdit.SetName(_T("set_gemini_prompt"))
		self.promptEdit.SetValue(str(sec.get("aiPrompt", "") or ""))
		sHelper.addItem(self.promptEdit)

		# I. Sekcja aktualizacji dodatku (WCAG 2.1 / 2.2 AA)
		self.chkAutoUpdates = wx.CheckBox(
			self,
			label="&Automatycznie sprawdzaj dostępność aktualizacji przy uruchomieniu NVDA"
		)
		self.chkAutoUpdates.SetValue(_to_bool(sec.get("autoCheckUpdates", True), default=True))
		sHelper.addItem(self.chkAutoUpdates)

		self.btnUpdBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnCheckUpdates = wx.Button(self, label="&Sprawdź dostępność aktualizacji teraz")
		self.btnCheckUpdates.Bind(wx.EVT_BUTTON, self.onCheckUpdatesNow)
		self.btnUpdBox.Add(self.btnCheckUpdates, 0, wx.RIGHT, 10)

		self.btnShareAddon = wx.Button(
			self,
			label="&Podziel się dodatkiem ze znajomymi..." if effLang == "pl" else "&Share add-on with others..."
		)
		self.btnShareAddon.Bind(wx.EVT_BUTTON, lambda evt: showShareAddonDialog(self))
		self.btnUpdBox.Add(self.btnShareAddon, 0)
		sHelper.addItem(self.btnUpdBox)

		# J. Narzędzia twórcy dodatku (NA SAMYM DOLE USTAWIEŃ ZAAWANSOWANYCH)
		self.btnAuthorTools = wx.Button(self, label=_T("set_btn_author_tools"))
		self.btnAuthorTools.Bind(wx.EVT_BUTTON, lambda evt: openCreatorAdminTools(self))
		sHelper.addItem(self.btnAuthorTools)

		# K. Informacje o dodatku (NA SAMYM DOLE CAŁEGO PANELU USTAWIEŃ - ZAWSZE WIDOCZNY)
		self.aboutBox = wx.BoxSizer(wx.HORIZONTAL)
		self.btnAboutAddon = wx.Button(
			self,
			label="&Informacje o dodatku..." if effLang == "pl" else "&About add-on..."
		)
		def onOpenAbout(evt):
			dlg = AddonAboutDialog(self)
			gui.mainFrame.prePopup()
			dlg.ShowModal()
			dlg.Destroy()
			gui.mainFrame.postPopup()
		self.btnAboutAddon.Bind(wx.EVT_BUTTON, onOpenAbout)
		self.aboutBox.Add(self.btnAboutAddon, 0)
		sHelper.addItem(self.aboutBox)

		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

		# Rejestracja formantów trybu zaawansowanego i wstępne ustawienie widoczności
		self.advancedControls = [
			self.langBox,
			self.lblLang,
			self.languageChoice,
			self.ocrBox,
			self.lblOcr,
			self.ocrChoice,
			self.lblCustomOcr,
			self.customOcrEdit,
			self.btnDiagnostics,
			self.aiNamingChk,
			self.provBox,
			self.lblAiProvider,
			self.aiProviderChoice,
			self.lblApiKey,
			self.apiKeyEdit,
			self.apiKeyEditPlain,
			self.chkShowGeminiKey,
			self.apiBtnsBox,
			self.lblModel,
			self.modelEdit,
			self.lblOpenAiKey,
			self.openAiKeyEdit,
			self.openAiKeyEditPlain,
			self.chkShowOpenAiKey,
			self.openAiBtnsBox,
			self.lblOpenAiModel,
			self.openAiModelEdit,
			self.lblPrompt,
			self.promptEdit,
			self.chkAutoUpdates,
			self.btnUpdBox,
			self.btnCheckUpdates,
			self.btnShareAddon,
			self.btnAuthorTools,
		]
		self.onShowAdvancedToggled()

	def onShowAdvancedToggled(self, evt=None):
		"""Przełącza widoczność opcji zaawansowanych i dynamicznie dopasowuje obszar przewijania."""
		show = self.chkShowAdvanced.GetValue()
		for ctrl in getattr(self, "advancedControls", []):
			try:
				if ctrl in (getattr(self, "apiKeyEditPlain", None), getattr(self, "apiKeyEdit", None)):
					continue
				if ctrl in (getattr(self, "openAiKeyEditPlain", None), getattr(self, "openAiKeyEdit", None)):
					continue
				if isinstance(ctrl, wx.Sizer):
					ctrl.ShowItems(show)
					if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
						self.sHelper.sizer.Show(ctrl, show)
				else:
					ctrl.Show(show)
					if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
						self.sHelper.sizer.Show(ctrl, show)
			except Exception as e:
				log.debug(f"outlookImageHelper: onShowAdvancedToggled ctrl show error: {e}")

		# Prawidłowa obsługa widoczności pól kluczy (hasło vs jawny) w trybie zaawansowanym
		if hasattr(self, "apiKeyEdit") and hasattr(self, "apiKeyEditPlain") and hasattr(self, "chkShowGeminiKey"):
			showPlain = self.chkShowGeminiKey.GetValue() if show else False
			showPass = (not self.chkShowGeminiKey.GetValue()) if show else False
			self.apiKeyEdit.Show(showPass)
			self.apiKeyEditPlain.Show(showPlain)
			if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
				self.sHelper.sizer.Show(self.apiKeyEdit, showPass)
				self.sHelper.sizer.Show(self.apiKeyEditPlain, showPlain)

		if hasattr(self, "openAiKeyEdit") and hasattr(self, "openAiKeyEditPlain") and hasattr(self, "chkShowOpenAiKey"):
			showPlain = self.chkShowOpenAiKey.GetValue() if show else False
			showPass = (not self.chkShowOpenAiKey.GetValue()) if show else False
			self.openAiKeyEdit.Show(showPass)
			self.openAiKeyEditPlain.Show(showPlain)
			if hasattr(self, "sHelper") and hasattr(self.sHelper, "sizer"):
				self.sHelper.sizer.Show(self.openAiKeyEdit, showPass)
				self.sHelper.sizer.Show(self.openAiKeyEditPlain, showPlain)

		self.Layout()
		p = self
		while p:
			try:
				if hasattr(p, "FitInside"):
					p.FitInside()
				if hasattr(p, "AdjustScrollbars"):
					p.AdjustScrollbars()
				p.Layout()
			except Exception:
				pass
			p = p.GetParent()

	def updateUILabels(self):
		"""Aktualizuje wszystkie etykiety kontrolek w panelu natychmiast po zmianie języka."""
		try:
			eff = getEffectiveLanguage()
			if hasattr(self, "lblLang") and self.lblLang:
				self.lblLang.SetLabel(_T("set_lang_lbl"))
			if hasattr(self, "languageChoice") and self.languageChoice:
				self.languageChoice.SetName(_T("set_lang_lbl"))
			if hasattr(self, "lblOcr") and self.lblOcr:
				self.lblOcr.SetLabel(_T("set_ocr_lbl"))
			if hasattr(self, "customOcrEdit") and self.customOcrEdit:
				self.customOcrEdit.SetName(_T("set_custom_ocr_lbl"))
			if hasattr(self, "lblCustomOcr") and self.lblCustomOcr:
				self.lblCustomOcr.SetLabel(_T("set_custom_ocr_lbl"))

			# Aktualizacja listy języków OCR w ocrChoice
			if hasattr(self, "ocrChoice") and self.ocrChoice:
				self.ocrChoice.SetName(_T("set_ocr_lbl"))
				curOcrIdx = self.ocrChoice.GetSelection()
				curOcrCode = self.ocrCodes[curOcrIdx] if (hasattr(self, "ocrCodes") and 0 <= curOcrIdx < len(self.ocrCodes)) else None
				ocrOpts = OCR_LANG_OPTIONS_EN if eff == "en" else OCR_LANG_OPTIONS_PL
				for i, (_, name) in enumerate(ocrOpts):
					try:
						self.ocrChoice.SetString(i, name)
					except Exception:
						pass
				if curOcrCode and hasattr(self, "ocrCodes") and curOcrCode in self.ocrCodes:
					self.ocrChoice.SetSelection(self.ocrCodes.index(curOcrCode))

			if hasattr(self, "autoAnnounceChk") and self.autoAnnounceChk:
				self.autoAnnounceChk.SetLabel(_T("set_auto_announce"))
			if hasattr(self, "playSoundChk") and self.playSoundChk:
				self.playSoundChk.SetLabel(_T("set_play_sound"))
			if hasattr(self, "contextMenuChk") and self.contextMenuChk:
				self.contextMenuChk.SetLabel(_T("set_context_menu"))
			if hasattr(self, "navSoundsChk") and self.navSoundsChk:
				self.navSoundsChk.SetLabel(_T("set_nav_sounds"))
			if hasattr(self, "lblTheme") and self.lblTheme:
				self.lblTheme.SetLabel(_T("set_sound_theme"))
			if hasattr(self, "themeChoice") and self.themeChoice:
				self.themeChoice.SetName(_T("set_sound_theme"))
				themeIdx = self.themeChoice.GetSelection()
				tChoices = self.themeChoicesEn if eff == "en" else self.themeChoicesPl
				for i, ch in enumerate(tChoices):
					try:
						self.themeChoice.SetString(i, ch)
					except Exception:
						pass
				if 0 <= themeIdx < len(tChoices):
					self.themeChoice.SetSelection(themeIdx)
			if hasattr(self, "lblTriggerMode") and self.lblTriggerMode:
				self.lblTriggerMode.SetLabel(_T("set_sound_trigger_mode"))
			if hasattr(self, "soundTriggerModeChoice") and self.soundTriggerModeChoice:
				self.soundTriggerModeChoice.SetName(_T("set_sound_trigger_mode"))
				trigIdx = self.soundTriggerModeChoice.GetSelection()
				choices = self.triggerChoicesEn if eff == "en" else self.triggerChoicesPl
				for i, ch in enumerate(choices):
					try:
						self.soundTriggerModeChoice.SetString(i, ch)
					except Exception:
						pass
				if 0 <= trigIdx < len(choices):
					self.soundTriggerModeChoice.SetSelection(trigIdx)
			if hasattr(self, "soundCoordNote") and self.soundCoordNote:
				is_sf = isSpeechFilterHandlingInlineImages()
				self.soundCoordNote.SetLabel(_T("set_sound_coord_detected") if is_sf else _T("set_sound_coord_standalone"))
			if hasattr(self, "lblVolume") and self.lblVolume:
				self.lblVolume.SetLabel(_T("set_sound_volume"))
			if hasattr(self, "soundVolumeSlider") and self.soundVolumeSlider:
				self.soundVolumeSlider.SetName(_T("set_sound_volume"))
			if hasattr(self, "testSoundBtn") and self.testSoundBtn:
				self.testSoundBtn.SetLabel(_T("set_test_volume"))
			if hasattr(self, "btnDiagnostics") and self.btnDiagnostics:
				self.btnDiagnostics.SetLabel(_T("set_btn_diagnostics"))
			if hasattr(self, "aiNamingChk") and self.aiNamingChk:
				self.aiNamingChk.SetLabel(_T("set_ai_naming"))
			if hasattr(self, "lblAiProvider") and self.lblAiProvider:
				self.lblAiProvider.SetLabel(_T("set_ai_provider"))
			if hasattr(self, "aiProviderChoice") and self.aiProviderChoice:
				self.aiProviderChoice.SetName(_T("set_ai_provider"))
			if hasattr(self, "lblApiKey") and self.lblApiKey:
				self.lblApiKey.SetLabel(_T("set_gemini_key"))
			if hasattr(self, "apiKeyEdit") and self.apiKeyEdit:
				self.apiKeyEdit.SetName(_T("set_gemini_key"))
			if hasattr(self, "btnGetApiKey") and self.btnGetApiKey:
				self.btnGetApiKey.SetLabel(_T("set_btn_get_api_key"))
			if hasattr(self, "btnTestApiKey") and self.btnTestApiKey:
				self.btnTestApiKey.SetLabel(_T("set_btn_test_api_key"))
			if hasattr(self, "lblModel") and self.lblModel:
				self.lblModel.SetLabel(_T("set_gemini_model"))
			if hasattr(self, "modelEdit") and self.modelEdit:
				self.modelEdit.SetName(_T("set_gemini_model"))
			if hasattr(self, "lblOpenAiKey") and self.lblOpenAiKey:
				self.lblOpenAiKey.SetLabel(_T("set_openai_key"))
			if hasattr(self, "openAiKeyEdit") and self.openAiKeyEdit:
				self.openAiKeyEdit.SetName(_T("set_openai_key"))
			if hasattr(self, "btnGetOpenAiKey") and self.btnGetOpenAiKey:
				self.btnGetOpenAiKey.SetLabel(_T("set_btn_get_openai_key"))
			if hasattr(self, "btnTestOpenAiKey") and self.btnTestOpenAiKey:
				self.btnTestOpenAiKey.SetLabel(_T("set_btn_test_openai_key"))
			if hasattr(self, "lblOpenAiModel") and self.lblOpenAiModel:
				self.lblOpenAiModel.SetLabel(_T("set_openai_model"))
			if hasattr(self, "openAiModelEdit") and self.openAiModelEdit:
				self.openAiModelEdit.SetName(_T("set_openai_model"))
			if hasattr(self, "lblPrompt") and self.lblPrompt:
				self.lblPrompt.SetLabel(_T("set_gemini_prompt"))
			if hasattr(self, "promptEdit") and self.promptEdit:
				self.promptEdit.SetName(_T("set_gemini_prompt"))
			if hasattr(self, "btnInputGestures") and self.btnInputGestures:
				self.btnInputGestures.SetLabel(_T("set_btn_gestures"))
			if hasattr(self, "btnUserGuide") and self.btnUserGuide:
				self.btnUserGuide.SetLabel(_T("set_btn_guide"))
			if hasattr(self, "btnDiagnostics") and self.btnDiagnostics:
				self.btnDiagnostics.SetLabel(_T("set_btn_diagnostics"))
			if hasattr(self, "chkShowAdvanced") and self.chkShowAdvanced:
				self.chkShowAdvanced.SetLabel(_T("set_chk_show_advanced"))
			if hasattr(self, "btnAuthorTools") and self.btnAuthorTools:
				self.btnAuthorTools.SetLabel(_T("set_btn_author_tools"))
			if hasattr(self, "btnAboutAddon") and self.btnAboutAddon:
				self.btnAboutAddon.SetLabel("&Informacje o dodatku..." if eff == "pl" else "&About add-on...")
			if hasattr(self, "chkShowGeminiKey") and self.chkShowGeminiKey:
				self.chkShowGeminiKey.SetLabel("Pokaż klucz API Google Gemini" if eff == "pl" else "Show Google Gemini API key")
			if hasattr(self, "chkShowOpenAiKey") and self.chkShowOpenAiKey:
				self.chkShowOpenAiKey.SetLabel("Pokaż klucz API OpenAI" if eff == "pl" else "Show OpenAI API key")

			# Zaktualizuj etykiety w languageChoice
			if hasattr(self, "languageChoice") and self.languageChoice:
				langIdx = self.languageChoice.GetSelection()
				lChoices = self.langChoicesEn if eff == "en" else self.langChoicesPl
				for i, ch in enumerate(lChoices):
					try:
						self.languageChoice.SetString(i, ch)
					except Exception:
						pass
				if 0 <= langIdx < len(lChoices):
					self.languageChoice.SetSelection(langIdx)

			# Bezpieczna aktualizacja nazwy kategorii w oknie ustawień NVDA
			try:
				parentDlg = self.GetParent()
				while parentDlg and not isinstance(parentDlg, gui.settingsDialogs.NVDASettingsDialog):
					parentDlg = parentDlg.GetParent()

				def _isMatch(t: str) -> bool:
					if not t:
						return False
					tl = t.lower()
					if "filtr" in tl or "filter" in tl or "speech" in tl or "mowy" in tl:
						return False
					keywords = [
						"pomocnik", "obraz", "image", "attach", "załącznik", "bild", "anhang",
						"adjunto", "pièce", "allegat"
					]
					return any(kw in tl for kw in keywords)

				newTitle = _T("addon_title")
				if parentDlg and hasattr(parentDlg, "catListCtrl"):
					curSel = parentDlg.catListCtrl.GetFirstSelected()
					if curSel != -1:
						parentDlg.catListCtrl.SetItemText(curSel, newTitle)
					else:
						for i in range(parentDlg.catListCtrl.GetItemCount()):
							text = parentDlg.catListCtrl.GetItemText(i)
							if _isMatch(text):
								parentDlg.catListCtrl.SetItemText(i, newTitle)
								break
				elif parentDlg and hasattr(parentDlg, "categoriesList"):
					curSel = parentDlg.categoriesList.GetSelection()
					if curSel != -1:
						parentDlg.categoriesList.SetString(curSel, newTitle)
					else:
						for i in range(parentDlg.categoriesList.GetCount()):
							text = parentDlg.categoriesList.GetString(i)
							if _isMatch(text):
								parentDlg.categoriesList.SetString(i, newTitle)
								break
			except Exception:
				pass

			self.Layout()
		except Exception as e:
			log.warning(f"outlookImageHelper: updateUILabels error: {e}")

	def onLanguageChanged(self, evt):
		idx = self.languageChoice.GetSelection()
		if 0 <= idx < len(self.langKeys):
			newKey = self.langKeys[idx]
			config.conf[CONFIG_SECTION]["language"] = newKey
			eff = getEffectiveLanguage()

			# Synchronizuj domyślny język OCR przy zmianie języka interfejsu dodatku
			curOcrIdx = self.ocrChoice.GetSelection() if hasattr(self, "ocrChoice") and self.ocrChoice else -1
			curOcrCode = self.ocrCodes[curOcrIdx] if (hasattr(self, "ocrCodes") and 0 <= curOcrIdx < len(self.ocrCodes)) else "pl-PL"
			targetOcr = "pl-PL" if eff == "pl" else "en-US"

			if curOcrCode in ("pl-PL", "en-US") or (hasattr(self, "ocrCodes") and curOcrCode not in self.ocrCodes):
				if hasattr(self, "ocrCodes") and targetOcr in self.ocrCodes:
					newOcrIdx = self.ocrCodes.index(targetOcr)
					self.ocrChoice.SetSelection(newOcrIdx)
					config.conf[CONFIG_SECTION]["defaultOcrLang"] = targetOcr
					if hasattr(self, "customOcrEdit") and self.customOcrEdit:
						self.customOcrEdit.Enable(False)

			self.updateUILabels()
			chosenName = self.languageChoice.GetString(idx)
			isPl = (eff == "pl")
			msg = f"Zmieniono język interfejsu dodatku na: {chosenName}" if isPl else f"Add-on interface language changed to: {chosenName}"
			ui.message(msg)

	def refreshLanguageChoices(self, selectCode: str = None):
		"""Odświeża listę języków w languageChoice (np. po wygenerowaniu lub zaimportowaniu nowego języka)."""
		avail = getAvailableLanguages()
		self.langKeys = [item[0] for item in avail]
		self.langChoicesPl = [item[1] for item in avail]
		self.langChoicesEn = [item[2] for item in avail]
		eff = getEffectiveLanguage()
		choices = self.langChoicesEn if eff == "en" else self.langChoicesPl
		if hasattr(self, "languageChoice") and self.languageChoice:
			self.languageChoice.Clear()
			for ch in choices:
				self.languageChoice.Append(ch)
			if selectCode and selectCode in self.langKeys:
				selIdx = self.langKeys.index(selectCode)
			else:
				curLang = config.conf[CONFIG_SECTION].get("language", "auto")
				selIdx = self.langKeys.index(curLang) if curLang in self.langKeys else 0
			self.languageChoice.SetSelection(selIdx)

	def onTranslateAi(self, evt):
		def onGen(code):
			self.refreshLanguageChoices(selectCode=code)
			config.conf[CONFIG_SECTION]["language"] = code
			self.updateUILabels()
		dlg = GenerateLanguageDialog(self, onGeneratedCallback=onGen)
		dlg.ShowModal()

	def onExportLang(self, evt):
		with wx.FileDialog(
			self,
			message=_T("dlg_export_title"),
			wildcard="JSON files (*.json)|*.json",
			style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
			defaultFile="outlookImageHelper_language_template.json",
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				path = dlg.GetPath()
				if exportLanguageTemplate(path):
					ui.message(_T("dlg_export_success", file=os.path.basename(path)))
				else:
					ui.message("Nie udało się wyeksportować szablonu.")

	def onImportLang(self, evt):
		with wx.FileDialog(
			self,
			message=_T("dlg_import_title"),
			wildcard="JSON files (*.json)|*.json",
			style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				path = dlg.GetPath()
				ok, code, name = importLanguageFile(path)
				if ok:
					self.refreshLanguageChoices(selectCode=code)
					config.conf[CONFIG_SECTION]["language"] = code
					self.updateUILabels()
					ui.message(_T("dlg_import_success", name=name, code=code))
				else:
					ui.message(_T("dlg_import_err_invalid", err=name))

	def onGetApiKeyWizard(self, evt):
		def onKeySaved(keyVal=None):
			savedKey = (keyVal or config.conf[CONFIG_SECTION].get("geminiApiKey", "")).strip()
			if hasattr(self, "apiKeyEdit") and self.apiKeyEdit:
				self.apiKeyEdit.SetValue(savedKey)
		dlg = GeminiApiKeyWizardDialog(self, onKeySavedCallback=onKeySaved)
		dlg.ShowModal()

	def onTestApiKey(self, evt):
		key = (self.apiKeyEditPlain.GetValue() if hasattr(self, "chkShowGeminiKey") and self.chkShowGeminiKey.GetValue() else self.apiKeyEdit.GetValue()).strip() if hasattr(self, "apiKeyEdit") and self.apiKeyEdit else ""
		if not key:
			ui.message(_T("wizard_key_empty"))
			return
		testMsg = _T("wizard_testing") or "Testowanie połączenia z Google Gemini, proszę czekać..."
		progress = AiProgressFeedback(
			startMessage=testMsg,
			ongoingMessage="Nadal trwa testowanie klucza API Google Gemini, proszę czekać..."
		)
		if hasattr(self, "btnTestApiKey") and self.btnTestApiKey:
			self.btnTestApiKey.Disable()
		model = self.modelEdit.GetValue().strip() if hasattr(self, "modelEdit") and self.modelEdit else _DEFAULT_GEMINI_MODEL

		def worker():
			ok = False
			msg = ""
			try:
				ok, msg = testGeminiApiKey(key, model=model)
			except Exception as ex:
				msg = f"Błąd testowania: {ex}"
			finally:
				def finish():
					if hasattr(self, "btnTestApiKey") and self.btnTestApiKey:
						self.btnTestApiKey.Enable()
					progress.stop(finishMessage=msg)
					try:
						if ok:
							tones.beep(880, 50)
							tones.beep(1175, 70)
						else:
							tones.beep(330, 150)
					except Exception:
						pass
				wx.CallAfter(finish)

		threading.Thread(target=worker, daemon=True).start()

	def onTestOpenAiKey(self, evt):
		key = (self.openAiKeyEditPlain.GetValue() if hasattr(self, "chkShowOpenAiKey") and self.chkShowOpenAiKey.GetValue() else self.openAiKeyEdit.GetValue()).strip() if hasattr(self, "openAiKeyEdit") and self.openAiKeyEdit else ""
		if not key:
			ui.message("Pole klucza API OpenAI jest puste.")
			return
		progress = AiProgressFeedback(
			startMessage="Testowanie połączenia z OpenAI, proszę czekać...",
			ongoingMessage="Nadal trwa testowanie klucza API OpenAI, proszę czekać..."
		)
		if hasattr(self, "btnTestOpenAiKey") and self.btnTestOpenAiKey:
			self.btnTestOpenAiKey.Disable()
		model = self.openAiModelEdit.GetValue().strip() if hasattr(self, "openAiModelEdit") and self.openAiModelEdit else "gpt-4o-mini"

		def worker():
			ok = False
			msg = ""
			try:
				ok, msg = testOpenAiApiKey(key, model=model)
			except Exception as ex:
				msg = f"Błąd testowania: {ex}"
			finally:
				def finish():
					if hasattr(self, "btnTestOpenAiKey") and self.btnTestOpenAiKey:
						self.btnTestOpenAiKey.Enable()
					progress.stop(finishMessage=msg)
					try:
						if ok:
							tones.beep(880, 50)
							tones.beep(1175, 70)
						else:
							tones.beep(330, 150)
					except Exception:
						pass
				wx.CallAfter(finish)

		threading.Thread(target=worker, daemon=True).start()

	def _updateSliderValue(self, newVal: int):
		newVal = max(0, min(100, newVal))
		self.soundVolumeSlider.SetValue(newVal)
		self.volValueText.SetLabel(f"{newVal}%")
		try:
			config.conf[CONFIG_SECTION]["soundVolume"] = newVal
		except Exception:
			pass
		speech.speakMessage(f"{newVal}%")

	def onSliderKeyDown(self, evt):
		keyCode = evt.GetKeyCode()
		currentVal = self.soundVolumeSlider.GetValue()
		if keyCode in (wx.WXK_PAGEUP, getattr(wx, "WXK_NUMPAD_PAGEUP", -1)):
			self._updateSliderValue(currentVal + 10)
			return
		elif keyCode in (wx.WXK_PAGEDOWN, getattr(wx, "WXK_NUMPAD_PAGEDOWN", -1)):
			self._updateSliderValue(currentVal - 10)
			return
		elif keyCode in (wx.WXK_UP, wx.WXK_RIGHT, getattr(wx, "WXK_NUMPAD_UP", -1), getattr(wx, "WXK_NUMPAD_RIGHT", -1)):
			self._updateSliderValue(currentVal + 1)
			return
		elif keyCode in (wx.WXK_DOWN, wx.WXK_LEFT, getattr(wx, "WXK_NUMPAD_DOWN", -1), getattr(wx, "WXK_NUMPAD_LEFT", -1)):
			self._updateSliderValue(currentVal - 1)
			return
		elif keyCode in (wx.WXK_HOME, getattr(wx, "WXK_NUMPAD_HOME", -1)):
			self._updateSliderValue(100)
			return
		elif keyCode in (wx.WXK_END, getattr(wx, "WXK_NUMPAD_END", -1)):
			self._updateSliderValue(0)
			return
		evt.Skip()

	def onSoundVolumeChanged(self, evt):
		val = self.soundVolumeSlider.GetValue()
		self.volValueText.SetLabel(f"{val}%")
		try:
			config.conf[CONFIG_SECTION]["soundVolume"] = int(val)
		except Exception:
			pass

	def onCharHook(self, event):
		if event.GetKeyCode() == wx.WXK_F1:
			showContextHelp("settings", parent=self)
			return
		if event.ControlDown() and event.ShiftDown() and event.GetKeyCode() in (ord('T'), ord('t')):
			openCreatorAdminTools(self)
			return
		event.Skip()

	def onSave(self):
		sec = config.conf[CONFIG_SECTION]

		# Język interfejsu
		idx = self.languageChoice.GetSelection()
		if 0 <= idx < len(self.langKeys):
			sec["language"] = self.langKeys[idx]
			self.initialLanguage = self.langKeys[idx]

		self.updateUILabels()

		# Język OCR
		ocrIdx = self.ocrChoice.GetSelection()
		if hasattr(self, "ocrCodes") and 0 <= ocrIdx < len(self.ocrCodes):
			selectedCode = self.ocrCodes[ocrIdx]
			if selectedCode == "custom":
				customCode = self.customOcrEdit.GetValue().strip()
				sec["defaultOcrLang"] = customCode if customCode else "pl-PL"
			else:
				sec["defaultOcrLang"] = selectedCode
			self.initialOcrLang = sec["defaultOcrLang"]

		sec["autoAnnounce"] = self.autoAnnounceChk.GetValue()
		sec["playSound"] = self.playSoundChk.GetValue()
		if hasattr(self, "contextMenuChk") and self.contextMenuChk:
			sec["contextMenuEnabled"] = self.contextMenuChk.GetValue()
		if hasattr(self, "soundTriggerModeChoice") and self.soundTriggerModeChoice:
			trigIdx = self.soundTriggerModeChoice.GetSelection()
			if 0 <= trigIdx < len(self.triggerKeys):
				sec["soundTriggerMode"] = self.triggerKeys[trigIdx]
		if hasattr(self, "soundVolumeSlider") and self.soundVolumeSlider:
			sec["soundVolume"] = int(self.soundVolumeSlider.GetValue())
		sec["navSounds"] = self.navSoundsChk.GetValue()
		themeIdx = self.themeChoice.GetSelection()
		if 0 <= themeIdx < len(self.themeKeys):
			sec["soundTheme"] = self.themeKeys[themeIdx]
		sec["aiNaming"] = self.aiNamingChk.GetValue()
		if hasattr(self, "aiProviderChoice") and self.aiProviderChoice:
			provIdx = self.aiProviderChoice.GetSelection()
			sec["aiProvider"] = "openai" if provIdx == 1 else "gemini"
		if hasattr(self, "apiKeyEdit") and self.apiKeyEdit:
			val = self.apiKeyEditPlain.GetValue() if hasattr(self, "chkShowGeminiKey") and self.chkShowGeminiKey.GetValue() else self.apiKeyEdit.GetValue()
			sec["geminiApiKey"] = val.strip()
		if hasattr(self, "modelEdit") and self.modelEdit:
			sec["geminiModel"] = self.modelEdit.GetValue().strip() or _DEFAULT_GEMINI_MODEL
		if hasattr(self, "openAiKeyEdit") and self.openAiKeyEdit:
			val = self.openAiKeyEditPlain.GetValue() if hasattr(self, "chkShowOpenAiKey") and self.chkShowOpenAiKey.GetValue() else self.openAiKeyEdit.GetValue()
			sec["openaiApiKey"] = val.strip()
		if hasattr(self, "openAiModelEdit") and self.openAiModelEdit:
			sec["openaiModel"] = self.openAiModelEdit.GetValue().strip() or "gpt-4o-mini"
		if hasattr(self, "promptEdit") and self.promptEdit:
			sec["aiPrompt"] = self.promptEdit.GetValue().strip()
		if hasattr(self, "chkShowAdvanced") and self.chkShowAdvanced:
			sec["showAdvancedSettings"] = self.chkShowAdvanced.GetValue()
		if hasattr(self, "chkAutoUpdates") and self.chkAutoUpdates:
			sec["autoCheckUpdates"] = bool(self.chkAutoUpdates.GetValue())
		if hasattr(self, "txtUpdateSource") and self.txtUpdateSource:
			sec["updateSourceUrl"] = self.txtUpdateSource.GetValue().strip()

	def onCheckUpdatesNow(self, evt):
		updater = AddonUpdaterEngine(
			addon_name=CONFIG_SECTION,
			addon_title=ADDON_TITLE,
			current_version=ADDON_VERSION,
			addon_dir=ADDON_DIR,
			config_section=CONFIG_SECTION
		)
		updater.check_for_updates(manual=True)

	def onDiscard(self):
		global _activeSettingsPanelRef
		_activeSettingsPanelRef = None
		try:
			sec = config.conf[CONFIG_SECTION]
			if hasattr(self, "initialLanguage"):
				sec["language"] = self.initialLanguage
			if hasattr(self, "initialOcrLang"):
				sec["defaultOcrLang"] = self.initialOcrLang
		except Exception:
			pass


# ==============================================================================
# GŁÓWNA KLASA GLOBALPLUGIN
# ==============================================================================


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = SettingsPanelTitleDescriptor()
	instance: GlobalPlugin = None

	__gestures__ = {
		# Lista grafik w bieżącej wiadomości
		"kb:NVDA+shift+g": "openGraphicList",
		"kb:alt+shift+g": "openGraphicList",
		"kb:NVDA+control+a": "openGraphicList",
		# Menu czynności dla grafiki pod kursorem w wiadomości Outlook
		"kb:alt+c": "graphicActionsMenu",
		"kb:alt+shift+c": "graphicActionsMenu",
		"kb:NVDA+alt+c": "graphicActionsMenu",
		"kb:apps": "graphicAppsMenu",
		"kb:shift+f10": "graphicAppsMenu",
		# Szybki OCR grafiki pod kursorem lub jedynej w mailu
		"kb:NVDA+windows+g": "quickOcr",
		"kb:NVDA+alt+g": "quickOcr",
		"kb:alt+shift+o": "quickOcr",
		# Szybki opis obrazu przez AI
		"kb:NVDA+windows+shift+g": "quickAi",
		"kb:NVDA+alt+shift+g": "quickAi",
		"kb:alt+shift+a": "quickAi",
		# Szybki podgląd w negatywie (inwersja kolorów dla słabowidzących)
		"kb:NVDA+windows+i": "quickInvert",
		"kb:NVDA+alt+i": "quickInvert",
		"kb:alt+shift+i": "quickInvert",
	}

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		GlobalPlugin.instance = self
		self._bypassNextContextMenu = False
		self._lastAnnouncedHwnd = 0
		self._currentGraphicItem = None
		self._lastNavigatedDocKey = None
		self._lastNavigatedIndex = -1
		self._lastNavigatedCaretPos = -9999
		self._suppressCaretSpeechUntil = 0.0
		self._cachedGraphicsDocKey = None
		self._cachedGraphicsItems = []
		self._cachedGraphicsTime = 0.0
		self._cachedGraphicsFast = True
		self.toolsMenuItem = None
		self.toolsGuideItem = None

		config.conf.spec[CONFIG_SECTION] = confspec
		if CONFIG_SECTION not in config.conf:
			config.conf[CONFIG_SECTION] = {}

		# Ścisła sanitizacja i konwersja typów kluczy w sekcji konfiguracji (ochrona przed błędem CheckBox.SetValue ze str)
		try:
			sec = config.conf[CONFIG_SECTION]
			for boolKey in ("autoAnnounce", "playSound", "navSounds", "aiNaming", "showAdvancedSettings"):
				if boolKey in sec:
					sec[boolKey] = _to_bool(sec[boolKey], default=(boolKey not in ("aiNaming", "showAdvancedSettings")))
			if "minDimension" in sec:
				sec["minDimension"] = _to_int(sec["minDimension"], default=15, min_val=1, max_val=500)
			if str(sec.get("geminiModel", "")).strip() in _RETIRED_GEMINI_MODELS:
				sec["geminiModel"] = _DEFAULT_GEMINI_MODEL
			# Migracja starej lokalizacji pliku aktualizacji na oficjalne repozytorium GitHub
			curUpd = str(sec.get("updateSourceUrl", "")).strip()
			isOldDefault = bool(curUpd and ("mój dysk" in curUpd.lower() and "antigravity" in curUpd.lower()))
			if not curUpd or isOldDefault:
				sec["updateSourceUrl"] = "https://raw.githubusercontent.com/strugo10/outlookImageHelper/main/update_info.json"
		except Exception as e:
			log.warning(f"outlookImageHelper: Błąd sanitizacji konfiguracji: {e}")

		try:
			config.conf.validate(config.conf.validator, copy=True)
		except Exception:
			pass

		# Bezpieczna rejestracja panelu ustawień bez duplikatów
		if OutlookImageHelperSettingsPanel not in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(OutlookImageHelperSettingsPanel)

		# Podłącz zoptymalizowaną nawigację G i Shift+G do buforów przeglądania wiadomości
		try:
			_patchOutlookTreeInterceptors()
		except Exception as ex:
			log.warning(f"outlookImageHelper: Błąd _patchOutlookTreeInterceptors: {ex}")

		# Jawne powiązanie gestów klawiszowych dla natychmiastowej responsywności
		for gestureStr, scriptName in self.__gestures__.items():
			try:
				self.bindGesture(gestureStr, scriptName)
			except Exception as ex:
				log.debugWarning(f"outlookImageHelper: Nie udało się powiązać gestu {gestureStr}: {ex}")

		# Inicjalizacja modułu aktualizacji i opóźnione sprawdzenie w tle
		self._updater = AddonUpdaterEngine(
			addon_name=CONFIG_SECTION,
			addon_title=ADDON_TITLE,
			current_version=ADDON_VERSION,
			addon_dir=ADDON_DIR,
			config_section=CONFIG_SECTION
		)
		wx.CallAfter(self._init_tools_menu_and_cleanup)
		wx.CallLater(15000, self._check_updates_auto)

	def _init_tools_menu_and_cleanup(self):
		# Dodanie skrótów do menu Narzędzia w menu głównym NVDA
		try:
			self.toolsMenuItem = gui.mainFrame.sysTrayIcon.toolsMenu.Append(
				wx.ID_ANY,
				"&Outlook - Pomocnik obrazów i OCR (Lista grafik)...\tCtrl+NVDA+A",
				"Otwiera listę grafik i zrzutów ekranu w otwartej wiadomości programu Outlook",
			)
			gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self.onToolsMenuOpen, self.toolsMenuItem)

			self.toolsGuideItem = gui.mainFrame.sysTrayIcon.toolsMenu.Append(
				wx.ID_ANY,
				"&Instrukcja obsługi dodatku Outlook Pomocnik obrazów i OCR...",
				"Otwiera pełną dokumentację i instrukcję obsługi dodatku w przeglądarce",
			)
			gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, lambda evt: openUserGuide(), self.toolsGuideItem)
		except Exception as e:
			log.warning(f"outlookImageHelper: Nie udało się dodać pozycji do menu Narzędzia: {e}")
		log.info("outlookImageHelper: Dodatek załadowany pomyślnie. Zarejestrowano panel ustawień, menu Narzędzia, gesty oraz obsługę G/Shift+G.")

		# Uruchomienie czyszczenia starych plików tymczasowych w tle
		try:
			threading.Thread(target=cleanOldTempFiles, daemon=True, name="oih_temp_cleanup").start()
		except Exception:
			pass

	def _check_updates_auto(self):
		conf = config.conf.get(CONFIG_SECTION, {})
		if not conf.get("autoCheckUpdates", True):
			return
		last_time = conf.get("lastUpdateCheckTime", 0.0)
		if time.time() - last_time < 86400:
			return
		self._updater.check_for_updates(manual=False)

	def terminate(self):
		super().terminate()
		try:
			_unpatchOutlookTreeInterceptors()
		except Exception as ex:
			log.debug(f"outlookImageHelper: _unpatchOutlookTreeInterceptors error: {ex}")
		try:
			while OutlookImageHelperSettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
				gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(OutlookImageHelperSettingsPanel)
		except Exception:
			pass
		if getattr(self, "toolsMenuItem", None):
			try:
				gui.mainFrame.sysTrayIcon.toolsMenu.RemoveItem(self.toolsMenuItem)
			except Exception:
				pass
		if getattr(self, "toolsGuideItem", None):
			try:
				gui.mainFrame.sysTrayIcon.toolsMenu.RemoveItem(self.toolsGuideItem)
			except Exception:
				pass
		if getattr(self, "_pendingInspectCall", None):
			try:
				self._pendingInspectCall.Stop()
			except Exception:
				pass
			self._pendingInspectCall = None
		self._cachedGraphicsDocKey = None
		self._cachedGraphicsItems = None
		self._cachedGraphicsTime = 0.0
		_cancel_item_sound()

		# Usunięcie wyizolowanych podkatalogów tymczasowych utworzonych w tej sesji
		for folder in list(_active_temp_attachment_folders):
			try:
				if os.path.isdir(folder):
					shutil.rmtree(folder, ignore_errors=True)
			except Exception:
				pass
		_active_temp_attachment_folders.clear()

		GlobalPlugin.instance = None

	def onToolsMenuOpen(self, evt):
		self.script_openGraphicList(None)

	def isOutlookActive(self) -> bool:
		"""Sprawdza, czy bieżące aktywne okno lub obiekt fokusu należy do programu Outlook."""
		focus = api.getFocusObject()
		if not focus:
			return False
		appMod = getattr(focus, "appModule", None)
		if appMod and getattr(appMod, "appName", "").lower() == "outlook":
			return True
		fgHwnd = winUser.getForegroundWindow()
		if fgHwnd:
			cls = winUser.getClassName(fgHwnd)
			if "renwnd32" in cls.lower():
				return True
		return False

	def _checkOutlookOrPassGesture(self, gesture) -> bool:
		"""
		Sprawdza, czy użytkownik znajduje się w programie Microsoft Outlook.
		Jeśli nie:
		- dla skrótów bez klawisza NVDA (np. Alt+Shift+G/O/A/I) przepuszcza skrót do aktywnej aplikacji (gesture.send()),
		  aby nie blokować skrótów innych programów ani przełączania układu klawiatury;
		- dla skrótów z klawiszem NVDA (lub wywołania z menu) wypowiada komunikat informacyjny.
		"""
		if self.isOutlookActive():
			return True
		if gesture is not None:
			try:
				idents = " ".join(getattr(gesture, "identifiers", []) or []).lower()
				if "nvda" not in idents and "insert" not in idents and "capslock" not in idents:
					gesture.send()
					return False
			except Exception:
				pass
		ui.message(_T("msg_outlook_only"))
		return False

	def event_gainFocus(self, obj, nextHandler):
		nextHandler()
		if not self.isOutlookActive():
			return

		sec = config.conf[CONFIG_SECTION]
		if not _to_bool(sec.get("autoAnnounce", True), default=True) and not _to_bool(sec.get("playSound", True), default=True):
			return

		triggerMode = sec.get("soundTriggerMode", "open_only")
		if triggerMode == "off":
			return

		hwnd = getattr(obj, "windowHandle", 0)
		if not hwnd:
			return

		cls = winUser.getClassName(hwnd)

		# 1. Sprawdzenie, czy fokus wszedł do właściwej treści otwartej wiadomości (_WwG)
		if cls == "_WwG":
			if triggerMode not in ("open_only", "both"):
				return

			msgKey = getMessageKey()
			currentTime = time.time()
			# Unikaj wielokrotnego odtwarzania dla tej samej wiadomości (okno 8 sekund)
			if msgKey and msgKey == getattr(self, "_lastAnnouncedMsgKey", "") and (currentTime - getattr(self, "_lastAnnouncedTime", 0)) < 8.0:
				return

			self._lastAnnouncedMsgKey = msgKey
			self._lastAnnouncedTime = currentTime
			self._lastNavigatedDocKey = None
			self._lastNavigatedIndex = -1
			self._lastNavigatedCaretPos = -9999
			self._currentGraphicItem = None
			if hasattr(self, "_pendingInspectCall") and self._pendingInspectCall:
				try:
					self._pendingInspectCall.Stop()
				except Exception:
					pass
			self._pendingInspectCall = wx.CallLater(150, self._inspectAndAnnounceMessage, msgKey)
			return

		# 2. Sprawdzenie, czy fokus wszedł na wiersz listy wiadomości w głównym oknie Outlooka (Explorer)
		if triggerMode in ("list_only", "both"):
			if isGroupHeader(obj):
				return

			cachedCls = getattr(getattr(obj, "UIAElement", None), "cachedClassName", "") or getattr(obj, "windowClassName", "") or ""
			role = getattr(obj, "role", None)
			isRow = cachedCls in ("LeafRow", "ThreadItem", "TableItem", "UIAGridRow") or role in (ROLE_TABLEROW, controlTypes.Role.TABLECELL, controlTypes.Role.LISTITEM)
			if not isRow:
				return

			# Koordynacja i deduplikacja z Outlook Speech Filter dla kategorii "obrazy / grafiki wklejone"
			if isSpeechFilterHandlingInlineImages():
				vol = sec.get("soundVolume", 80)
				_log_sound_event("image_alert", vol, "Lista wiadomości Outlook (Explorer)", "Pominięto (deduplikacja: obsługa przez Outlook Speech Filter)")
				return

			# Jeśli Speech Filter nie obsługuje grafik na liście, Pomocnik obrazów przejmuje obsługę
			mailItem = getCurrentMailItem()
			if mailItem and mailItemHasInlineImages(mailItem):
				msgKey = getMessageKey(mailItem)
				currentTime = time.time()
				if msgKey and msgKey == getattr(self, "_lastListAnnouncedMsgKey", "") and (currentTime - getattr(self, "_lastListAnnouncedTime", 0)) < 1.5:
					return
				self._lastListAnnouncedMsgKey = msgKey
				self._lastListAnnouncedTime = currentTime
				vol = sec.get("soundVolume", 80)
				playAlertSound(soundName="image_alert", volume=vol, context="Lista wiadomości Outlook (Explorer)")

	def _inspectAndAnnounceMessage(self, expectedMsgKey: str):
		self._pendingInspectCall = None
		if not self.isOutlookActive():
			return

		doc = getActiveWordDocument()
		if not doc:
			return

		curKey = getMessageKey()
		if expectedMsgKey and curKey and expectedMsgKey != curKey:
			return

		items = getOrCollectGraphics(
			doc,
			minDimension=config.conf[CONFIG_SECTION].get("minDimension", 15),
			fast=True,
		)
		if not items:
			return

		sec = config.conf[CONFIG_SECTION]
		unlabeledCount = sum(1 for it in items if it.isScreenshotOrUnlabeled)

		if sec.get("playSound", True) and unlabeledCount > 0:
			vol = sec.get("soundVolume", 80)
			playAlertSound(soundName="image_alert", volume=vol, context="Otwarta treść wiadomości (_WwG)")

		if sec.get("autoAnnounce", True):
			if len(items) == 1 and unlabeledCount == 1:
				ui.message(_T("msg_one_screenshot"))
			elif unlabeledCount > 0:
				ui.message(_T("msg_items_found", total=len(items), unlabeled=unlabeledCount))
			else:
				ui.message(_T("msg_items_all_labeled", total=len(items)))

	# ==========================================================================
	# SKRYPTY GESTÓW
	# ==========================================================================

	@script(
		description="Otwiera okno dialogowe z listą wszystkich grafik oraz załączników w bieżącej wiadomości programu Outlook",
		category="Outlook - Pomocnik obrazów, załączników i OCR",
		gestures=[
			"kb:NVDA+shift+g",
			"kb:alt+shift+g",
			"kb:NVDA+control+a",
		],
	)
	def script_openGraphicList(self, gesture):
		if not self._checkOutlookOrPassGesture(gesture):
			return

		focus = api.getFocusObject()
		if isGroupHeader(focus):
			ui.message("Fokus znajduje się na nagłówku grupy (np. 'Dzisiaj'), a nie na wiadomości e-mail. Zaznacz konkretną wiadomość na liście lub otwórz ją klawiszem Enter.")
			return

		openGraphicListDialog()

	@script(
		description="Szybkie rozpoznanie tekstu (OCR) bieżącej grafiki lub zrzutu ekranu w wiadomości",
		category="Outlook - Pomocnik obrazów, załączników i OCR",
		gestures=[
			"kb:NVDA+windows+g",
			"kb:NVDA+alt+g",
			"kb:alt+shift+o",
		],
	)
	def script_quickOcr(self, gesture):
		if not self._checkOutlookOrPassGesture(gesture):
			return

		focus = api.getFocusObject()
		if isGroupHeader(focus):
			ui.message("Fokus znajduje się na nagłówku grupy, a nie na wiadomości e-mail. Zaznacz konkretną wiadomość na liście lub otwórz ją klawiszem Enter.")
			return

		selCount, sel, isMulti = getSelectedMessagesInfo()
		if isMulti:
			handleMultipleMessagesSelected(selCount, sel, action="ocr")
			return

		mailItem = getCurrentMailItem()
		doc = getActiveWordDocument(mailItem=mailItem)
		if not mailItem and doc:
			mailItem = getCurrentMailItem(doc=doc)

		items = []
		totalLen = 0
		if doc:
			items = getOrCollectGraphics(doc, minDimension=config.conf[CONFIG_SECTION].get("minDimension", 15), fast=True)
			try:
				totalLen = len(doc.Content.Text)
			except Exception:
				pass

		attachments = []
		if mailItem:
			attachments = extractAttachmentsFromCurrentMessage(mailItem=mailItem, includeInline=False)

		if not items and not attachments:
			if isMessageListFocused() or not isMessageBodyFocused():
				handleReadingPaneOff(mailItem, action="ocr")
				return
			if mailItem:
				try:
					subj = str(getattr(mailItem, "Subject", "") or "").strip()
					if subj:
						ui.message(_T("msg_no_ocr_subj", subj=subj))
						return
				except Exception:
					pass
			ui.message(_T("msg_no_elements"))
			return

		targetItem = getGraphicAtCaret(doc, strict=True) if doc else None
		if not targetItem:
			unlabeled = [it for it in items if it.isScreenshotOrUnlabeled]
			if len(items) == 1 and not attachments:
				targetItem = items[0]
			elif len(unlabeled) == 1 and not attachments:
				targetItem = unlabeled[0]
			elif not items and len(attachments) == 1:
				targetItem = attachments[0]
			else:
				ui.message(f"Wiadomość zawiera {len(items)} grafik oraz {len(attachments)} załączników. Otwieram listę wyboru...")
				d = GraphicElementsDialog(gui.mainFrame, items, attachments=attachments, totalDocLength=totalLen)
				gui.mainFrame.prePopup()
				d.Show()
				gui.mainFrame.postPopup()
				return

		self.triggerOcrForItem(targetItem)

	@script(
		description="Szybki opis bieżącej grafiki lub załącznika w wiadomości przez sztuczną inteligencję (AI)",
		category="Outlook - Pomocnik obrazów, załączników i OCR",
		gestures=[
			"kb:NVDA+windows+shift+g",
			"kb:NVDA+alt+shift+g",
			"kb:alt+shift+a",
		],
	)
	def script_quickAi(self, gesture):
		if not self._checkOutlookOrPassGesture(gesture):
			return

		focus = api.getFocusObject()
		if isGroupHeader(focus):
			ui.message("Fokus znajduje się na nagłówku grupy, a nie na wiadomości e-mail. Zaznacz konkretną wiadomość na liście lub otwórz ją klawiszem Enter.")
			return

		selCount, sel, isMulti = getSelectedMessagesInfo()
		if isMulti:
			handleMultipleMessagesSelected(selCount, sel, action="ai")
			return

		mailItem = getCurrentMailItem()
		doc = getActiveWordDocument(mailItem=mailItem)
		if not mailItem and doc:
			mailItem = getCurrentMailItem(doc=doc)

		items = []
		totalLen = 0
		if doc:
			items = getOrCollectGraphics(doc, minDimension=config.conf[CONFIG_SECTION].get("minDimension", 15), fast=True)
			try:
				totalLen = len(doc.Content.Text)
			except Exception:
				pass

		attachments = []
		if mailItem:
			attachments = extractAttachmentsFromCurrentMessage(mailItem=mailItem, includeInline=False)

		if not items and not attachments:
			if isMessageListFocused() or not isMessageBodyFocused():
				handleReadingPaneOff(mailItem, action="ai")
				return
			if mailItem:
				try:
					subj = str(getattr(mailItem, "Subject", "") or "").strip()
					if subj:
						ui.message(_T("msg_no_ai_subj", subj=subj))
						return
				except Exception:
					pass
			ui.message(_T("msg_no_elements"))
			return

		targetItem = getGraphicAtCaret(doc, strict=True) if doc else None
		if not targetItem:
			if len(items) == 1 and not attachments:
				targetItem = items[0]
			elif not items and len(attachments) == 1:
				targetItem = attachments[0]
			else:
				unlabeled = [it for it in items if it.isScreenshotOrUnlabeled]
				if len(unlabeled) == 1 and not attachments:
					targetItem = unlabeled[0]

		if not targetItem:
			ui.message(f"Wiadomość zawiera {len(items)} grafik oraz {len(attachments)} załączników. Wybierz element z listy:")
			d = GraphicElementsDialog(gui.mainFrame, items, attachments=attachments, totalDocLength=totalLen)
			gui.mainFrame.prePopup()
			d.Show()
			gui.mainFrame.postPopup()
			return

		self.triggerAiDescriptionForItem(targetItem)

	@script(
		description=_T("script_quick_invert_doc"),
		category="Outlook - Pomocnik obrazów, załączników i OCR",
		gestures=[
			"kb:NVDA+windows+i",
			"kb:NVDA+alt+i",
			"kb:alt+shift+i",
		],
	)
	def script_quickInvert(self, gesture):
		if not self._checkOutlookOrPassGesture(gesture):
			return

		focus = api.getFocusObject()
		if isGroupHeader(focus):
			ui.message("Fokus znajduje się na nagłówku grupy, a nie na wiadomości e-mail. Zaznacz konkretną wiadomość na liście lub otwórz ją klawiszem Enter.")
			return

		selCount, sel, isMulti = getSelectedMessagesInfo()
		if isMulti:
			handleMultipleMessagesSelected(selCount, sel, action="invert")
			return

		mailItem = getCurrentMailItem()
		doc = getActiveWordDocument(mailItem=mailItem)
		if not mailItem and doc:
			mailItem = getCurrentMailItem(doc=doc)

		items = []
		totalLen = 0
		if doc:
			items = getOrCollectGraphics(doc, minDimension=config.conf[CONFIG_SECTION].get("minDimension", 15), fast=True)
			try:
				totalLen = len(doc.Content.Text)
			except Exception:
				pass

		attachments = []
		if mailItem:
			attachments = extractAttachmentsFromCurrentMessage(mailItem=mailItem, includeInline=False)

		if not items and not attachments:
			if isMessageListFocused() or not isMessageBodyFocused():
				handleReadingPaneOff(mailItem, action="invert")
				return
			if mailItem:
				try:
					subj = str(getattr(mailItem, "Subject", "") or "").strip()
					if subj:
						ui.message(_T("msg_no_preview_subj", subj=subj))
						return
				except Exception:
					pass
			ui.message(_T("msg_no_elements"))
			return

		targetItem = getGraphicAtCaret(doc, strict=True) if doc else None
		if not targetItem:
			realGraphics = [it for it in items if not it.isDivider]
			if len(realGraphics) == 1:
				targetItem = realGraphics[0]
			elif len(items) == 1:
				targetItem = items[0]
			elif not items:
				imgAtts = [a for a in attachments if getattr(a, "isImage", False)]
				if len(imgAtts) == 1:
					targetItem = imgAtts[0]
			else:
				unlabeled = [it for it in realGraphics if it.isScreenshotOrUnlabeled]
				if len(unlabeled) == 1:
					targetItem = unlabeled[0]

		if not targetItem:
			ui.message("Wybierz element do inwersji kolorów z listy:")
			d = GraphicElementsDialog(gui.mainFrame, items, attachments=attachments, totalDocLength=totalLen)
			gui.mainFrame.prePopup()
			d.Show()
			gui.mainFrame.postPopup()
			return

		# Wyodrębnij obraz i uruchom podgląd w negatywie
		if isinstance(targetItem, AttachmentItem):
			if not getattr(targetItem, "isImage", False):
				ui.message(_T("preview_only_images"))
				return
			tempPath = saveAttachmentToTempFile(targetItem)
		else:
			ui.message(_T("msg_extracting_preview"))
			tempPath = extractGraphicToPngFile(targetItem)

		if not tempPath or not os.path.isfile(tempPath):
			ui.message(_T("err_extract_img"))
			return

		gui.mainFrame.prePopup()
		try:
			d = ImagePreviewDialog(gui.mainFrame, targetItem, tempPath, startInverted=True)
			d.ShowModal()
			d.Destroy()
		finally:
			gui.mainFrame.postPopup()
		try:
			if tempPath and os.path.isfile(tempPath):
				os.remove(tempPath)
		except Exception:
			pass

	@script(
		description="Otwiera menu czynności dla bieżącej grafiki w wiadomości e-mail programu Outlook",
		category="Outlook - Pomocnik obrazów, załączników i OCR",
		gestures=[
			"kb(outlook):alt+c",
			"kb(outlook):alt+shift+c",
			"kb(outlook):NVDA+alt+c",
		],
	)
	def script_graphicActionsMenu(self, gesture):
		if not self.isOutlookActive():
			gesture.send()
			return

		doc = getActiveWordDocument()
		if not doc:
			gesture.send()
			return

		item = getGraphicAtCaret(doc, strict=False)
		if not item:
			items = getOrCollectGraphics(doc, fast=True)
			if items:
				item = items[0]
			else:
				if isMessageListFocused() or not isMessageBodyFocused():
					ui.message("Aby otworzyć menu czynności dla grafiki w tej wiadomości, otwórz ją klawiszem Enter.")
				else:
					ui.message("W bieżącej wiadomości nie odnaleziono grafik.")
				return

		showGraphicActionsMenu(item)

	@script(
		description="Kontekstowe menu czynności dla grafiki w wiadomości lub standardowe menu programu Outlook",
		category="Outlook - Pomocnik obrazów, załączników i OCR",
		gestures=[
			"kb(outlook):applications",
			"kb(outlook):shift+f10",
			"kb:applications",
			"kb:shift+f10",
		],
	)
	def script_graphicAppsMenu(self, gesture):
		if not self.isOutlookActive():
			gesture.send()
			return

		contextMenuEnabled = True
		try:
			contextMenuEnabled = _to_bool(config.conf[CONFIG_SECTION].get("contextMenuEnabled", True), default=True)
		except Exception:
			pass

		if not contextMenuEnabled:
			gesture.send()
			return

		focus = api.getFocusObject()
		ti = getattr(focus, "treeInterceptor", None)

		doc = getActiveWordDocument()
		if not doc:
			gesture.send()
			return

		item = getStrictGraphicUnderCaret(doc, treeInterceptor=ti)
		if item:
			import wx
			wx.CallAfter(showGraphicActionsMenu, item)
			return

		# Jeśli kursor NIE stoi ściśle na grafice, przepuść gest do standardowego menu Outlooka
		gesture.send()


	# ==========================================================================
	# WYKONYWANIE ZADAŃ OCR ORAZ AI
	# ==========================================================================

	def triggerOcrForItem(self, item: GraphicItem | AttachmentItem, parentDialog: GraphicElementsDialog | None = None):
		"""Uruchamia proces OCR lub ekstrakcji tekstu dla wskazanej grafiki lub załącznika."""
		if not ensureGraphicReady(item, parent=parentDialog, actionName="rozpoznania tekstu (OCR)"):
			return

		if item.ocrText:
			def showCached():
				if item.ocrText in ("(Brak rozpoznanego tekstu)", "(Brak tekstu na obrazie)", "(Brak tekstu w dokumencie PDF)"):
					if isinstance(item, AttachmentItem):
						targetLabel = item.fileName
						body = _T("ocr_no_text_att_body", filename=item.fileName, size=item.formattedSize)
					else:
						targetLabel = f"grafika {item.index}"
						body = _T("ocr_no_text_body", index=item.index, width=item.widthPx, height=item.heightPx)
					ui.message(_T("ocr_no_text_msg", target=targetLabel))
					ui.browseableMessage(body, title=_T("ocr_no_text_title", target=targetLabel))
				else:
					lineCount = len([l for l in item.ocrText.splitlines() if l.strip()])
					label = f"Załącznik {item.index}: {item.fileName}" if isinstance(item, AttachmentItem) else f"Grafika {item.index}"
					ui.message(f"Rozpoznany tekst z {label} ({lineCount} wierszy).")
					title = f"Rozpoznany tekst — {item.fileName}" if isinstance(item, AttachmentItem) else f"Rozpoznany tekst — Grafika {item.index} ({item.widthPx}×{item.heightPx} px)"
					ui.browseableMessage(item.ocrText, title=title)
			wx.CallAfter(showCached)
			return

		if isinstance(item, AttachmentItem):
			self._triggerOcrForAttachment(item, parentDialog=parentDialog)
			return

		# Standardowa grafika w treści
		if getattr(item, "isBlankOrSolid", False):
			reason = getattr(item, "blankReason", "") or "brak widocznej zawartości"
			if not getattr(item, "isDivider", False):
				showBlockedImageInstructionsDialog(parentDialog, item, actionName="rozpoznania tekstu (OCR)")
				return
			ui.message(f"Grafika {item.index} to pusty ozdobnik lub element przezroczysty ({reason}). Brak tekstu.")
			item.ocrText = "(Brak tekstu na obrazie)"
			return

		ui.message(f"Pobieranie grafiki {item.index}...")
		tempPng = extractGraphicToPngFile(item)
		if not tempPng or not os.path.isfile(tempPng):
			ui.message("Błąd: Nie udało się wyodrębnić obrazu z edytora wiadomości.")
			return

		if getattr(item, "isBlankOrSolid", False):
			reason = getattr(item, "blankReason", "") or "brak widocznej zawartości"
			if not getattr(item, "isDivider", False):
				showBlockedImageInstructionsDialog(parentDialog, item, actionName="rozpoznania tekstu (OCR)")
				return
			ui.message(f"Grafika {item.index} to pusty ozdobnik lub element przezroczysty ({reason}). Brak tekstu.")
			item.ocrText = "(Brak tekstu na obrazie)"
			return

		ui.message(f"Trwa rozpoznawanie tekstu z grafiki {item.index} (OCR)...")

		def worker():
			lang = config.conf[CONFIG_SECTION].get("defaultOcrLang", "pl-PL")
			try:
				text = runWindowsOcr(tempPng, lang=lang)
				if not text:
					item.ocrText = "(Brak rozpoznanego tekstu)"
					def showEmpty():
						if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
							parentDialog.refreshListLabels()
						targetLabel = f"grafika {item.index}"
						ui.message(_T("ocr_no_text_msg", target=targetLabel))
						msg = _T(
							"ocr_no_text_body",
							index=item.index,
							width=item.widthPx,
							height=item.heightPx,
						)
						ui.browseableMessage(
							msg,
							title=_T("ocr_no_text_title", target=targetLabel),
						)
					wx.CallAfter(showEmpty)
					return

				item.ocrText = text
				lineCount = len([l for l in text.splitlines() if l.strip()])
				summaryMsg = f"Rozpoznano tekst z grafiki {item.index} ({lineCount} wierszy)."

				def showResult():
					if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
						parentDialog.refreshListLabels()
					ui.message(summaryMsg)
					ui.browseableMessage(
						text,
						title=f"Rozpoznany tekst — Grafika {item.index} ({item.widthPx}×{item.heightPx} px)",
					)

				wx.CallAfter(showResult)
			except Exception as ex:
				log.error(f"outlookImageHelper: Błąd OCR: {ex}", exc_info=True)
				wx.CallAfter(ui.message, f"Błąd podczas rozpoznawania OCR: {ex}")
			finally:
				try:
					if tempPng and os.path.isfile(tempPng):
						os.remove(tempPng)
				except Exception:
					pass

		t = threading.Thread(target=worker, daemon=True)
		t.start()

	def _triggerOcrForAttachment(self, item: AttachmentItem, parentDialog: GraphicElementsDialog | None = None):
		"""Obsługuje OCR lub ekstrakcję tekstu dla pliku załącznika (PDF, DOCX, grafika, plik tekstowy)."""
		ui.message(f"Pobieranie załącznika: {item.fileName}...")
		tempPath = saveAttachmentToTempFile(item)
		if not tempPath or not os.path.isfile(tempPath):
			ui.message(f"Błąd: Nie udało się pobrać pliku {item.fileName}.")
			return

		if getattr(item, "isDoc", False):
			ui.message(f"Plik '{item.fileName}' jest w starszym formacie binarnym (.doc). Zapisz go jako .docx, aby odczytać treść.")
			return

		if item.isDocx:
			ui.message(f"Odczytywanie zawartości dokumentu Word: {item.fileName}...")
			def docxWorker():
				try:
					txt = extractTextFromDocxFile(tempPath)
					if not txt:
						txt = "(Dokument Word jest pusty lub zawiera wyłącznie obiekty graficzne)"
					item.ocrText = txt
					def showDocx():
						if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
							parentDialog.refreshListLabels()
						ui.message(f"Odczytano zawartość dokumentu {item.fileName}.")
						ui.browseableMessage(txt, title=f"Zawartość dokumentu: {item.fileName}")
					wx.CallAfter(showDocx)
				except Exception as ex:
					log.error(f"outlookImageHelper: Błąd odczytu DOCX: {ex}", exc_info=True)
					wx.CallAfter(ui.message, f"Błąd odczytu pliku Word: {ex}")
			threading.Thread(target=docxWorker, daemon=True).start()
			return

		if item.isText:
			ui.message(f"Odczytywanie pliku tekstowego: {item.fileName}...")
			def textWorker():
				try:
					with open(tempPath, "r", encoding="utf-8", errors="replace") as f:
						txt = f.read()
					item.ocrText = txt or "(Plik jest pusty)"
					def showTxt():
						if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
							parentDialog.refreshListLabels()
						ui.message(f"Odczytano plik {item.fileName}.")
						ui.browseableMessage(item.ocrText, title=f"Zawartość pliku: {item.fileName}")
					wx.CallAfter(showTxt)
				except Exception as ex:
					wx.CallAfter(ui.message, f"Błąd odczytu pliku: {ex}")
			threading.Thread(target=textWorker, daemon=True).start()
			return

		if item.isPdf:
			# Jeśli plik PDF jest chroniony hasłem, odblokuj go przed wysłaniem do OCR
			unlockedPdf = ensurePdfUnlocked(item, parentWin=parentDialog)
			if not unlockedPdf or not os.path.isfile(unlockedPdf):
				ui.message("Anulowano odblokowywanie dokumentu PDF lub błędne hasło.")
				return
			tempPath = unlockedPdf

			# Dla PDF tekst rozpoznawany jest za pomocą multimodalnej sztucznej inteligencji
			prov = getActiveAiProvider()
			provName = "Google Gemini" if prov == "gemini" else "OpenAI (ChatGPT)"
			apiKey = getActiveAiApiKey()
			if not apiKey:
				def onKeySavedPdf():
					self._triggerOcrForAttachment(item, parentDialog)
				checkAndPromptAiApiKey(parentWin=parentDialog, onKeyReadyCallback=onKeySavedPdf)
				return

			fSize = os.path.getsize(tempPath) if os.path.isfile(tempPath) else 0
			if fSize > 20 * 1024 * 1024:
				ui.message(f"Plik PDF przekracza maksymalny limit 20 MB ({round(fSize / (1024 * 1024), 1)} MB).")
				return

			fSizeMb = round(fSize / (1024 * 1024), 2)
			confirmMsg = (
				f"Plik PDF '{item.fileName}' ({fSizeMb} MB) zostanie przesłany do chmury {provName} "
				f"w celu precyzyjnego przepisania tekstu (OCR AI).\n\n"
				f"Czy wyrażasz zgodę na transmisję tego dokumentu do {provName}?"
			)
			if gui.messageBox(confirmMsg, f"Potwierdzenie OCR w chmurze ({provName})", wx.YES_NO | wx.ICON_QUESTION, parent=parentDialog) != wx.YES:
				ui.message("Anulowano wysyłanie dokumentu PDF do chmury.")
				return

			ui.message(f"Przepisywanie tekstu ze skanu PDF ({item.fileName}) przez chmurę {provName}...")
			def pdfWorker():
				sec = config.conf[CONFIG_SECTION]
				model = sec.get("geminiModel", _DEFAULT_GEMINI_MODEL) if prov == "gemini" else sec.get("openaiModel", "gpt-4o-mini")
				try:
					if prov == "gemini":
						txt = describeWithGemini(
							tempPath,
							apiKey,
							model=model,
							customPrompt="Dokonaj dokładnego OCR i przepisz cały tekst z tego dokumentu PDF w oryginalnym układzie i strukturze. Zachowaj tabele, nagłówki, punkty i kwoty. Nie dodawaj wstępów ani komentarzy.",
							mimeType="application/pdf",
						)
					else:
						txt = describeWithOpenAi(
							tempPath,
							apiKey,
							model=model,
							customPrompt="Dokonaj dokładnego OCR i przepisz cały tekst z tego dokumentu PDF w oryginalnym układzie i strukturze. Zachowaj tabele, nagłówki, punkty i kwoty. Nie dodawaj wstępów ani komentarzy.",
							mimeType="application/pdf",
						)
					if not txt:
						item.ocrText = "(Brak rozpoznanego tekstu)"
						def showPdfEmpty():
							if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
								parentDialog.refreshListLabels()
							ui.message(_T("ocr_no_text_msg", target=item.fileName))
							msg = _T("ocr_no_text_att_body", filename=item.fileName, size=item.formattedSize)
							ui.browseableMessage(msg, title=_T("ocr_no_text_title", target=item.fileName))
						wx.CallAfter(showPdfEmpty)
					else:
						item.ocrText = txt
						def showPdf():
							if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
								parentDialog.refreshListLabels()
							ui.message(f"Ukończono rozpoznawanie tekstu z PDF: {item.fileName}.")
							ui.browseableMessage(item.ocrText, title=f"Rozpoznany tekst z PDF: {item.fileName}")
						wx.CallAfter(showPdf)
				except Exception as ex:
					log.error(f"outlookImageHelper: Błąd OCR PDF: {ex}", exc_info=True)
					wx.CallAfter(ui.message, f"Błąd OCR PDF: {ex}")
			threading.Thread(target=pdfWorker, daemon=True).start()
			return

		if item.isImage:
			ui.message("Trwa rozpoznawanie tekstu z załączonego obrazu (OCR)...")
			def imgWorker():
				lang = config.conf[CONFIG_SECTION].get("defaultOcrLang", "pl-PL")
				try:
					txt = runWindowsOcr(tempPath, lang=lang)
					if not txt:
						item.ocrText = "(Brak rozpoznanego tekstu)"
						def showImgEmpty():
							if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
								parentDialog.refreshListLabels()
							ui.message(_T("ocr_no_text_msg", target=item.fileName))
							msg = _T("ocr_no_text_att_body", filename=item.fileName, size=item.formattedSize)
							ui.browseableMessage(msg, title=_T("ocr_no_text_title", target=item.fileName))
						wx.CallAfter(showImgEmpty)
					else:
						item.ocrText = txt
						def showImg():
							if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
								parentDialog.refreshListLabels()
							ui.message(f"Zakończono OCR dla: {item.fileName}.")
							ui.browseableMessage(item.ocrText, title=f"Rozpoznany tekst — {item.fileName}")
						wx.CallAfter(showImg)
				except Exception as ex:
					log.error(f"outlookImageHelper: Błąd OCR załącznika obrazu: {ex}", exc_info=True)
					wx.CallAfter(ui.message, f"Błąd OCR obrazu: {ex}")
			threading.Thread(target=imgWorker, daemon=True).start()
			return

		ui.message(f"Dla typu pliku {item.ext} OCR nie jest bezpośrednio obsługiwany.")

	def triggerBatchOcr(self, items: list[GraphicItem | AttachmentItem], parentDialog: GraphicElementsDialog | None = None):
		"""Zbiorcze rozpoznanie OCR dla listy zaznaczonych grafik i załączników."""
		if not items:
			return

		pdfItems = [it for it in items if isinstance(it, AttachmentItem) and it.isPdf]
		allowCloudPdf = False
		prov = getActiveAiProvider()
		provName = "Google Gemini" if prov == "gemini" else "OpenAI (ChatGPT)"
		apiKey = getActiveAiApiKey()

		if pdfItems:
			if apiKey:
				totalPdfBytes = sum(getattr(it.rawAttachment, "Size", 0) for it in pdfItems if hasattr(it, "rawAttachment"))
				totalPdfMb = round(totalPdfBytes / (1024 * 1024), 2)
				confirmMsg = (
					f"Wśród zaznaczonych elementów znajduje się {len(pdfItems)} dokument(ów) PDF "
					f"o łącznym rozmiarze ok. {totalPdfMb} MB, które wymagają wysłania do chmury {provName} w celu OCR AI.\n\n"
					f"Czy wyrażasz zgodę na transmisję tych plików do usługi {provName}?"
				)
				allowCloudPdf = (gui.messageBox(confirmMsg, f"Potwierdzenie OCR w chmurze ({provName})", wx.YES_NO | wx.ICON_QUESTION, parent=parentDialog) == wx.YES)
			else:
				ui.message(f"Brak klucza API dla {provName}. OCR dla plików PDF zostanie pominięty.")

		ui.message(f"Rozpoczynanie OCR dla {len(items)} elementów...")

		# Pre-ekstrakcja plików na głównym wątku STA przed uruchomieniem wątku roboczego
		prepared_items = []
		for it in items:
			try:
				if isinstance(it, AttachmentItem):
					p = saveAttachmentToTempFile(it)
					if it.isPdf and p:
						unlocked = ensurePdfUnlocked(it, parentWin=parentDialog)
						if not unlocked:
							ui.message(f"Pominięto zaszyfrowany plik PDF: {it.fileName}")
							p = None
						else:
							p = unlocked
					prepared_items.append((it, p))
				else:
					p = extractGraphicToPngFile(it)
					prepared_items.append((it, p))
			except Exception as ex:
				log.debug(f"outlookImageHelper: błąd pre-ekstrakcji elementu: {ex}")
				prepared_items.append((it, None))

		def worker():
			lang = config.conf[CONFIG_SECTION].get("defaultOcrLang", "pl-PL")
			sec = config.conf[CONFIG_SECTION]
			results = []

			for idx, (it, tempPath) in enumerate(prepared_items):
				if isinstance(it, AttachmentItem):
					header = f"=== Załącznik {it.index}: {it.fileName} ({it.formattedSize}) ==="
					if getattr(it, "isDoc", False):
						results.append(f"{header}\n(Format .doc nie jest obsługiwany — wymagany format .docx)")
					elif it.isDocx and tempPath:
						txt = extractTextFromDocxFile(tempPath)
						it.ocrText = txt or "(Pusty dokument)"
						results.append(f"{header}\n{it.ocrText}")
					elif it.isText and tempPath:
						with open(tempPath, "r", encoding="utf-8", errors="replace") as f:
							it.ocrText = f.read(500000)
						results.append(f"{header}\n{it.ocrText}")
					elif it.isPdf and tempPath:
						if not allowCloudPdf:
							results.append(f"{header}\n(Pominięto OCR w chmurze — brak zgody użytkownika na transmisję)")
						elif apiKey:
							try:
								if prov == "gemini":
									txt = describeWithGemini(
										tempPath,
										apiKey,
										model=sec.get("geminiModel", _DEFAULT_GEMINI_MODEL),
										customPrompt="Przepisz dokładnie tekst z tego dokumentu PDF.",
										mimeType="application/pdf",
									)
								else:
									txt = describeWithOpenAi(
										tempPath,
										apiKey,
										model=sec.get("openaiModel", "gpt-4o-mini"),
										customPrompt="Przepisz dokładnie tekst z tego dokumentu PDF.",
										mimeType="application/pdf",
									)
								it.ocrText = txt
								results.append(f"{header}\n{it.ocrText}")
							except Exception as pe:
								results.append(f"{header}\nBłąd OCR PDF: {pe}")
					elif it.isImage and tempPath:
						try:
							txt = runWindowsOcr(tempPath, lang=lang)
							it.ocrText = txt or "(Brak tekstu)"
							results.append(f"{header}\n{it.ocrText}")
						except Exception as ie:
							results.append(f"{header}\nBłąd OCR obrazu: {ie}")
					else:
						results.append(f"{header}\n(Nieobsługiwany format do OCR)")
				else:
					tempPng = tempPath
					header = f"=== Grafika {it.index}"
					if it.contextText:
						header += f" [{it.contextText}]"
					header += f" ({it.widthPx}×{it.heightPx} px) ==="
					if tempPng and os.path.isfile(tempPng):
						try:
							txt = runWindowsOcr(tempPng, lang=lang)
							it.ocrText = txt or "(Brak tekstu)"
							results.append(f"{header}\n{it.ocrText}")
						except Exception as ex:
							results.append(f"{header}\nBłąd OCR: {ex}")
						finally:
							try:
								os.remove(tempPng)
							except Exception:
								pass
					else:
						results.append(f"{header}\nNie udało się wyodrębnić grafiki.")

			combined = ("\n\n" + "=" * 60 + "\n\n").join(results)

			def showDone():
				if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
					parentDialog.refreshListLabels()
				ui.message(f"Zakończono OCR dla {len(items)} elementów.")
				ui.browseableMessage(combined, title=f"Zbiorczy wynik OCR ({len(items)} elementów)")

			wx.CallAfter(showDone)

		t = threading.Thread(target=worker, daemon=True)
		t.start()

	def triggerAiDescriptionForItem(
		self,
		item: GraphicItem | AttachmentItem,
		parentDialog: GraphicElementsDialog | None = None,
		customPrompt: str = "",
	):
		"""
		Uruchamia proces opisywania/analizy grafiki lub załącznika przez AI.
		Po zakończeniu inicjalizuje historię rozmowy i natychmiast otwiera
		dostępne okno dialogowe rozmowy (AiConversationDialog).
		"""
		if not ensureGraphicReady(item, parent=parentDialog, actionName="analizy i opisu przez AI"):
			return

		prov = getActiveAiProvider()
		provName = "OpenAI (ChatGPT)" if prov == "openai" else "Google Gemini"
		apiKey = getActiveAiApiKey()
		if not apiKey:
			def onKeyReady():
				if GlobalPlugin.instance:
					GlobalPlugin.instance.triggerAiDescriptionForItem(
						item, parentDialog=parentDialog, customPrompt=customPrompt
					)
			checkAndPromptAiApiKey(parentWin=parentDialog, onKeyReadyCallback=onKeyReady)
			return

		# Jeśli rozmowa lub opis już istnieje i nie podano nowego prompta, otwórz okno rozmowy od razu
		if item.chatHistory and not customPrompt:
			def showChat():
				actualParent = parentDialog if (parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0))) else gui.mainFrame
				d = AiConversationDialog(actualParent, item)
				d.ShowModal()
				d.Destroy()
				if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
					parentDialog.refreshListLabels()
			wx.CallAfter(showChat)
			return

		isAtt = isinstance(item, AttachmentItem)
		itemLabel = f"załącznika {item.fileName}" if isAtt else f"grafiki {item.index}"

		if isAtt:
			if item.isPdf:
				targetPath = ensurePdfUnlocked(item, parentWin=parentDialog)
				if not targetPath:
					return
			else:
				ui.message(f'Pobieranie {itemLabel}...')
				targetPath = saveAttachmentToTempFile(item)
				if not targetPath or not os.path.isfile(targetPath):
					ui.message(f'Błąd: Nie udało się pobrać {itemLabel}.')
					return
			tempPng = None
		else:
			if getattr(item, "isBlankOrSolid", False):
				reason = getattr(item, "blankReason", "") or "brak widocznej zawartości"
				if not getattr(item, "isDivider", False):
					showBlockedImageInstructionsDialog(parentDialog, item, actionName="analizy i opisu przez AI")
					return
				ui.message(f"Grafika {item.index} to pusty ozdobnik lub element przezroczysty ({reason}). Pominięto wysyłanie do AI.")
				return

			ui.message(f"Pobieranie grafiki {item.index}...")
			tempPng = extractGraphicToPngFile(item)
			targetPath = tempPng
			if not targetPath or not os.path.isfile(targetPath):
				ui.message("Błąd: Nie udało się pobrać grafiki z edytora wiadomości.")
				return

			if getattr(item, "isBlankOrSolid", False):
				reason = getattr(item, "blankReason", "") or "brak widocznej zawartości"
				if not getattr(item, "isDivider", False):
					showBlockedImageInstructionsDialog(parentDialog, item, actionName="analizy i opisu przez AI")
					return
				ui.message(f"Grafika {item.index} to pusty ozdobnik lub element przezroczysty ({reason}). Pominięto wysyłanie do AI.")
				return

			try:
				with open(targetPath, "rb") as f:
					item.base64Image = base64.b64encode(f.read()).decode("ascii")
			except Exception as be:
				log.debug(f"outlookImageHelper: base64 cache error: {be}")

		feedback = AiProgressFeedback(f"Trwa analiza {itemLabel} przez sztuczną inteligencję ({provName}), proszę czekać...")

		def worker():
			sec = config.conf[CONFIG_SECTION]
			prompt = customPrompt or sec.get("aiPrompt", "")
			try:
				mimeType = "image/png"
				prefaceText = ""
				sendPath = targetPath
				b64Data = getattr(item, "base64Data", "") or getattr(item, "base64Image", "")

				if isAtt:
					if item.isPdf:
						mimeType = "application/pdf"
						fileSize = os.path.getsize(targetPath) if os.path.isfile(targetPath) else 0
						if fileSize > 20 * 1024 * 1024:
							raise ValueError(f"Rozmiar pliku PDF ({round(fileSize / (1024 * 1024), 1)} MB) przekracza dopuszczalny limit 20 MB.")
						try:
							with open(targetPath, "rb") as f:
								item.base64Data = base64.b64encode(f.read()).decode("ascii")
								b64Data = item.base64Data
						except Exception:
							pass
					elif item.isImage:
						mimeType = "image/png" if item.ext == ".png" else "image/jpeg"
						fileSize = os.path.getsize(targetPath) if os.path.isfile(targetPath) else 0
						if fileSize > 20 * 1024 * 1024:
							raise ValueError(f"Rozmiar obrazu ({round(fileSize / (1024 * 1024), 1)} MB) przekracza dopuszczalny limit 20 MB.")
						try:
							with open(targetPath, "rb") as f:
								item.base64Data = base64.b64encode(f.read()).decode("ascii")
								b64Data = item.base64Data
						except Exception:
							pass
					elif item.isDocx:
						sendPath = ""
						docxText = extractTextFromDocxFile(targetPath)
						prefaceText = f"Oto treść załączonego dokumentu Word ({item.fileName}):\n\n{docxText[:200000]}"
					elif item.isText:
						sendPath = ""
						with open(targetPath, "r", encoding="utf-8", errors="replace") as f:
							prefaceText = f"Oto treść załączonego pliku ({item.fileName}):\n\n{f.read(40000)}"
					else:
						sendPath = ""
						prefaceText = f"Załączony plik w wiadomości: {item.fileName} ({item.formattedSize})"

				fullPrompt = prompt
				if prefaceText:
					fullPrompt = f"{prefaceText}\n\n{fullPrompt}" if fullPrompt else prefaceText

				description = callAiVisionApi(
					prompt=fullPrompt,
					imagePath=sendPath,
					base64Data=b64Data,
					mimeType=mimeType,
					parentDialog=parentDialog,
				)
				item.aiDescription = description

				firstTurn = {
					"role": "user",
					"text": fullPrompt or prompt or ("(Szczegółowa analiza załącznika)" if isAtt else "(Domyślny szczegółowy opis)"),
					"timestamp": time.strftime("%H:%M"),
				}
				b64 = getattr(item, "base64Data", "") or getattr(item, "base64Image", "")
				if b64:
					firstTurn["imageData"] = b64
					firstTurn["mimeType"] = mimeType

				item.chatHistory = [
					firstTurn,
					{
						"role": "model",
						"text": description,
						"timestamp": time.strftime("%H:%M"),
					},
				]

				def openChatWindow():
					feedback.stop()
					playSuccessChime()
					if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
						parentDialog.refreshListLabels()
					ui.message(f"Analiza AI dla {itemLabel} zakończona sukcesem. Otwieram okno rozmowy...")
					actualParent = parentDialog if (parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0))) else gui.mainFrame
					d = AiConversationDialog(actualParent, item)
					d.ShowModal()
					d.Destroy()
					if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
						parentDialog.refreshListLabels()

				wx.CallAfter(openChatWindow)
			except Exception as ex:
				errStr = str(ex)
				log.error(f"outlookImageHelper: Błąd AI: {errStr}", exc_info=True)
				def onFail(errStr=errStr):
					feedback.stop(f"Błąd generowania opisu AI: {errStr}")
					if "503" in errStr or "Service Unavailable" in errStr:
						msg = (
							"BŁĄD POŁĄCZENIA Z SERWEREM AI (HTTP 503 SERVICE UNAVAILABLE)\n"
							+ "=" * 60 + "\n\n"
							f"Serwery usługi {provName} są w tej chwili przeciążone lub chwilowo niedostępne.\n\n"
							"Co możesz zrobić:\n"
							"1. Odczekaj chwilę (ok. 15-30 sekund) i spróbuj ponownie.\n"
							"2. W ustawieniach dodatku możesz zmienić model AI (np. Gemini 1.5 Flash lub Gemini 2.0 Flash) albo przełączyć się na drugiego dostawcę (OpenAI ChatGPT).\n"
							"3. Upewnij się, że masz stabilne połączenie z internetem."
						)
					elif "429" in errStr:
						msg = (
							f"PRZEKROCZONO LIMIT ZAPYTAŃ DLA {provName.upper()} (HTTP 429 TOO MANY REQUESTS)\n"
							+ "=" * 60 + "\n\n"
							"Wyczerpano dopuszczalny limit zapytań na minutę dla Twojego klucza API.\n"
							"Odczekaj minutę przed ponowieniem próby."
						)
					elif "401" in errStr or "403" in errStr:
						msg = (
							f"BŁĄD AUTORYZACJI KLUCZA API DLA {provName.upper()} (HTTP 401 / 403)\n"
							+ "=" * 60 + "\n\n"
							"Podany klucz API jest nieprawidłowy lub utracił ważność.\n"
							"Sprawdź i zaktualizuj klucz API w ustawieniach dodatku."
						)
					elif '400' in errStr or 'Bad Request' in errStr:
						msg = (
							f'BŁĄD ŻĄDANIA DO {provName.upper()} (HTTP 400 BAD REQUEST)\n'
							+ '=' * 60 + '\n\n'
							'Serwer sztucznej inteligencji odrzucił przesłany materiał.\n\n'
							f'Szczegóły błędu:\n{errStr}\n\n'
							'Możliwe przyczyny:\n'
							'1. Plik PDF jest zaszyfrowany lub zabezpieczony hasłem.\n'
							'2. Format pliku jest uszkodzony lub nieobsługiwany przez model AI.\n'
							'3. Rozmiar lub zawartość załącznika przekracza limity zapytania.'
						)
					else:
						msg = f"Wystąpił błąd podczas analizy materiału przez {provName}:\n\n{errStr}"
					showAccessibleMessageBox(msg, title=f"Błąd analizy AI ({provName})", parent=parentDialog)

				wx.CallAfter(onFail)
			finally:
				if tempPng and os.path.isfile(tempPng):
					try:
						os.remove(tempPng)
					except Exception:
						pass

		t = threading.Thread(target=worker, daemon=True)
		t.start()

	def triggerBatchAi(self, items: list[GraphicItem | AttachmentItem], parentDialog: GraphicElementsDialog | None = None):
		"""Zbiorcze opisywanie przez AI dla listy zaznaczonych grafik i załączników."""
		prov = getActiveAiProvider()
		provName = "OpenAI (ChatGPT)" if prov == "openai" else "Google Gemini"
		apiKey = getActiveAiApiKey()
		if not apiKey:
			def onReadyBatch():
				if GlobalPlugin.instance:
					GlobalPlugin.instance.triggerBatchAi(items, parentDialog=parentDialog)
			checkAndPromptAiApiKey(parentWin=parentDialog, onKeyReadyCallback=onReadyBatch)
			return

		if not items:
			ui.message("Nie wybrano żadnych elementów do analizy AI.")
			return

		ui.message(f"Wysyłanie {len(items)} elementów do sztucznej inteligencji ({provName})...")

		# Pre-ekstrakcja załączników i grafik na wątku głównym STA
		prepared_items = []
		for it in items:
			try:
				if isinstance(it, AttachmentItem):
					p = saveAttachmentToTempFile(it)
					if it.isPdf and p:
						unlocked = ensurePdfUnlocked(it, parentWin=parentDialog)
						if not unlocked:
							ui.message(f"Pominięto zaszyfrowany plik PDF: {it.fileName}")
							p = None
						else:
							p = unlocked
					prepared_items.append((it, p))
				else:
					p = extractGraphicToPngFile(it)
					prepared_items.append((it, p))
			except Exception as ex:
				log.debug(f"outlookImageHelper: błąd pre-ekstrakcji w triggerBatchAi: {ex}")
				prepared_items.append((it, None))

		def worker():
			results = []

			for idx, (it, targetPath) in enumerate(prepared_items):
				isAtt = isinstance(it, AttachmentItem)
				tempClean = None
				try:
					if isAtt:
						mimeType = "image/png"
						prefaceText = ""
						sendPath = targetPath
						b64Data = getattr(it, "base64Data", "")

						if it.isPdf:
							targetPath = getattr(it, 'unlockedPdfPath', '') or targetPath
							sendPath = targetPath
							mimeType = "application/pdf"
							try:
								with open(targetPath, "rb") as f:
									it.base64Data = base64.b64encode(f.read()).decode("ascii")
									b64Data = it.base64Data
							except Exception:
								pass
						elif it.isImage:
							mimeType = "image/png" if it.ext == ".png" else "image/jpeg"
							try:
								with open(targetPath, "rb") as f:
									it.base64Data = base64.b64encode(f.read()).decode("ascii")
									b64Data = it.base64Data
							except Exception:
								pass
						elif it.isDocx and targetPath:
							sendPath = ""
							docxText = extractTextFromDocxFile(targetPath)
							prefaceText = f"Zawartość dokumentu Word ({it.fileName}):\n\n{docxText}"
						elif it.isText and targetPath:
							sendPath = ""
							with open(targetPath, "r", encoding="utf-8", errors="replace") as f:
								prefaceText = f"Zawartość pliku ({it.fileName}):\n\n{f.read(40000)}"
						else:
							sendPath = ""
							prefaceText = f"Załącznik: {it.fileName} ({it.formattedSize})"

						fullPrompt = f"{prefaceText}\n\nOpisz ten załącznik szczegółowo dla osoby niewidomej." if prefaceText else "Opisz ten załącznik szczegółowo dla osoby niewidomej."
						desc = callAiVisionApi(
							prompt=fullPrompt,
							imagePath=sendPath,
							base64Data=b64Data,
							mimeType=mimeType,
							parentDialog=parentDialog,
						)
						it.aiDescription = desc
						b64 = getattr(it, "base64Data", "")
						uTurn = {"role": "user", "text": fullPrompt or "(Szczegółowa analiza załącznika)", "timestamp": time.strftime("%H:%M")}
						if b64:
							uTurn["imageData"] = b64
							uTurn["mimeType"] = mimeType
						it.chatHistory = [uTurn, {"role": "model", "text": desc, "timestamp": time.strftime("%H:%M")}]

						header = f"=== Opis AI — Załącznik {it.index}: {it.fileName} ({it.formattedSize}) ==="
						results.append(f"{header}\n\n{desc}")
					else:
						tempPng = targetPath
						tempClean = tempPng
						if tempPng and os.path.isfile(tempPng):
							try:
								with open(tempPng, "rb") as f:
									it.base64Image = base64.b64encode(f.read()).decode("ascii")
							except Exception:
								pass
							prompt = "Opisz ten obraz wyczerpująco i dokładnie dla osoby niewidomej."
							if it.contextText:
								prompt = f"Kontekst z treści e-maila: {it.contextText}\n\n{prompt}"
							desc = callAiVisionApi(
								prompt=prompt,
								imagePath=tempPng,
								base64Data=it.base64Image,
								mimeType="image/png",
								parentDialog=parentDialog,
							)
							it.aiDescription = desc
							it.chatHistory = [
								{
									"role": "user",
									"text": "(Szczegółowy opis)",
									"imageData": it.base64Image,
									"timestamp": time.strftime("%H:%M"),
								},
								{
									"role": "model",
									"text": desc,
									"timestamp": time.strftime("%H:%M"),
								},
							]
							header = f"=== Opis AI — Grafika {it.index}"
							if it.contextText:
								header += f" [{it.contextText}]"
							header += f" ({it.widthPx}×{it.heightPx} px) ==="
							results.append(f"{header}\n\n{desc}")
						else:
							results.append(f"=== Grafika {it.index} ===\nNie udało się pobrać obrazu.")
				except Exception as ex:
					lbl = f"Załącznik {it.fileName}" if isAtt else f"Grafika {it.index}"
					results.append(f"=== {lbl} ===\nBłąd AI: {ex}")
				finally:
					if tempClean and os.path.isfile(tempClean):
						try:
							os.remove(tempClean)
						except Exception:
							pass

			combined = ("\n\n" + "=" * 60 + "\n\n").join(results)

			def showDone():
				if parentDialog and winUser.isWindow(getattr(parentDialog, "Handle", 0)):
					parentDialog.refreshListLabels()
				ui.message(f"Zakończono analizę AI dla {len(items)} elementów.")
				ui.browseableMessage(combined, title=f"Zbiorcze opisy AI ({len(items)} elementów)")

			wx.CallAfter(showDone)

		t = threading.Thread(target=worker, daemon=True)
		t.start()
