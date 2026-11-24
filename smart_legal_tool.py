import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
from tkinter import filedialog, messagebox
import customtkinter as ctk
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

try:
  import fitz  # PyMuPDF

  PDF_OK = True
except Exception:
  PDF_OK = False

try:
  from PIL import Image, ImageDraw, ImageFont

  PIL_OK = True
except Exception:
  PIL_OK = False

try:
  from odf import text as odf_text
  from odf.opendocument import OpenDocumentText, load as load_odt

  ODT_OK = True
except Exception:
  ODT_OK = False

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


TRANSLATIONS = {
    "gu": {
        "lang_label": "ભાષા:",
        "total_fmt": "• કુલ ફોર્મેટ: {n}",
        "add_fmt": "+ ફોર્મેટ એડ કરો",
        "dl_all": "બધા ફોર્મેટ ડાઉનલોડ કરો",
        "del_fmt": "ફોર્મેટ કાઢો",
        "guide_btn": "ગાઇડ / મદદ",
        "search_ph": "ફોર્મેટ શોધો...",
        "case_no": "કેસ નંબર:",
        "case_no_ph": "કેસ નંબર...",
        "ps": "પોલીસ સ્ટેશન:",
        "ps_ph": "પોલીસ સ્ટેશન...",
        "fir": "FIR નંબર:",
        "fir_ph": "FIR નંબર...",
        "font_small": "અ-",
        "font_large": "અ+",
        "align_left": "ડાબે",
        "align_center": "મધ્યે",
        "align_right": "જમણે",
        "insert_line": "+ લીટી (____)",
        "editor_title": (
            "ફોર્મેટ એડિટર (નીચે સીધો સુધારો કરો — મૂળ ફાઇલ અકબંધ રહેશે)"
        ),
        "out_format": "આઉટપુટ ફોર્મેટ:",
        "save_as_btn": "સેવ કરો",
        "generate": "જનરેટ કરો",
        "save_dialog_title": "ફાઇલ ક્યાં સેવ કરવી છે તે પસંદ કરો",
        "no_match": "મેચ થતું કોઈ ફોર્મેટ મળ્યું નથી.",
        "select_fmt_first": "કૃપા કરીને પહેલા ફોર્મેટ પસંદ કરો.",
        "saved_status": "સેવ થયું: {path}",
        "missing_deps_note": (
            " (PDF/JPG/JPEG માટે 'pip install pymupdf pillow' ઇન્સ્ટોલ કરો)"
        ),
        "add_dialog_title": "ફોર્મેટ સેવ કરો",
        "add_dialog_q": "આ ફોર્મેટ કયા નામે સેવ કરવું છે?",
        "enter_fmt_name": "કૃપા કરીને ફોર્મેટનું નામ લખો.",
        "valid_name_err": "માન્ય નામ લખો.",
        "exists_err": "'{fname}' નામે ફોર્મેટ પહેલાથી હાજર છે! બીજું નામ લખો.",
        "fmt_saved_msg": "'{fname}' સેવ થઈ ગયું છે.",
        "save_fmt_btn": "ફોર્મેટ સેવ કરો",
        "del_dialog_title": "ફોર્મેટ કાઢો",
        "del_dialog_q": "જે ફોર્મેટ કાઢવું હોય તે સર્ચ કરીને પસંદ કરો:",
        "del_search_ph": "ફોર્મેટનું નામ શોધો...",
        "sel_fmt_none": "પસંદ કરેલ ફોર્મેટ: કોઈ નહીં",
        "sel_fmt_val": "પસંદ કરેલ ફોર્મેટ: {name}",
        "del_sel_btn": "પસંદ કરેલ ફોર્મેટ કાઢો",
        "no_fmt_to_del": "કાઢવા માટે કોઈ ફોર્મેટ ઉપલબ્ધ નથી.",
        "pick_fmt_warn": "કૃપા કરીને લિસ્ટમાંથી જે ફોર્મેટ કાઢવું હોય તે પસંદ કરો.",
        "confirm_title": "ખાતરી કરો",
        "confirm_del_msg": "શું તમે '{target}' ફોર્મેટ કાયમ માટે દૂર કરવા માંગો છો?",
        "deleted_status": "'{target}' દૂર કરવામાં આવ્યું.",
        "info_title": "માહિતી",
        "no_fmt_dl": "ડાઉનલોડ કરવા માટે કોઈ ફોર્મેટ નથી.",
        "pick_dl_dir": "બધા ફોર્મેટ ડાઉનલોડ કરવા માટે ફોલ્ડર પસંદ કરો",
        "dl_done_msg": "{n} ફોર્મેટ '{folder}' માં ડાઉનલોડ થઈ ગયા.",
        "fmt_type_title": "ફોર્મેટ",
        "fmt_type_warn": "કૃપા કરીને .docx, .odt અથવા .pdf ફાઇલ પસંદ કરો.",
        "open_dialog_title": "નવું ફોર્મેટ પસંદ કરો",
        "guide_title": "CourtWale — ઉપયોગની ગાઇડલાઇન",
        "close_btn": "બંધ કરો",
        "guide_body": (
            "૧. ભાષા પસંદગી:\n"
            "   ટોપ બારમાંથી ગુજરાતી, હિન્દી અથવા અંગ્રેજી ભાષા પસંદ કરો.\n\n"
            "૨. ફોર્મેટ શોધવું અને એડિટ કરવું:\n"
            "   સર્ચ બોક્સમાં નામ લખી ફોર્મેટ પસંદ કરો. Ctrl+A થી આખું લખાણ\n"
            "   સિલેક્ટ થશે. ડાબે / મધ્યે / જમણે ગોઠવણી, અ- / અ+ સાઇઝ અને\n"
            "   '+ લીટી (____)' વાપરી શકો છો.\n\n"
            "૩. સેવ અને જનરેટ:\n"
            "   નીચેથી DOCX/ODT, DOC, PDF, JPEG અથવા JPG પસંદ કરો.\n"
            "   • 'સેવ કરો': તમને જોઈએ તે ફોલ્ડર અને નામથી ફાઇલ સેવ કરો.\n"
            "   • 'જનરેટ કરો': કેસ નંબર / FIR મુજબના ફોલ્ડરમાં સીધી ફાઇલ બનાવો."
        ),
    },
    "hi": {
        "lang_label": "भाषा:",
        "total_fmt": "• कुल फ़ॉर्मेट: {n}",
        "add_fmt": "+ फ़ॉर्मेट जोड़ें",
        "dl_all": "सभी फ़ॉर्मेट डाउनलोड करें",
        "del_fmt": "फ़ॉर्मेट हटाएं",
        "guide_btn": "गाइड / मदद",
        "search_ph": "फ़ॉर्मेट खोजें...",
        "case_no": "केस नंबर:",
        "case_no_ph": "केस नंबर...",
        "ps": "पुलिस स्टेशन:",
        "ps_ph": "पुलिस स्टेशन...",
        "fir": "FIR नंबर:",
        "fir_ph": "FIR नंबर...",
        "font_small": "अ-",
        "font_large": "अ+",
        "align_left": "बाएं",
        "align_center": "मध्य",
        "align_right": "दाएं",
        "insert_line": "+ लाइन (____)",
        "editor_title": (
            "फ़ॉर्मेट एडिटर (नीचे सीधे सुधार करें — मूल फ़ाइल सुरक्षित रहेगी)"
        ),
        "out_format": "आउटपुट फ़ॉर्मेट:",
        "save_as_btn": "सेव करें",
        "generate": "जनरेट करें",
        "save_dialog_title": "फ़ाइल सेव करने का स्थान चुनें",
        "no_match": "कोई मेल खाता फ़ॉर्मेट नहीं मिला।",
        "select_fmt_first": "कृपया पहले फ़ॉर्मेट चुनें।",
        "saved_status": "सहेजा गया: {path}",
        "missing_deps_note": (
            " (PDF/JPG/JPEG के लिए 'pip install pymupdf pillow' इंस्टॉल करें)"
        ),
        "add_dialog_title": "फ़ॉर्मेट सहेजें",
        "add_dialog_q": "इस फ़ॉर्मेट को किस नाम से सहेजना है?",
        "enter_fmt_name": "कृपया फ़ॉर्मेट का नाम लिखें।",
        "valid_name_err": "मान्य नाम लिखें।",
        "exists_err": "'{fname}' नाम का फ़ॉर्मेट पहले से मौजूद है! दूसरा नाम लिखें।",
        "fmt_saved_msg": "'{fname}' सहेजा गया है।",
        "save_fmt_btn": "फ़ॉर्मेट सहेजें",
        "del_dialog_title": "फ़ॉर्मेट हटाएं",
        "del_dialog_q": "जो फ़ॉर्मेट हटाना हो उसे खोजकर चुनें:",
        "del_search_ph": "फ़ॉर्मेट का नाम खोजें...",
        "sel_fmt_none": "चयनित फ़ॉर्मेट: कोई नहीं",
        "sel_fmt_val": "चयनित फ़ॉर्मेट: {name}",
        "del_sel_btn": "चयनित फ़ॉर्मेट हटाएं",
        "no_fmt_to_del": "हटाने के लिए कोई फ़ॉर्मेट उपलब्ध नहीं है।",
        "pick_fmt_warn": "कृपया सूची से जो फ़ॉर्मेट हटाना हो उसे चुनें।",
        "confirm_title": "पुष्टि करें",
        "confirm_del_msg": (
            "क्या आप '{target}' फ़ॉर्मेट हमेशा के लिए हटाना चाहते हैं?"
        ),
        "deleted_status": "'{target}' हटा दिया गया।",
        "info_title": "सूचना",
        "no_fmt_dl": "डाउनलोड करने के लिए कोई फ़ॉर्मेट नहीं है।",
        "pick_dl_dir": "सभी फ़ॉर्मेट डाउनलोड करने के लिए फ़ोल्डर चुनें",
        "dl_done_msg": "{n} फ़ॉर्मेट '{folder}' में डाउनलोड हो गए।",
        "fmt_type_title": "फ़ॉर्मेट",
        "fmt_type_warn": "कृपया .docx, .odt या .pdf फ़ाइल चुनें।",
        "open_dialog_title": "नया फ़ॉर्मेट चुनें",
        "guide_title": "CourtWale — उपयोग गाइडलाइन",
        "close_btn": "बंद करें",
        "guide_body": (
            "१. भाषा चयन:\n"
            "   शीर्ष बार से गुजराती, हिन्दी या अंग्रेज़ी भाषा चुनें।\n\n"
            "२. फ़ॉर्मेट खोजना और संपादित करना:\n"
            "   सर्च बॉक्स में नाम लिखकर फ़ॉर्मेट चुनें। Ctrl+A से पूरा टेक्स्ट\n"
            "   चयनित होगा। बाएं / मध्य / दाएं संरेखण, अ- / अ+ साइज़ और\n"
            "   '+ लाइन (____)' का उपयोग करें।\n\n"
            "३. सेव और जनरेट:\n"
            "   नीचे से DOCX/ODT, DOC, PDF, JPEG या JPG चुनें।\n"
            "   • 'सेव करें': अपनी पसंद के फ़ोल्डर और नाम से फ़ाइल सहेजें।\n"
            "   • 'जनरेट करें': केस नंबर / FIR के अनुसार फ़ोल्डर में फ़ाइल बनाएं।"
        ),
    },
    "en": {
        "lang_label": "Language:",
        "total_fmt": "• Total Formats: {n}",
        "add_fmt": "+ Add Format",
        "dl_all": "Download All Formats",
        "del_fmt": "Remove Format",
        "guide_btn": "Guide / Help",
        "search_ph": "Search format...",
        "case_no": "Case No:",
        "case_no_ph": "Case No...",
        "ps": "Police Station:",
        "ps_ph": "Police Station...",
        "fir": "FIR No:",
        "fir_ph": "FIR No...",
        "font_small": "A-",
        "font_large": "A+",
        "align_left": "Left",
        "align_center": "Center",
        "align_right": "Right",
        "insert_line": "+ Line (____)",
        "editor_title": (
            "Format Editor (Edit directly below — original file stays intact)"
        ),
        "out_format": "Output Format:",
        "save_as_btn": "Save",
        "generate": "Generate",
        "save_dialog_title": "Choose where to save the file",
        "no_match": "No matching format found.",
        "select_fmt_first": "Please select a format first.",
        "saved_status": "Saved: {path}",
        "missing_deps_note": (
            " (Install 'pip install pymupdf pillow' for PDF/JPG/JPEG export)"
        ),
        "add_dialog_title": "Save Format",
        "add_dialog_q": "Under what name should this format be saved?",
        "enter_fmt_name": "Please enter a format name.",
        "valid_name_err": "Please enter a valid name.",
        "exists_err": "A format named '{fname}' already exists! Enter another name.",
        "fmt_saved_msg": "'{fname}' has been saved.",
        "save_fmt_btn": "Save Format",
        "del_dialog_title": "Remove Format",
        "del_dialog_q": "Search and select the format you want to remove:",
        "del_search_ph": "Search format name...",
        "sel_fmt_none": "Selected Format: None",
        "sel_fmt_val": "Selected Format: {name}",
        "del_sel_btn": "Remove Selected Format",
        "no_fmt_to_del": "No formats available to remove.",
        "pick_fmt_warn": "Please select a format from the list to remove.",
        "confirm_title": "Confirm",
        "confirm_del_msg": (
            "Are you sure you want to permanently remove '{target}'?"
        ),
        "deleted_status": "'{target}' removed.",
        "info_title": "Info",
        "no_fmt_dl": "No formats available to download.",
        "pick_dl_dir": "Select a folder to download all formats",
        "dl_done_msg": "{n} formats downloaded to '{folder}'.",
        "fmt_type_title": "Format",
        "fmt_type_warn": "Please select a .docx, .odt, or .pdf file.",
        "open_dialog_title": "Select New Format",
        "guide_title": "CourtWale — User Guideline",
        "close_btn": "Close",
        "guide_body": (
            "1. Language Selection:\n"
            "   Select Gujarati, Hindi, or English from the top bar.\n\n"
            "2. Search & Edit Format:\n"
            "   Search and pick a format. Ctrl+A selects all text in the editor.\n"
            "   Use Left/Center/Right alignment, A-/A+ font size, or '+ Line"
            " (____)'.\n\n"
            "3. Save & Generate:\n"
            "   Select DOCX/ODT, DOC, PDF, JPEG, or JPG below.\n"
            "   • 'Save': Pick any folder and path to save.\n"
            "   • 'Generate': Auto-save into Case No / FIR output folder."
        ),
    },
}


class WinUnicodeIME:
  """Windows user32.dll થી વર્તમાન કીબોર્ડ લેઆઉટ (ગુજરાતી / હિન્દી / Inscript / IME) મુજબ સાચો Unicode અક્ષર મેળવવા માટે."""

  def __init__(self):
    self.enabled = sys.platform == "win32"
    if self.enabled:
      try:
        self.user32 = ctypes.WinDLL("user32", use_last_error=True)
        self.user32.GetKeyboardLayout.argtypes = [ctypes.c_ulong]
        self.user32.GetKeyboardLayout.restype = ctypes.c_void_p

        self.user32.GetKeyboardState.argtypes = [ctypes.POINTER(ctypes.c_byte)]
        self.user32.GetKeyboardState.restype = ctypes.c_int

        self.user32.ToUnicodeEx.argtypes = [
            ctypes.c_uint,
            ctypes.c_uint,
            ctypes.POINTER(ctypes.c_byte),
            ctypes.c_wchar_p,
            ctypes.c_int,
            ctypes.c_uint,
            ctypes.c_void_p,
        ]
        self.user32.ToUnicodeEx.restype = ctypes.c_int
      except Exception:
        self.enabled = False

  def translate_event_to_unicode(self, event) -> str:
    if not self.enabled:
      return ""
    try:
      vk = int(getattr(event, "keycode", 0) or 0)
      if not vk:
        return ""
      key_state = (ctypes.c_byte * 256)()
      if not self.user32.GetKeyboardState(key_state):
        return ""

      state = int(getattr(event, "state", 0) or 0)
      if state & 0x0001:  # Shift
        key_state[0x10] = -128
        key_state[0xA0] = -128
      if state & 0x0004:  # Control
        key_state[0x11] = -128
      if state & 0x0008:  # Alt
        key_state[0x12] = -128

      hkl = self.user32.GetKeyboardLayout(0)
      buf = ctypes.create_unicode_buffer(8)
      rc = self.user32.ToUnicodeEx(vk, 0, key_state, buf, 8, 0, hkl)
      if rc > 0:
        out = buf.value[:rc]
        if out and any(ord(c) >= 0x0900 for c in out):
          return out
        if out and out != "?":
          return out
    except Exception:
      pass
    return ""


class CourtWaleCleanApp(ctk.CTk):

  def __init__(self):
    super().__init__()
    self.title("CourtWale")
    self.geometry("1060x780")
    self.minsize(860, 640)

    self.bg_page = "#f4f6f9"
    self.bg_card = "#ffffff"
    self.bg_soft = "#eef2f6"
    self.border_col = "#dce3ec"
    self.text_dark = "#1f2937"
    self.text_muted = "#64748b"
    self.accent_blue = "#1a73e8"
    self.accent_blue_hover = "#1557b0"
    self.accent_green = "#15803d"
    self.accent_green_hover = "#166534"
    self.accent_red = "#dc2626"
    self.accent_red_hover = "#b91c1c"
    self.hover_row = "#dbeafe"
    self.status_ok = "#15803d"
    self.status_warn = "#b45309"

    self.configure(fg_color=self.bg_page)

    if getattr(sys, "frozen", False):
      self.bundle_dir = sys._MEIPASS
      self.app_dir = os.path.dirname(sys.executable)
    else:
      self.bundle_dir = os.path.dirname(os.path.abspath(__file__))
      self.app_dir = self.bundle_dir

    icon_path = os.path.join(self.app_dir, "court_icon.ico")
    if not os.path.exists(icon_path):
      icon_path = os.path.join(self.bundle_dir, "court_icon.ico")
    if os.path.exists(icon_path):
      try:
        self.iconbitmap(icon_path)
      except Exception:
        pass

    self.editor_font_size = 15
    self.font_ui = ("Nirmala UI", 13)
    self.font_ui_bold = ("Nirmala UI", 13, "bold")
    self.font_title = ("Nirmala UI", 17, "bold")
    self.font_editor = ("Nirmala UI", self.editor_font_size)

    self.lang_config_path = os.path.join(self.app_dir, "lang_config.json")
    self.lang = self._load_saved_lang()

    self.templates_dir = os.path.join(self.app_dir, "Templates")
    self.output_dir = os.path.join(self.app_dir, "Generated_Output")
    os.makedirs(self.templates_dir, exist_ok=True)
    os.makedirs(self.output_dir, exist_ok=True)

    self._seed_builtin_templates()

    self.ime_helper = WinUnicodeIME()
    self.all_templates = []
    self.matched_templates = []
    self.current_template = None
    self._ignore_search_trace = False
    self.dropdown_win = None

    self._build_ui()
    self._init_editor_align_tags()
    self._bind_select_all_shortcuts()
    self.apply_language(self.lang)
    self.refresh_templates(select_first=True)

  def _bind_select_all_shortcuts(self):
    def select_all_in_textbox(event=None):
      tb = self.inline_editor._textbox
      tb.tag_remove("sel", "1.0", "end")
      tb.tag_add("sel", "1.0", "end-1c")
      tb.mark_set("insert", "end-1c")
      tb.see("insert")
      return "break"

    self.inline_editor._textbox.bind("<Control-a>", select_all_in_textbox)
    self.inline_editor._textbox.bind("<Control-A>", select_all_in_textbox)

  def _init_editor_align_tags(self):
    tb = self.inline_editor._textbox
    tb.tag_configure("align_left", justify="left")
    tb.tag_configure("align_center", justify="center")
    tb.tag_configure("align_right", justify="right")

  def _load_saved_lang(self) -> str:
    try:
      if os.path.exists(self.lang_config_path):
        with open(self.lang_config_path, "r", encoding="utf-8") as f:
          data = json.load(f)
          val = data.get("lang", "gu")
          if val in TRANSLATIONS:
            return val
    except Exception:
      pass
    return "gu"

  def _save_lang(self, lang_code: str):
    try:
      with open(self.lang_config_path, "w", encoding="utf-8") as f:
        json.dump({"lang": lang_code}, f)
    except Exception:
      pass

  def t(self, key: str, **kwargs) -> str:
    pack = TRANSLATIONS.get(self.lang, TRANSLATIONS["gu"])
    text = pack.get(key, TRANSLATIONS["gu"].get(key, key))
    if kwargs:
      return text.format(**kwargs)
    return text

  def _on_lang_selected(self, choice: str):
    mapping = {
        "ગુજરાતી": "gu",
        "हिन्दी": "hi",
        "English": "en",
    }
    code = mapping.get(choice, "gu")
    self.apply_language(code)

  def apply_language(self, lang_code: str):
    if lang_code not in TRANSLATIONS:
      lang_code = "gu"
    self.lang = lang_code
    self._save_lang(lang_code)

    reverse_map = {
        "gu": "ગુજરાતી",
        "hi": "हिन्दी",
        "en": "English",
    }
    self.lang_menu.set(reverse_map[self.lang])

    self.lang_lbl.configure(text=self.t("lang_label"))
    self.count_label.configure(
        text=self.t("total_fmt", n=len(self.all_templates))
    )
    self.btn_guide.configure(text=self.t("guide_btn"))
    self.btn_del_fmt.configure(text=self.t("del_fmt"))
    self.btn_dl_all.configure(text=self.t("dl_all"))
    self.btn_add_fmt.configure(text=self.t("add_fmt"))

    self.search_entry.configure(placeholder_text=self.t("search_ph"))
    self.case_no_lbl.configure(text=self.t("case_no"))
    self.case_no_entry.configure(placeholder_text=self.t("case_no_ph"))
    self.ps_lbl.configure(text=self.t("ps"))
    self.ps_entry.configure(placeholder_text=self.t("ps_ph"))
    self.fir_lbl.configure(text=self.t("fir"))
    self.fir_entry.configure(placeholder_text=self.t("fir_ph"))

    self.btn_font_dec.configure(text=self.t("font_small"))
    self.btn_font_inc.configure(text=self.t("font_large"))
    self.btn_align_left.configure(text=self.t("align_left"))
    self.btn_align_center.configure(text=self.t("align_center"))
    self.btn_align_right.configure(text=self.t("align_right"))
    self.btn_insert_line.configure(text=self.t("insert_line"))

    self.lower_title_lbl.configure(text=self.t("editor_title"))
    self.out_fmt_lbl.configure(text=self.t("out_format"))
    self.btn_save_path.configure(text=self.t("save_as_btn"))
    self.btn_generate.configure(text=self.t("generate"))

  def _seed_builtin_templates(self):
    bundled_tpl_dir = os.path.join(self.bundle_dir, "Templates")
    if os.path.abspath(bundled_tpl_dir) == os.path.abspath(self.templates_dir):
      return
    if not os.path.exists(bundled_tpl_dir):
      return
    for fname in os.listdir(bundled_tpl_dir):
      if fname.lower().endswith(
          (".docx", ".odt", ".pdf")
      ) and not fname.startswith("~$"):
        src = os.path.join(bundled_tpl_dir, fname)
        dst = os.path.join(self.templates_dir, fname)
        if not os.path.exists(dst):
          try:
            shutil.copy2(src, dst)
          except Exception:
            pass

  def _bind_unicode_input(self, widget, insert_func):
    def on_key(event):
      if event.state & (0x4 | 0x8):
        return None
      if event.char == "?":
        fixed = self.ime_helper.translate_event_to_unicode(event)
        if fixed and fixed != "?":
          insert_func(fixed)
          return "break"
      return None

    widget.bind("<Key>", on_key, add="+")

  def _build_ui(self):
    top_bar = ctk.CTkFrame(
        self,
        fg_color=self.bg_card,
        corner_radius=14,
        border_width=1,
        border_color=self.border_col,
    )
    top_bar.pack(fill="x", padx=14, pady=(8, 4))

    ctk.CTkLabel(
        top_bar,
        text="CourtWale",
        font=self.font_title,
        text_color=self.accent_blue,
    ).pack(side="left", padx=(16, 8), pady=7)

    self.count_label = ctk.CTkLabel(
        top_bar,
        text="",
        font=("Nirmala UI", 12),
        text_color=self.text_muted,
    )
    self.count_label.pack(side="left", padx=4)

    self.lang_lbl = ctk.CTkLabel(
        top_bar,
        text="",
        font=("Nirmala UI", 12, "bold"),
        text_color=self.text_dark,
    )
    self.lang_lbl.pack(side="left", padx=(10, 2), pady=7)

    self.lang_menu = ctk.CTkOptionMenu(
        top_bar,
        values=["ગુજરાતી", "हिन्दी", "English"],
        command=self._on_lang_selected,
        font=("Nirmala UI", 12, "bold"),
        fg_color=self.bg_soft,
        button_color=self.border_col,
        button_hover_color=self.hover_row,
        text_color=self.text_dark,
        dropdown_fg_color=self.bg_card,
        dropdown_text_color=self.text_dark,
        dropdown_hover_color=self.hover_row,
        corner_radius=14,
        height=30,
        width=110,
    )
    self.lang_menu.pack(side="left", padx=(4, 6), pady=7)

    self.btn_guide = ctk.CTkButton(
        top_bar,
        text="",
        command=self.open_guide_dialog,
        font=("Nirmala UI", 12, "bold"),
        fg_color=self.bg_soft,
        hover_color=self.hover_row,
        text_color=self.text_dark,
        border_width=1,
        border_color=self.border_col,
        corner_radius=16,
        height=32,
        width=105,
    )
    self.btn_guide.pack(side="left", padx=4, pady=7)

    self.btn_del_fmt = ctk.CTkButton(
        top_bar,
        text="",
        command=self.open_remove_format_dialog,
        font=self.font_ui_bold,
        fg_color=self.accent_red,
        hover_color=self.accent_red_hover,
        text_color="#ffffff",
        corner_radius=18,
        height=34,
        width=130,
    )
    self.btn_del_fmt.pack(side="right", padx=(4, 14), pady=7)

    self.btn_dl_all = ctk.CTkButton(
        top_bar,
        text="",
        command=self.download_all_templates,
        font=self.font_ui_bold,
        fg_color=self.accent_green,
        hover_color=self.accent_green_hover,
        text_color="#ffffff",
        corner_radius=18,
        height=34,
        width=185,
    )
    self.btn_dl_all.pack(side="right", padx=4, pady=7)

    self.btn_add_fmt = ctk.CTkButton(
        top_bar,
        text="",
        command=self.add_new_format_flow,
        font=self.font_ui_bold,
        fg_color=self.accent_blue,
        hover_color=self.accent_blue_hover,
        text_color="#ffffff",
        corner_radius=18,
        height=34,
    )
    self.btn_add_fmt.pack(side="right", padx=4, pady=7)

    ctrl_card = ctk.CTkFrame(
        self,
        fg_color=self.bg_card,
        corner_radius=14,
        border_width=1,
        border_color=self.border_col,
    )
    ctrl_card.pack(fill="x", padx=14, pady=(0, 4))

    search_row = ctk.CTkFrame(ctrl_card, fg_color="transparent")
    search_row.pack(fill="x", padx=12, pady=(7, 5))

    self.search_var = ctk.StringVar()
    self.search_var.trace_add("write", self._on_search_change)

    self.search_entry = ctk.CTkEntry(
        search_row,
        textvariable=self.search_var,
        height=34,
        font=self.font_ui,
        fg_color=self.bg_soft,
        border_color=self.border_col,
        text_color=self.text_dark,
        corner_radius=17,
    )
    self.search_entry.pack(fill="x", expand=True)
    self.search_entry.bind("<KeyRelease>", self._on_key_release)
    self.search_entry.bind("<Tab>", self._on_tab_accept)
    self.search_entry.bind("<Return>", self._on_return_open)
    self.search_entry.bind("<Down>", self._focus_dropdown_list)

    self._bind_unicode_input(
        self.search_entry._entry,
        lambda txt: self.search_entry._entry.insert("insert", txt),
    )

    meta_row = ctk.CTkFrame(ctrl_card, fg_color="transparent")
    meta_row.pack(fill="x", padx=12, pady=(0, 7))

    self.case_no_lbl = ctk.CTkLabel(
        meta_row,
        text="",
        font=("Nirmala UI", 12, "bold"),
        text_color=self.text_dark,
    )
    self.case_no_lbl.pack(side="left", padx=(2, 4))

    self.case_no_entry = ctk.CTkEntry(
        meta_row,
        width=130,
        height=30,
        font=("Nirmala UI", 12),
        fg_color=self.bg_soft,
        border_color=self.border_col,
        text_color=self.text_dark,
    )
    self.case_no_entry.pack(side="left", padx=(0, 8))
    self._bind_unicode_input(
        self.case_no_entry._entry,
        lambda txt: self.case_no_entry._entry.insert("insert", txt),
    )

    self.ps_lbl = ctk.CTkLabel(
        meta_row,
        text="",
        font=("Nirmala UI", 12, "bold"),
        text_color=self.text_dark,
    )
    self.ps_lbl.pack(side="left", padx=(2, 4))

    self.ps_entry = ctk.CTkEntry(
        meta_row,
        width=145,
        height=30,
        font=("Nirmala UI", 12),
        fg_color=self.bg_soft,
        border_color=self.border_col,
        text_color=self.text_dark,
    )
    self.ps_entry.pack(side="left", padx=(0, 8))
    self._bind_unicode_input(
        self.ps_entry._entry,
        lambda txt: self.ps_entry._entry.insert("insert", txt),
    )

    self.fir_lbl = ctk.CTkLabel(
        meta_row,
        text="",
        font=("Nirmala UI", 12, "bold"),
        text_color=self.text_dark,
    )
    self.fir_lbl.pack(side="left", padx=(2, 4))

    self.fir_entry = ctk.CTkEntry(
        meta_row,
        width=130,
        height=30,
        font=("Nirmala UI", 12),
        fg_color=self.bg_soft,
        border_color=self.border_col,
        text_color=self.text_dark,
    )
    self.fir_entry.pack(side="left", padx=(0, 8))
    self._bind_unicode_input(
        self.fir_entry._entry,
        lambda txt: self.fir_entry._entry.insert("insert", txt),
    )

    self.btn_insert_line = ctk.CTkButton(
        meta_row,
        text="",
        width=92,
        height=28,
        font=("Nirmala UI", 11, "bold"),
        fg_color=self.bg_soft,
        hover_color=self.hover_row,
        text_color=self.accent_blue,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        command=self._insert_blank_line,
    )
    self.btn_insert_line.pack(side="right", padx=2)

    self.btn_align_right = ctk.CTkButton(
        meta_row,
        text="",
        width=52,
        height=28,
        font=("Nirmala UI", 11),
        fg_color=self.bg_soft,
        hover_color=self.hover_row,
        text_color=self.text_dark,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        command=lambda: self._set_editor_align("right"),
    )
    self.btn_align_right.pack(side="right", padx=2)

    self.btn_align_center = ctk.CTkButton(
        meta_row,
        text="",
        width=52,
        height=28,
        font=("Nirmala UI", 11),
        fg_color=self.bg_soft,
        hover_color=self.hover_row,
        text_color=self.text_dark,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        command=lambda: self._set_editor_align("center"),
    )
    self.btn_align_center.pack(side="right", padx=2)

    self.btn_align_left = ctk.CTkButton(
        meta_row,
        text="",
        width=50,
        height=28,
        font=("Nirmala UI", 11),
        fg_color=self.bg_soft,
        hover_color=self.hover_row,
        text_color=self.text_dark,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        command=lambda: self._set_editor_align("left"),
    )
    self.btn_align_left.pack(side="right", padx=2)

    self.btn_font_inc = ctk.CTkButton(
        meta_row,
        text="",
        width=34,
        height=28,
        font=("Nirmala UI", 11, "bold"),
        fg_color=self.bg_soft,
        hover_color=self.hover_row,
        text_color=self.text_dark,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        command=self._increase_editor_font,
    )
    self.btn_font_inc.pack(side="right", padx=2)

    self.btn_font_dec = ctk.CTkButton(
        meta_row,
        text="",
        width=34,
        height=28,
        font=("Nirmala UI", 11, "bold"),
        fg_color=self.bg_soft,
        hover_color=self.hover_row,
        text_color=self.text_dark,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        command=self._decrease_editor_font,
    )
    self.btn_font_dec.pack(side="right", padx=2)

    lower_card = ctk.CTkFrame(
        self,
        fg_color=self.bg_card,
        corner_radius=14,
        border_width=1,
        border_color=self.border_col,
    )
    lower_card.pack(fill="both", expand=True, padx=14, pady=(0, 4))

    lower_header = ctk.CTkFrame(lower_card, fg_color="transparent")
    lower_header.pack(fill="x", padx=14, pady=(5, 2))

    self.lower_title_lbl = ctk.CTkLabel(
        lower_header,
        text="",
        font=self.font_ui_bold,
        text_color=self.accent_blue,
    )
    self.lower_title_lbl.pack(side="left")

    self.inline_editor = ctk.CTkTextbox(
        lower_card,
        font=self.font_editor,
        fg_color="#ffffff",
        text_color="#111827",
        border_width=1,
        border_color=self.border_col,
        corner_radius=12,
        wrap="word",
        undo=True,
    )
    self.inline_editor.pack(fill="both", expand=True, padx=12, pady=(2, 8))
    self._bind_unicode_input(
        self.inline_editor._textbox,
        lambda txt: self.inline_editor._textbox.insert("insert", txt),
    )
    self.inline_editor._textbox.bind(
        "<FocusIn>", lambda e: self._close_dropdown(), add="+"
    )

    bottom_bar = ctk.CTkFrame(self, fg_color="transparent")
    bottom_bar.pack(fill="x", padx=14, pady=(0, 2))

    action_line = ctk.CTkFrame(bottom_bar, fg_color="transparent")
    action_line.pack(fill="x")

    self.out_fmt_lbl = ctk.CTkLabel(
        action_line,
        text="",
        font=("Nirmala UI", 12, "bold"),
        text_color=self.text_dark,
    )
    self.out_fmt_lbl.pack(side="left", padx=(4, 8))

    self.out_fmt_menu = ctk.CTkOptionMenu(
        action_line,
        values=["DOCX / ODT", "DOC", "PDF", "JPEG", "JPG"],
        font=("Nirmala UI", 12, "bold"),
        fg_color=self.bg_card,
        button_color=self.border_col,
        button_hover_color=self.hover_row,
        text_color=self.text_dark,
        dropdown_fg_color=self.bg_card,
        dropdown_text_color=self.text_dark,
        dropdown_hover_color=self.hover_row,
        corner_radius=14,
        height=34,
        width=145,
    )
    self.out_fmt_menu.pack(side="left", padx=(0, 10))

    self.btn_save_path = ctk.CTkButton(
        action_line,
        text="",
        command=self.save_file_to_chosen_path,
        font=("Nirmala UI", 13, "bold"),
        fg_color=self.accent_green,
        hover_color=self.accent_green_hover,
        text_color="#ffffff",
        corner_radius=18,
        height=36,
        width=140,
    )
    self.btn_save_path.pack(side="left")

    self.btn_generate = ctk.CTkButton(
        action_line,
        text="",
        command=self.generate_file,
        font=("Nirmala UI", 14, "bold"),
        fg_color=self.accent_blue,
        hover_color=self.accent_blue_hover,
        text_color="#ffffff",
        corner_radius=18,
        height=36,
        width=200,
    )
    self.btn_generate.pack(side="right")

    self.status_lbl = ctk.CTkLabel(
        bottom_bar,
        text="",
        font=("Nirmala UI", 11),
        text_color=self.text_muted,
    )
    self.status_lbl.pack(pady=(1, 0), anchor="w")

  def _decrease_editor_font(self):
    if self.editor_font_size > 11:
      self.editor_font_size -= 1
      self.font_editor = ("Nirmala UI", self.editor_font_size)
      self.inline_editor.configure(font=self.font_editor)

  def _increase_editor_font(self):
    if self.editor_font_size < 28:
      self.editor_font_size += 1
      self.font_editor = ("Nirmala UI", self.editor_font_size)
      self.inline_editor.configure(font=self.font_editor)

  def _apply_line_align(self, start_line: int, end_line: int, align_mode: str):
    tb = self.inline_editor._textbox
    for tag in ("align_left", "align_center", "align_right"):
      tb.tag_remove(tag, f"{start_line}.0", f"{end_line}.end")
    tb.tag_add(f"align_{align_mode}", f"{start_line}.0", f"{end_line}.end")

  def _set_editor_align(self, align_mode: str):
    tb = self.inline_editor._textbox
    try:
      sel_start = tb.index("sel.first")
      sel_end = tb.index("sel.last")
      s_line = int(sel_start.split(".")[0])
      e_line = int(sel_end.split(".")[0])
      self._apply_line_align(s_line, e_line, align_mode)
    except Exception:
      cur_line = int(tb.index("insert").split(".")[0])
      self._apply_line_align(cur_line, cur_line, align_mode)

  def _insert_blank_line(self):
    self.inline_editor._textbox.insert("insert", "________________________")

  def open_guide_dialog(self):
    self._close_dropdown()
    win = ctk.CTkToplevel(self)
    win.title(self.t("guide_title"))
    win.geometry("540x430")
    win.minsize(460, 380)
    win.configure(fg_color=self.bg_card)
    win.grab_set()

    ctk.CTkLabel(
        win,
        text=self.t("guide_title"),
        font=self.font_title,
        text_color=self.accent_blue,
    ).pack(pady=(14, 6), padx=16, anchor="w")

    txt_box = ctk.CTkTextbox(
        win,
        font=self.font_ui,
        fg_color=self.bg_soft,
        text_color=self.text_dark,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        wrap="word",
    )
    txt_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))
    txt_box.insert("1.0", self.t("guide_body"))
    txt_box.configure(state="disabled")

    ctk.CTkButton(
        win,
        text=self.t("close_btn"),
        command=win.destroy,
        font=self.font_ui_bold,
        fg_color=self.accent_blue,
        hover_color=self.accent_blue_hover,
        text_color="#ffffff",
        height=36,
        corner_radius=16,
    ).pack(padx=16, pady=(0, 14), fill="x")

  def _close_dropdown(self, event=None):
    if self.dropdown_win and self.dropdown_win.winfo_exists():
      self.dropdown_win.destroy()
    self.dropdown_win = None

  def _open_or_update_dropdown(self, matches: list):
    self.update_idletasks()
    x = self.search_entry.winfo_rootx()
    y = self.search_entry.winfo_rooty() + self.search_entry.winfo_height() + 2
    w = self.search_entry.winfo_width()

    if self.dropdown_win and self.dropdown_win.winfo_exists():
      self.dropdown_win.destroy()

    pop = ctk.CTkToplevel(self)
    self.dropdown_win = pop
    pop.overrideredirect(True)
    pop.configure(fg_color=self.bg_card)

    list_h = min(220, max(44, len(matches) * 34 + 10))
    pop.geometry(f"{w}x{list_h}+{x}+{y}")

    scroll = ctk.CTkScrollableFrame(
        pop,
        fg_color=self.bg_card,
        border_width=1,
        border_color=self.accent_blue,
        corner_radius=10,
    )
    scroll.pack(fill="both", expand=True)

    if not matches:
      ctk.CTkLabel(
          scroll,
          text=self.t("no_match"),
          font=self.font_ui,
          text_color=self.text_muted,
      ).pack(pady=8, padx=10, anchor="w")
      return

    for fname in matches:
      btn = ctk.CTkButton(
          scroll,
          text=fname,
          anchor="w",
          font=self.font_ui,
          fg_color="transparent",
          hover_color=self.hover_row,
          text_color=self.text_dark,
          height=30,
          corner_radius=6,
          command=lambda name=fname: self._select_from_dropdown(name),
      )
      btn.pack(fill="x", padx=4, pady=1)

  def _select_from_dropdown(self, filename: str):
    self._ignore_search_trace = True
    self.search_var.set(filename)
    self._ignore_search_trace = False
    self._close_dropdown()
    self.open_saved_template(filename)

  def _focus_dropdown_list(self, event=None):
    if self.dropdown_win and self.dropdown_win.winfo_exists():
      self.dropdown_win.focus_set()
    return "break"

  def _on_key_release(self, event):
    if event.keysym in (
        "BackSpace",
        "Delete",
        "Left",
        "Up",
        "Down",
        "Escape",
    ):
      if event.keysym == "Escape":
        self._close_dropdown()
      return

    typed = self.search_var.get()
    if not typed:
      self._close_dropdown()
      return

    typed_lower = typed.lower()
    prefix_matches = [
        f for f in self.all_templates if f.lower().startswith(typed_lower)
    ]
    if prefix_matches:
      best = prefix_matches[0]
      self._ignore_search_trace = True
      self.search_entry.delete(0, "end")
      self.search_entry.insert(0, best)
      self.search_entry.select_range(len(typed), len(best))
      self.search_entry.icursor(len(typed))
      self._ignore_search_trace = False

  def _on_search_change(self, *args):
    if self._ignore_search_trace:
      return

    typed = self.search_var.get().strip()
    if not typed:
      self.matched_templates = []
      self._close_dropdown()
      return

    q = typed.lower()
    starts = [f for f in self.all_templates if f.lower().startswith(q)]
    contains = [
        f for f in self.all_templates if q in f.lower() and f not in starts
    ]
    self.matched_templates = starts + contains

    self._open_or_update_dropdown(self.matched_templates)

    if self.matched_templates:
      self.open_saved_template(self.matched_templates[0])

  def _on_tab_accept(self, event=None):
    full_val = self.search_var.get()
    if full_val:
      self.search_entry.select_clear()
      self.search_entry.icursor("end")
      self._close_dropdown()
      if full_val in self.all_templates:
        self.open_saved_template(full_val)
    return "break"

  def _on_return_open(self, event=None):
    full_val = self.search_var.get()
    self.search_entry.select_clear()
    self.search_entry.icursor("end")
    self._close_dropdown()
    if full_val in self.all_templates:
      self.open_saved_template(full_val)
    elif self.matched_templates:
      self.open_saved_template(self.matched_templates[0])
    return "break"

  def refresh_templates(self, select_first=False, select_name=None):
    files = [
        f
        for f in os.listdir(self.templates_dir)
        if f.lower().endswith((".docx", ".odt", ".pdf"))
        and not f.startswith("~$")
    ]
    files.sort(key=str.lower)
    self.all_templates = files
    self.count_label.configure(
        text=self.t("total_fmt", n=len(self.all_templates))
    )

    if select_name and select_name in self.all_templates:
      self.open_saved_template(select_name)
    elif select_first and self.all_templates:
      self.open_saved_template(self.all_templates[0])
    elif not self.all_templates:
      self.current_template = None
      self._set_editor_blocks([("", "left")])

  def _map_docx_align(self, align_val) -> str:
    if align_val == WD_ALIGN_PARAGRAPH.CENTER:
      return "center"
    if align_val == WD_ALIGN_PARAGRAPH.RIGHT:
      return "right"
    return "left"

  def _set_editor_blocks(self, blocks: list):
    self.inline_editor.configure(state="normal")
    tb = self.inline_editor._textbox
    tb.delete("1.0", "end")
    for tag in ("align_left", "align_center", "align_right"):
      tb.tag_remove(tag, "1.0", "end")

    if not blocks:
      blocks = [("", "left")]

    for idx, (line_txt, align_mode) in enumerate(blocks, start=1):
      if idx > 1:
        tb.insert("end", "\n")
      line_start = f"{idx}.0"
      if line_txt:
        tb.insert("end", line_txt)
      line_end = f"{idx}.end"
      tb.tag_add(f"align_{align_mode}", line_start, line_end)

  def _sanitize_name(self, text):
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", text.strip())
    cleaned = re.sub(r"\s+", "_", cleaned).strip("_")
    return cleaned

  def add_new_format_flow(self):
    self._close_dropdown()
    filetypes = [
        ("Supported Formats (Word, Writer, PDF)", "*.docx *.odt *.pdf"),
        ("Word Document", "*.docx"),
        ("Ubuntu Writer", "*.odt"),
        ("PDF Document", "*.pdf"),
    ]
    src_path = filedialog.askopenfilename(
        title=self.t("open_dialog_title"),
        filetypes=filetypes,
    )
    if not src_path:
      return

    ext = os.path.splitext(src_path)[1].lower()
    if ext not in (".docx", ".odt", ".pdf"):
      messagebox.showwarning(self.t("fmt_type_title"), self.t("fmt_type_warn"))
      return

    dialog = ctk.CTkToplevel(self)
    dialog.title(self.t("add_dialog_title"))
    dialog.geometry("420x190")
    dialog.configure(fg_color=self.bg_card)
    dialog.grab_set()

    ctk.CTkLabel(
        dialog,
        text=self.t("add_dialog_q"),
        font=self.font_ui_bold,
        text_color=self.text_dark,
    ).pack(pady=(16, 6), padx=16, anchor="w")

    default_base = os.path.splitext(os.path.basename(src_path))[0]
    name_var = ctk.StringVar(value=default_base)
    name_entry = ctk.CTkEntry(
        dialog,
        textvariable=name_var,
        height=34,
        font=self.font_ui,
        fg_color=self.bg_soft,
        border_color=self.border_col,
        text_color=self.text_dark,
    )
    name_entry.pack(fill="x", padx=16, pady=4)
    self._bind_unicode_input(
        name_entry._entry,
        lambda txt: name_entry._entry.insert("insert", txt),
    )
    name_entry.focus_set()

    err_lbl = ctk.CTkLabel(
        dialog,
        text="",
        font=("Nirmala UI", 12),
        text_color=self.accent_red,
    )
    err_lbl.pack(padx=16, anchor="w")

    def on_save_click():
      raw_name = name_var.get().strip()
      if not raw_name:
        err_lbl.configure(text=self.t("enter_fmt_name"))
        return

      clean_base = self._sanitize_name(raw_name)
      if not clean_base:
        err_lbl.configure(text=self.t("valid_name_err"))
        return

      final_fname = f"{clean_base}{ext}"
      existing_lower = {f.lower() for f in os.listdir(self.templates_dir)}
      if final_fname.lower() in existing_lower:
        err_lbl.configure(text=self.t("exists_err", fname=final_fname))
        return

      dst_path = os.path.join(self.templates_dir, final_fname)
      shutil.copy2(src_path, dst_path)
      dialog.destroy()
      self.refresh_templates(select_name=final_fname)
      self.status_lbl.configure(
          text=self.t("fmt_saved_msg", fname=final_fname),
          text_color=self.status_ok,
      )

    ctk.CTkButton(
        dialog,
        text=self.t("save_fmt_btn"),
        command=on_save_click,
        font=self.font_ui_bold,
        fg_color=self.accent_blue,
        hover_color=self.accent_blue_hover,
        text_color="#ffffff",
        height=36,
        corner_radius=16,
    ).pack(pady=(8, 14), padx=16, fill="x")

  def open_remove_format_dialog(self):
    self._close_dropdown()
    if not self.all_templates:
      messagebox.showinfo(self.t("info_title"), self.t("no_fmt_to_del"))
      return

    dialog = ctk.CTkToplevel(self)
    dialog.title(self.t("del_dialog_title"))
    dialog.geometry("460x440")
    dialog.minsize(420, 420)
    dialog.configure(fg_color=self.bg_card)
    dialog.grab_set()

    ctk.CTkLabel(
        dialog,
        text=self.t("del_dialog_q"),
        font=self.font_ui_bold,
        text_color=self.text_dark,
    ).pack(pady=(14, 6), padx=16, anchor="w")

    del_search_var = ctk.StringVar()
    del_search_entry = ctk.CTkEntry(
        dialog,
        textvariable=del_search_var,
        height=34,
        font=self.font_ui,
        fg_color=self.bg_soft,
        border_color=self.border_col,
        text_color=self.text_dark,
        placeholder_text=self.t("del_search_ph"),
    )
    del_search_entry.pack(fill="x", padx=16, pady=(0, 8))
    self._bind_unicode_input(
        del_search_entry._entry,
        lambda txt: del_search_entry._entry.insert("insert", txt),
    )

    selected_var = ctk.StringVar(value="")
    selected_lbl = ctk.CTkLabel(
        dialog,
        text=self.t("sel_fmt_none"),
        font=("Nirmala UI", 12, "bold"),
        text_color=self.accent_red,
    )
    selected_lbl.pack(padx=16, anchor="w", pady=(0, 6))

    list_frame = ctk.CTkScrollableFrame(
        dialog,
        fg_color=self.bg_soft,
        border_width=1,
        border_color=self.border_col,
        corner_radius=10,
        height=165,
    )
    list_frame.pack(fill="x", expand=False, padx=16, pady=(0, 14))

    def choose_item(fname: str):
      selected_var.set(fname)
      selected_lbl.configure(text=self.t("sel_fmt_val", name=fname))

    def rebuild_list(*args):
      for w in list_frame.winfo_children():
        w.destroy()
      q = del_search_var.get().strip().lower()
      filtered = [f for f in self.all_templates if not q or q in f.lower()]
      if not filtered:
        ctk.CTkLabel(
            list_frame,
            text=self.t("no_match"),
            font=self.font_ui,
            text_color=self.text_muted,
        ).pack(pady=10, padx=10, anchor="w")
        return

      for fname in filtered:
        is_sel = selected_var.get() == fname
        btn = ctk.CTkButton(
            list_frame,
            text=fname,
            anchor="w",
            font=self.font_ui_bold if is_sel else self.font_ui,
            fg_color=self.hover_row if is_sel else "transparent",
            hover_color=self.hover_row,
            text_color=self.accent_red if is_sel else self.text_dark,
            height=30,
            corner_radius=6,
            command=lambda name=fname: (choose_item(name), rebuild_list()),
        )
        btn.pack(fill="x", padx=4, pady=2)

    del_search_var.trace_add("write", rebuild_list)
    rebuild_list()
    del_search_entry.focus_set()

    def confirm_delete():
      target = selected_var.get().strip()
      if not target:
        messagebox.showwarning(self.t("info_title"), self.t("pick_fmt_warn"))
        return
      if not messagebox.askyesno(
          self.t("confirm_title"), self.t("confirm_del_msg", target=target)
      ):
        return

      fpath = os.path.join(self.templates_dir, target)
      was_open = self.current_template == target
      if os.path.exists(fpath):
        os.remove(fpath)
      dialog.destroy()

      if was_open:
        self.refresh_templates(select_first=True)
      else:
        self.refresh_templates(select_name=self.current_template)

      self.status_lbl.configure(
          text=self.t("deleted_status", target=target),
          text_color=self.status_warn,
      )

    ctk.CTkButton(
        dialog,
        text=self.t("del_sel_btn"),
        command=confirm_delete,
        font=self.font_ui_bold,
        fg_color=self.accent_red,
        hover_color=self.accent_red_hover,
        text_color="#ffffff",
        height=38,
        corner_radius=18,
    ).pack(padx=16, pady=(0, 16), fill="x")

  def download_all_templates(self):
    self._close_dropdown()
    if not self.all_templates:
      messagebox.showinfo(self.t("info_title"), self.t("no_fmt_dl"))
      return

    target_dir = filedialog.askdirectory(title=self.t("pick_dl_dir"))
    if not target_dir:
      return

    export_folder = os.path.join(target_dir, "CourtWale_Formats")
    os.makedirs(export_folder, exist_ok=True)

    copied = 0
    for fname in self.all_templates:
      src = os.path.join(self.templates_dir, fname)
      dst = os.path.join(export_folder, fname)
      if os.path.exists(src):
        shutil.copy2(src, dst)
        copied += 1

    self.status_lbl.configure(
        text=self.t("dl_done_msg", n=copied, folder=export_folder),
        text_color=self.status_ok,
    )
    try:
      os.startfile(export_folder)
    except Exception:
      pass

  def _extract_blocks(self, src_path: str) -> list:
    ext = os.path.splitext(src_path)[1].lower()
    blocks = []

    if ext == ".docx":
      doc = Document(src_path)
      for p in doc.paragraphs:
        align_mode = self._map_docx_align(p.alignment)
        blocks.append((p.text, align_mode))
      for t in doc.tables:
        for row in t.rows:
          row_txt = " | ".join(c.text.strip() for c in row.cells)
          blocks.append((row_txt, "left"))
    elif ext == ".odt" and ODT_OK:
      odt = load_odt(src_path)
      for p in odt.getElementsByType(odf_text.P):
        txt = "".join(
            n.data
            for n in p.childNodes
            if getattr(n, "nodeType", None) == n.TEXT_NODE
        )
        blocks.append((txt, "left"))
    elif ext == ".pdf" and PDF_OK:
      doc = fitz.open(src_path)
      for page in doc:
        for line in page.get_text().splitlines():
          blocks.append((line, "left"))
      doc.close()

    return blocks

  def open_saved_template(self, filename):
    full_path = os.path.join(self.templates_dir, filename)
    if not os.path.exists(full_path):
      return

    self.current_template = filename
    blocks = self._extract_blocks(full_path)
    self._set_editor_blocks(blocks)

  def _build_file_prefix(self):
    case_val = self._sanitize_name(self.case_no_entry.get())
    ps_val = self._sanitize_name(self.ps_entry.get())
    fir_val = self._sanitize_name(self.fir_entry.get())

    if case_val:
      return case_val
    if ps_val and fir_val:
      return f"{ps_val}_{fir_val}"
    if ps_val:
      return ps_val
    if fir_val:
      return fir_val
    return "Blank_Case"

  def _replace_placeholders(self, text: str) -> str:
    case_txt = self.case_no_entry.get().strip()
    ps_txt = self.ps_entry.get().strip()
    fir_txt = self.fir_entry.get().strip()

    out = text
    if case_txt:
      out = re.sub(r"\{\{\s*case_no\s*\}\}", case_txt, out)
    if ps_txt:
      out = re.sub(r"\{\{\s*police_station\s*\}\}", ps_txt, out)
    if fir_txt:
      out = re.sub(r"\{\{\s*fir_no\s*\}\}", fir_txt, out)
    return out

  def _collect_editor_blocks_with_align(self) -> list:
    tb = self.inline_editor._textbox
    total_lines = int(tb.index("end-1c").split(".")[0])
    result = []

    for line_no in range(1, total_lines + 1):
      raw_line = tb.get(f"{line_no}.0", f"{line_no}.end")
      final_line = self._replace_placeholders(raw_line)
      tags = tb.tag_names(f"{line_no}.0")

      if "align_center" in tags:
        align_mode = "center"
      elif "align_right" in tags:
        align_mode = "right"
      else:
        align_mode = "left"

      result.append((final_line, align_mode))

    return result

  def _find_unicode_font_path(self):
    candidates = [
        r"C:\Windows\Fonts\Nirmala.ttf",
        r"C:\Windows\Fonts\NirmalaUI.ttf",
        r"C:\Windows\Fonts\NirmalaB.ttf",
        r"C:\Windows\Fonts\shruti.ttf",
        r"C:\Windows\Fonts\arialuni.ttf",
    ]
    for p in candidates:
      if os.path.exists(p):
        return p
    return None

  def _find_libreoffice(self):
    for cmd in [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]:
      if shutil.which(cmd) or os.path.exists(cmd):
        return cmd
    return None

  def _convert_docx_via_libreoffice(
      self, docx_path: str, out_dir: str, target_ext: str
  ):
    soffice = self._find_libreoffice()
    if not soffice:
      return None
    try:
      subprocess.run(
          [
              soffice,
              "--headless",
              "--convert-to",
              target_ext.lstrip("."),
              "--outdir",
              out_dir,
              docx_path,
          ],
          stdout=subprocess.DEVNULL,
          stderr=subprocess.DEVNULL,
          timeout=40,
      )
      gen_file = os.path.join(
          out_dir,
          os.path.splitext(os.path.basename(docx_path))[0] + target_ext,
      )
      if os.path.exists(gen_file):
        return gen_file
    except Exception:
      pass
    return None

  def _write_docx_from_blocks(self, line_blocks: list, target_docx: str):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Nirmala UI"
    style.font.size = Pt(self.editor_font_size)

    for line_txt, align_mode in line_blocks:
      p = doc.add_paragraph()
      if align_mode == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
      elif align_mode == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
      else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

      run = p.add_run(line_txt)
      run.font.name = "Nirmala UI"
      run.font.size = Pt(self.editor_font_size)

    doc.save(target_docx)

  def _render_image_from_blocks_pil(
      self, line_blocks: list, img_path: str
  ) -> bool:
    """Pillow + Windows Nirmala UI થી સીધી JPG/JPEG ઇમેજ બનાવે છે."""
    if not PIL_OK:
      return False
    try:
      page_w, page_h = 1654, 2339  # A4 @ 200 DPI
      margin_x = 120
      margin_y = 130
      px_size = max(24, int(self.editor_font_size * 2.2))
      line_h = int(px_size * 1.45)

      font_path = self._find_unicode_font_path()
      if font_path:
        font = ImageFont.truetype(font_path, px_size)
      else:
        font = ImageFont.load_default()

      pages = []
      img = Image.new("RGB", (page_w, page_h), "#ffffff")
      draw = ImageDraw.Draw(img)
      y = margin_y

      for line_txt, align_mode in line_blocks:
        sub_lines = line_txt.splitlines() if line_txt else [""]
        for sub_line in sub_lines:
          if y + line_h > page_h - margin_y:
            pages.append(img)
            img = Image.new("RGB", (page_w, page_h), "#ffffff")
            draw = ImageDraw.Draw(img)
            y = margin_y

          if sub_line.strip():
            bbox = draw.textbbox((0, 0), sub_line, font=font)
            text_w = bbox[2] - bbox[0]
            if align_mode == "center":
              x = max(margin_x, (page_w - text_w) // 2)
            elif align_mode == "right":
              x = max(margin_x, page_w - margin_x - text_w)
            else:
              x = margin_x

            draw.text((x, y), sub_line, font=font, fill="#111827")

          y += line_h

      pages.append(img)
      pages[0].save(img_path, format="JPEG", quality=95)
      return os.path.exists(img_path)
    except Exception:
      return False

  def _create_pdf_direct(
      self, line_blocks: list, docx_path: str, pdf_path: str
  ) -> bool:
    """LibreOffice હોય તો તેનાથી, ન હોય તો PyMuPDF થી સીધી PDF બનાવે છે."""
    gen_pdf = self._convert_docx_via_libreoffice(
        docx_path, os.path.dirname(pdf_path), ".pdf"
    )
    if gen_pdf and os.path.exists(gen_pdf):
      if os.path.abspath(gen_pdf) != os.path.abspath(pdf_path):
        if os.path.exists(pdf_path):
          os.remove(pdf_path)
        os.rename(gen_pdf, pdf_path)
      return True

    if not PDF_OK:
      return False

    try:
      doc = fitz.open()
      page_w, page_h = 595.0, 842.0
      margin_x = 45.0
      margin_y = 50.0
      fsize = float(self.editor_font_size)
      line_h = fsize + 9.0

      font_path = self._find_unicode_font_path()
      custom_font = fitz.Font(fontfile=font_path) if font_path else None

      page = doc.new_page(width=page_w, height=page_h)
      y = margin_y

      for line_txt, align_mode in line_blocks:
        wrapped = line_txt.splitlines() if line_txt else [""]
        for sub_line in wrapped:
          if y + line_h > page_h - margin_y:
            page = doc.new_page(width=page_w, height=page_h)
            y = margin_y

          if sub_line.strip():
            if custom_font is not None:
              text_w = custom_font.text_length(sub_line, fontsize=fsize)
              if align_mode == "center":
                x = max(margin_x, (page_w - text_w) / 2.0)
              elif align_mode == "right":
                x = max(margin_x, page_w - margin_x - text_w)
              else:
                x = margin_x
              tw = fitz.TextWriter(page.rect)
              tw.append(
                  fitz.Point(x, y + fsize),
                  sub_line,
                  font=custom_font,
                  fontsize=fsize,
              )
              tw.write_text(page, color=(0.07, 0.09, 0.12))
            else:
              page.insert_text(
                  fitz.Point(margin_x, y + fsize),
                  sub_line,
                  fontsize=fsize,
                  fontname="helv",
                  color=(0.07, 0.09, 0.12),
              )
          y += line_h

      doc.save(pdf_path)
      doc.close()
      return os.path.exists(pdf_path)
    except Exception:
      return False

  def _create_jpeg_direct(
      self, line_blocks: list, docx_path: str, img_path: str
  ) -> bool:
    """સીધી JPG/JPEG ફાઇલ બનાવે છે (LibreOffice -> PyMuPDF -> Pillow fallback)."""
    tmp_pdf = os.path.splitext(img_path)[0] + "_render_tmp.pdf"
    gen_pdf = self._convert_docx_via_libreoffice(
        docx_path, os.path.dirname(tmp_pdf), ".pdf"
    )
    if gen_pdf and os.path.exists(gen_pdf) and PDF_OK:
      try:
        doc = fitz.open(gen_pdf)
        if len(doc) > 0:
          pix = doc[0].get_pixmap(dpi=200)
          pix.save(img_path, jpg_quality=95)
          doc.close()
          os.remove(gen_pdf)
          if os.path.exists(img_path):
            return True
      except Exception:
        pass
    if os.path.exists(gen_pdf or ""):
      try:
        os.remove(gen_pdf)
      except Exception:
        pass

    return self._render_image_from_blocks_pil(line_blocks, img_path)

  def _export_to_target_path(
      self, line_blocks: list, ext_tpl: str, selected_fmt: str, target_path: str
  ):
    out_dir = os.path.dirname(target_path)
    os.makedirs(out_dir, exist_ok=True)
    tmp_docx = os.path.join(
        out_dir, f"_tmp_export_{os.getpid()}_{os.path.basename(target_path)}.docx"
    )
    self._write_docx_from_blocks(line_blocks, tmp_docx)

    final_path = target_path
    extra_note = ""

    try:
      if selected_fmt == "DOCX / ODT":
        if ext_tpl == ".odt" and ODT_OK:
          final_path = os.path.splitext(target_path)[0] + ".odt"
          odt = OpenDocumentText()
          for line_txt, _ in line_blocks:
            odt.text.addElement(odf_text.P(text=line_txt))
          odt.save(final_path)
        else:
          final_path = os.path.splitext(target_path)[0] + ".docx"
          shutil.copy2(tmp_docx, final_path)

      elif selected_fmt == "DOC":
        final_path = os.path.splitext(target_path)[0] + ".doc"
        gen_doc = self._convert_docx_via_libreoffice(tmp_docx, out_dir, ".doc")
        if gen_doc and os.path.exists(gen_doc):
          if os.path.abspath(gen_doc) != os.path.abspath(final_path):
            if os.path.exists(final_path):
              os.remove(final_path)
            os.rename(gen_doc, final_path)
        else:
          shutil.copy2(tmp_docx, final_path)

      elif selected_fmt == "PDF":
        final_path = os.path.splitext(target_path)[0] + ".pdf"
        if not self._create_pdf_direct(line_blocks, tmp_docx, final_path):
          final_path = os.path.splitext(target_path)[0] + ".docx"
          shutil.copy2(tmp_docx, final_path)
          extra_note = self.t("missing_deps_note")

      elif selected_fmt in ("JPEG", "JPG"):
        img_ext = ".jpeg" if selected_fmt == "JPEG" else ".jpg"
        final_path = os.path.splitext(target_path)[0] + img_ext
        if not self._create_jpeg_direct(line_blocks, tmp_docx, final_path):
          final_path = os.path.splitext(target_path)[0] + ".docx"
          shutil.copy2(tmp_docx, final_path)
          extra_note = self.t("missing_deps_note")
    finally:
      if os.path.exists(tmp_docx):
        try:
          os.remove(tmp_docx)
        except Exception:
          pass

    return final_path, extra_note

  def save_file_to_chosen_path(self):
    self._close_dropdown()
    if not self.current_template:
      self.status_lbl.configure(
          text=self.t("select_fmt_first"), text_color=self.accent_red
      )
      return

    ext_tpl = os.path.splitext(self.current_template)[1].lower()
    base_name = os.path.splitext(self.current_template)[0]
    prefix = self._build_file_prefix()
    selected_fmt = self.out_fmt_menu.get()

    fmt_ext_map = {
        "DOCX / ODT": (".odt" if ext_tpl == ".odt" else ".docx"),
        "DOC": ".doc",
        "PDF": ".pdf",
        "JPEG": ".jpeg",
        "JPG": ".jpg",
    }
    default_ext = fmt_ext_map.get(selected_fmt, ".docx")
    default_fname = f"{prefix}_{base_name}{default_ext}"

    chosen_path = filedialog.asksaveasfilename(
        title=self.t("save_dialog_title"),
        initialdir=self.output_dir,
        initialfile=default_fname,
        defaultextension=default_ext,
        filetypes=[
            ("Selected Format", f"*{default_ext}"),
            ("All Files", "*.*"),
        ],
    )
    if not chosen_path:
      return

    line_blocks = self._collect_editor_blocks_with_align()
    final_path, extra_note = self._export_to_target_path(
        line_blocks, ext_tpl, selected_fmt, chosen_path
    )

    self.status_lbl.configure(
        text=self.t("saved_status", path=final_path) + extra_note,
        text_color=self.status_ok,
    )

  def generate_file(self):
    self._close_dropdown()
    if not self.current_template:
      self.status_lbl.configure(
          text=self.t("select_fmt_first"), text_color=self.accent_red
      )
      return

    ext_tpl = os.path.splitext(self.current_template)[1].lower()
    base_name = os.path.splitext(self.current_template)[0]
    prefix = self._build_file_prefix()

    out_folder = os.path.join(self.output_dir, prefix)
    os.makedirs(out_folder, exist_ok=True)

    line_blocks = self._collect_editor_blocks_with_align()
    selected_fmt = self.out_fmt_menu.get()
    base_out_name = f"{prefix}_{base_name}"
    initial_target = os.path.join(out_folder, f"{base_out_name}.docx")

    final_saved_path, extra_note = self._export_to_target_path(
        line_blocks, ext_tpl, selected_fmt, initial_target
    )

    self.status_lbl.configure(
        text=self.t("saved_status", path=final_saved_path) + extra_note,
        text_color=self.status_ok,
    )
    try:
      os.startfile(out_folder)
    except Exception:
      pass


if __name__ == "__main__":
  app = CourtWaleCleanApp()
  app.mainloop()