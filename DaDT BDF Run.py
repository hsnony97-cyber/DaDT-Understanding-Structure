#!/usr/bin/env python3
"""
DaDT Structure Definition Tool v16.0
=====================================
Tab 1: BDF Merge Preparation
Tab 2: Understanding Structure Type (maneuver->thermal offset, 'Bar Property Structure Type' sheet)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import threading
import csv
import shutil
import tempfile
import subprocess
import time
import concurrent.futures
from datetime import datetime
import pandas as pd
from pyNastran.bdf.bdf import BDF
from pyNastran.op2.op2 import OP2
import numpy as np

_APP_COLORS = {}  # populated by _configure_app_style(); shared palette for both tabs


def _configure_app_style(root):
    """One-time visual theme for the whole app. Both tabs share this root's
    ttk.Style singleton, so configuring it once here restyles every ttk
    widget in both tabs without touching how any of them behave. Pure
    look-and-feel - no widget's command/callback is affected."""
    colors = {
        'bg': '#f1f5f9',
        'surface': '#ffffff',
        'text': '#0f172a',
        'muted': '#64748b',
        'border': '#cbd5e1',
        'primary': '#2563eb',
        'primary_dark': '#1d4ed8',
        'success': '#16a34a',
        'success_dark': '#15803d',
        'danger': '#dc2626',
        'danger_dark': '#b91c1c',
        'log_bg': '#0f172a',
        'log_fg': '#e2e8f0',
    }

    root.configure(bg=colors['bg'])

    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except tk.TclError:
        pass

    base_font = ('Segoe UI', 10)
    section_font = ('Segoe UI', 10, 'bold')

    style.configure('.', background=colors['bg'], foreground=colors['text'], font=base_font)
    style.configure('TFrame', background=colors['bg'])
    style.configure('TLabel', background=colors['bg'], foreground=colors['text'])
    style.configure('TCheckbutton', background=colors['bg'])

    style.configure('Header.TLabel', font=('Segoe UI', 15, 'bold'),
                    foreground=colors['primary_dark'], background=colors['bg'])
    style.configure('Subtitle.TLabel', font=('Segoe UI', 9),
                    foreground=colors['muted'], background=colors['bg'])

    style.configure('TLabelframe', background=colors['bg'], bordercolor=colors['border'])
    style.configure('TLabelframe.Label', background=colors['bg'],
                    foreground=colors['primary_dark'], font=section_font)

    style.configure('TNotebook', background=colors['bg'], borderwidth=0)
    style.configure('TNotebook.Tab', font=('Segoe UI', 10, 'bold'), padding=(16, 8),
                    background=colors['border'], foreground=colors['text'])
    style.map('TNotebook.Tab',
              background=[('selected', colors['surface'])],
              foreground=[('selected', colors['primary_dark'])])

    style.configure('TButton', font=base_font, padding=(10, 5),
                    background=colors['surface'], foreground=colors['text'])
    style.map('TButton', background=[('active', colors['border'])])

    style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'), padding=(14, 7),
                    background=colors['primary'], foreground='white')
    style.map('Accent.TButton',
              background=[('active', colors['primary_dark']), ('disabled', colors['border'])],
              foreground=[('disabled', colors['muted'])])

    style.configure('Danger.TButton', font=('Segoe UI', 10, 'bold'), padding=(14, 7),
                    background=colors['danger'], foreground='white')
    style.map('Danger.TButton',
              background=[('active', colors['danger_dark']), ('disabled', colors['border'])],
              foreground=[('disabled', colors['muted'])])

    style.configure('TEntry', padding=4, fieldbackground=colors['surface'], bordercolor=colors['border'])
    style.configure('TCombobox', padding=4)

    style.configure('Horizontal.TProgressbar', troughcolor=colors['border'],
                    background=colors['primary'], bordercolor=colors['border'], thickness=14)

    _APP_COLORS.update(colors)
    return colors


class IntegratedBDFRFTool:
    def __init__(self, root):
        self.root = root
        self.root.title("DaDT Structure Definition Tool v16.0")
        self.root.geometry("1150x950")
        self.root.minsize(1000, 700)
        self.colors = _configure_app_style(self.root)

        # Tab 1 variables
        self.thermal_bdfs = []
        self.maneuver_bdfs = []
        self.excel_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_thermal_name = tk.StringVar(value="merged_thermal.bdf")
        self.output_maneuver_name = tk.StringVar(value="merged_maneuver.bdf")
        self.set_id = tk.StringVar(value="99")
        self.node_set_id = tk.StringVar(value="199")
        self.temp_initial = tk.StringVar(value="10")
        

        self.setup_ui()
    
    def setup_ui(self):
        banner = ttk.Frame(self.root, padding=(16, 12))
        banner.pack(fill=tk.X)
        ttk.Label(banner, text="\U0001F527 DaDT Structure Definition Tool", style='Header.TLabel').pack(side=tk.LEFT)
        ttk.Label(banner, text="  v16.0", style='Subtitle.TLabel').pack(side=tk.LEFT, pady=(6, 0))
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="\U0001F4C4  BDF Merge Preparation")
        self.setup_tab1()

        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="\U0001F9EE  Understanding Structure Type")
        self.tab3_tool = BarPropertySolverTab(self.tab3, self.root)

    
    def setup_tab1(self):
        main = ttk.Frame(self.tab1, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main, text="BDF Merge Preparation v8", style='Header.TLabel').pack(pady=(0,10))

        # === THERMAL SECTION ===
        thermal_main = ttk.LabelFrame(main, text="\U0001F321  THERMAL", padding="8")
        thermal_main.pack(fill=tk.X, pady=5)
        
        thm_master_f = ttk.Frame(thermal_main)
        thm_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(thm_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(thm_master_f, text="Add...", command=self.add_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(thm_master_f, text="Clear", command=self.clear_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        self.thermal_count = tk.StringVar(value="0 files")
        ttk.Label(thm_master_f, textvariable=self.thermal_count).pack(side=tk.LEFT, padx=5)
        
        self.thermal_listbox = tk.Listbox(thermal_main, height=3, width=100, font=('Segoe UI', 9),
                                           bg=self.colors['surface'], fg=self.colors['text'],
                                           selectbackground=self.colors['primary'], selectforeground='white',
                                           relief='flat', highlightthickness=1,
                                           highlightbackground=self.colors['border'],
                                           highlightcolor=self.colors['primary'])
        self.thermal_listbox.pack(fill=tk.X, pady=2)
        
        # === MANEUVER SECTION ===
        maneuver_main = ttk.LabelFrame(main, text="\U0001F6EB  MANEUVER", padding="8")
        maneuver_main.pack(fill=tk.X, pady=5)
        
        man_master_f = ttk.Frame(maneuver_main)
        man_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(man_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(man_master_f, text="Add...", command=self.add_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(man_master_f, text="Clear", command=self.clear_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        self.maneuver_count = tk.StringVar(value="0 files")
        ttk.Label(man_master_f, textvariable=self.maneuver_count).pack(side=tk.LEFT, padx=5)
        
        self.maneuver_listbox = tk.Listbox(maneuver_main, height=3, width=100, font=('Segoe UI', 9),
                                            bg=self.colors['surface'], fg=self.colors['text'],
                                            selectbackground=self.colors['primary'], selectforeground='white',
                                            relief='flat', highlightthickness=1,
                                            highlightbackground=self.colors['border'],
                                            highlightcolor=self.colors['primary'])
        self.maneuver_listbox.pack(fill=tk.X, pady=2)
        
        # === SETTINGS ===
        sf = ttk.LabelFrame(main, text="⚙  Settings", padding="10")
        sf.pack(fill=tk.X, pady=5)
        ttk.Label(sf, text="Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.excel_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        ttk.Label(sf, text="Output:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.output_folder, width=70).grid(row=1, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_output).grid(row=1, column=2)
        ttk.Label(sf, text="SET ID:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.set_id, width=10).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Label(sf, text="NODE SET ID:").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.node_set_id, width=10).grid(row=3, column=1, sticky=tk.W, padx=5)
        ttk.Label(sf, text="TEMP(INIT):").grid(row=4, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.temp_initial, width=10).grid(row=4, column=1, sticky=tk.W, padx=5)
        
        bf = ttk.Frame(main)
        bf.pack(fill=tk.X, pady=10)
        self.process_btn = ttk.Button(bf, text="▶  PROCESS & MERGE", command=self.start_processing,
                                       style='Accent.TButton')
        self.process_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Clear Log", command=self.clear_log1).pack(side=tk.LEFT)

        self.progress1 = ttk.Progressbar(main, mode='indeterminate')
        self.progress1.pack(fill=tk.X, pady=5)

        lf = ttk.LabelFrame(main, text="\U0001F4DC  Log", padding="10")
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text1 = scrolledtext.ScrolledText(lf, height=15, font=('Consolas', 10),
                                                    bg=self.colors['log_bg'], fg=self.colors['log_fg'],
                                                    insertbackground=self.colors['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text1.pack(fill=tk.BOTH, expand=True)
    
    # ============= TAB 1 HELPERS =============
    def add_thermal_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.thermal_bdfs:
                self.thermal_bdfs.append(f)
                self.thermal_listbox.insert(tk.END, f)
        self.thermal_count.set(f"{len(self.thermal_bdfs)} files")
    
    def clear_thermal_bdfs(self):
        self.thermal_bdfs.clear()
        self.thermal_listbox.delete(0, tk.END)
        self.thermal_count.set("0 files")
    
    def add_maneuver_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, f)
        self.maneuver_count.set(f"{len(self.maneuver_bdfs)} files")
    
    def clear_maneuver_bdfs(self):
        self.maneuver_bdfs.clear()
        self.maneuver_listbox.delete(0, tk.END)
        self.maneuver_count.set("0 files")
    
    def browse_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f: self.excel_path.set(f)
    
    def browse_output(self):
        f = filedialog.askdirectory()
        if f: self.output_folder.set(f)
    
    def log1(self, msg):
        self.log_text1.insert(tk.END, msg + "\n")
        self.log_text1.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log1(self):
        self.log_text1.delete(1.0, tk.END)
    
    def format_include_nastran(self, abs_path):
        """Uzun INCLUDE path'lerini Nastran formatına uygun böler."""
        include_line = f"INCLUDE '{abs_path}'"
        if len(include_line) <= 72:
            return [include_line]
        parts = abs_path.split('/')
        lines = []
        current_line = "INCLUDE '"
        for i, part in enumerate(parts):
            is_last = (i == len(parts) - 1)
            segment = part if is_last else part + '/'
            if len(current_line + segment) <= 72:
                current_line += segment
            else:
                if current_line != "INCLUDE '":
                    lines.append(current_line)
                current_line = segment
        if current_line:
            current_line += "'"
            lines.append(current_line)
        return lines
    
    def read_file_safe(self, fpath):
        """Dosyayı güvenli şekilde oku."""
        for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
            try:
                with open(fpath, 'r', encoding=enc, errors='replace') as f:
                    return f.read()
            except:
                continue
        return ""
    
    def extract_subcase_load_info(self, bdf_content):
        """
        BDF içeriğinden TÜM SUBCASE bilgilerini çıkarır.
        Bir dosyada birden fazla SUBCASE olabilir (MASTER_GUST.BDF gibi).
        
        Returns: list of dicts, her biri:
            - subcase_id
            - load_id
            - temp_load_id
            - subtitle
        """
        results = []
        lines = bdf_content.split('\n')
        
        current_subcase = None
        current_load = None
        current_temp_load = None
        current_subtitle = None
        
        for line in lines:
            upper = line.upper().strip()
            original = line.strip()
            
            # SUBCASE satırı
            if upper.startswith('SUBCASE'):
                # Önceki subcase'i kaydet
                if current_subcase is not None:
                    results.append({
                        'subcase_id': current_subcase,
                        'load_id': current_load,
                        'temp_load_id': current_temp_load,
                        'subtitle': current_subtitle
                    })
                
                # Yeni subcase başlat
                parts = upper.split()
                if len(parts) >= 2:
                    try:
                        current_subcase = int(parts[1])
                        current_load = None
                        current_temp_load = None
                        current_subtitle = None
                    except:
                        pass
            
            # LOAD = satırı
            elif current_subcase and upper.startswith('LOAD') and '=' in upper:
                match = re.search(r'LOAD\s*=\s*(\d+)', upper)
                if match:
                    current_load = int(match.group(1))
            
            # TEMPERATURE(LOAD) = satırı
            elif current_subcase and 'TEMPERATURE' in upper and 'LOAD' in upper and '=' in upper:
                match = re.search(r'TEMPERATURE\s*\(\s*LOAD\s*\)\s*=\s*(\d+)', upper)
                if match:
                    current_temp_load = int(match.group(1))
            
            # SUBTITLE satırı
            elif current_subcase and upper.startswith('SUBTITLE'):
                m = re.search(r'SUBTITLE\s+(.+)', original, re.IGNORECASE)
                if m:
                    current_subtitle = m.group(1).strip()
            
            # BEGIN BULK'a ulaştıysak case control bitti
            elif upper.startswith('BEGIN') and 'BULK' in upper:
                break
        
        # Son subcase'i kaydet
        if current_subcase is not None:
            results.append({
                'subcase_id': current_subcase,
                'load_id': current_load,
                'temp_load_id': current_temp_load,
                'subtitle': current_subtitle
            })
        
        return results
    
    def parse_multiline_includes(self, content, bdf_dir):
        """
        Çok satırlı INCLUDE'ları parse eder.
        Nastran formatında INCLUDE şöyle olabilir:
        
        INCLUDE '../../../_COMMON_/INTERFACE/STRUCTURAL/INTER_AF_HT/
        INTER_AF_HT_STRU.BDF'
        
        Returns: list of dict with keys: lines, full_text, abs_path, start_idx, end_idx
        """
        lines = content.split('\n')
        includes = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            upper = line.upper().strip()
            
            if upper.startswith('INCLUDE'):
                # INCLUDE başladı - tırnak içindeki path'i bul
                include_lines = [line]
                full_text = line
                
                # Tırnak sayısını kontrol et - tek tırnak veya çift tırnak
                quote_char = None
                if "'" in line:
                    quote_char = "'"
                elif '"' in line:
                    quote_char = '"'
                
                if quote_char:
                    # Tırnak sayısını say
                    quote_count = full_text.count(quote_char)
                    
                    # Eğer tek tırnak varsa (açılmış ama kapanmamış), devam satırlarını oku
                    j = i + 1
                    while quote_count % 2 != 0 and j < len(lines):
                        next_line = lines[j]
                        include_lines.append(next_line)
                        full_text += '\n' + next_line
                        quote_count = full_text.count(quote_char)
                        j += 1
                    
                    # Path'i çıkar - newline'ları temizle
                    clean_text = full_text.replace('\n', '')
                    match = re.search(rf"INCLUDE\s*{quote_char}([^{quote_char}]*){quote_char}", 
                                     clean_text, re.IGNORECASE)
                    if match:
                        inc_path = match.group(1).strip()
                        # Absolute path'e çevir
                        if not os.path.isabs(inc_path):
                            abs_path = os.path.normpath(os.path.join(bdf_dir, inc_path))
                        else:
                            abs_path = os.path.normpath(inc_path)
                        abs_path = abs_path.replace('\\', '/')
                        
                        includes.append({
                            'lines': include_lines,
                            'full_text': full_text,
                            'abs_path': abs_path,
                            'start_idx': i,
                            'end_idx': j - 1 if j > i + 1 else i
                        })
                    
                    i = j
                else:
                    i += 1
            else:
                i += 1
        
        return includes
    
    def collect_all_lines_from_masters(self, bdf_files, load_case_set, common_type):
        """
        Tüm MASTER BDF'lerden TÜM SATIRLARI toplar.
        
        1. Excel'deki SUBCASE ID'lere uyan MASTER BDF'leri bulur
        2. Her birinden TÜM satırları alır ve alt alta yapıştırır
        3. INCLUDE'ları kategorize et:
           - COMMON LOAD/THERMAL → Ayrı tut (sonra INCLUDE olarak eklenecek)
           - Diğerleri (STRUCTURE, INTERFACE, vs.) → Satır olarak ekle (pyNastran açacak)
        4. Duplicate satırları çıkar
        
        NOT: Bir BDF birden fazla SUBCASE içerebilir (MASTER_GUST.BDF gibi)
        
        common_type: 'LOAD' veya 'THERMAL'
        
        Returns: (all_lines, common_includes, subcase_info_map)
        """
        all_lines_raw = []  # Tüm satırlar (INCLUDE satırları dahil - COMMON hariç)
        common_includes_raw = []  # COMMON INCLUDE path'leri
        subcase_info_map = {}
        processed_files = set()  # Aynı dosyayı birden fazla kez işlememek için
        
        self.log1(f"    Reading ALL lines from {len(bdf_files)} MASTER BDFs...")
        matched_count = 0
        matched_subcases = 0
        
        for bdf_path in bdf_files:
            bdf_dir = os.path.dirname(os.path.abspath(bdf_path))
            content = self.read_file_safe(bdf_path)
            
            # INREL dosyalarını atla
            bdf_basename = os.path.basename(bdf_path).upper()
            if 'INREL' in bdf_basename:
                self.log1(f"      SKIP (INREL): {os.path.basename(bdf_path)}")
                continue
            
            # TÜM SUBCASE bilgilerini al (birden fazla olabilir)
            all_subcases = self.extract_subcase_load_info(content)
            
            # Bu dosyadaki hangi subcase'ler Excel listesinde?
            matching_subcases = []
            for sc_info in all_subcases:
                sc_id = sc_info['subcase_id']
                if sc_id and sc_id in load_case_set:
                    matching_subcases.append(sc_info)
            
            # Eşleşen subcase varsa bu dosyayı işle
            if matching_subcases:
                # Dosya daha önce işlendiyse sadece subcase info'ları ekle
                if bdf_path in processed_files:
                    for sc_info in matching_subcases:
                        sc_id = sc_info['subcase_id']
                        if sc_id not in subcase_info_map:
                            if common_type == 'THERMAL':
                                subcase_info_map[sc_id] = {
                                    'temp_load_id': sc_info['temp_load_id'] if sc_info['temp_load_id'] else sc_id,
                                    'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Thermal Case {sc_id}"
                                }
                            else:
                                subcase_info_map[sc_id] = {
                                    'load_id': sc_info['load_id'] if sc_info['load_id'] else sc_id,
                                    'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Manoeuvre {sc_id}"
                                }
                            matched_subcases += 1
                    continue
                
                processed_files.add(bdf_path)
                matched_count += 1
                
                # Log - kaç subcase eşleşti
                sc_ids = [str(sc['subcase_id']) for sc in matching_subcases]
                self.log1(f"      MATCH: {os.path.basename(bdf_path)} ({len(matching_subcases)} subcases: {', '.join(sc_ids[:5])}{'...' if len(sc_ids) > 5 else ''})")
                
                # Subcase info'ları kaydet
                for sc_info in matching_subcases:
                    sc_id = sc_info['subcase_id']
                    matched_subcases += 1
                    if common_type == 'THERMAL':
                        subcase_info_map[sc_id] = {
                            'temp_load_id': sc_info['temp_load_id'] if sc_info['temp_load_id'] else sc_id,
                            'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Thermal Case {sc_id}"
                        }
                    else:
                        subcase_info_map[sc_id] = {
                            'load_id': sc_info['load_id'] if sc_info['load_id'] else sc_id,
                            'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Manoeuvre {sc_id}"
                        }
                
                # Önce tüm INCLUDE'ları parse et (çok satırlı dahil)
                all_includes = self.parse_multiline_includes(content, bdf_dir)
                
                # INCLUDE'ları kategorize et
                common_include_indices = set()  # COMMON INCLUDE satır indeksleri
                structure_include_count = 0
                common_include_count = 0
                
                for inc in all_includes:
                    abs_path_upper = inc['abs_path'].upper()
                    
                    # Bu INCLUDE COMMON LOAD/THERMAL mı?
                    is_common = False
                    if common_type == 'LOAD':
                        if '_COMMON_/LOAD' in abs_path_upper or '/COMMON/LOAD' in abs_path_upper:
                            is_common = True
                    elif common_type == 'THERMAL':
                        if '_COMMON_/THERMAL' in abs_path_upper or '/COMMON/THERMAL' in abs_path_upper:
                            is_common = True
                    
                    if is_common:
                        # COMMON INCLUDE - path'i kaydet, satırları atla
                        common_includes_raw.append(inc['abs_path'])
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        common_include_count += 1
                    else:
                        # Structure/Interface/vs INCLUDE - tek satır INCLUDE olarak ekle (pyNastran açacak)
                        # Absolute path ile yeni INCLUDE satırı oluştur
                        include_line = f"INCLUDE '{inc['abs_path']}'"
                        all_lines_raw.append(include_line)
                        # Orijinal satırları atla (common_include_indices'e ekle)
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        structure_include_count += 1
                
                # TÜM SATIRLARI oku (INCLUDE satırları hariç - hem COMMON hem diğerleri)
                lines = content.split('\n')
                line_count = 0
                
                for idx, line in enumerate(lines):
                    # INCLUDE satırı mı? Atla (zaten yukarıda işledik)
                    if idx in common_include_indices:
                        continue
                    
                    stripped = line.strip()
                    if not stripped:
                        continue
                    
                    # Case control satırlarını atla
                    upper = stripped.upper()
                    skip_keywords = ['SOL ', 'SOL\t', 'CEND', 'TITLE', 'SUBTITLE', 'ECHO', 
                                    'SUBCASE', 'LOAD =', 'LOAD=', 'SPC =', 'SPC=', 
                                    'TEMPERATURE', 'DISPLACEMENT', 'FORCE', 'GPFORCE', 
                                    'OLOAD', 'SPCFORCE', 'SET ', 'BEGIN BULK', 
                                    'BEGIN,BULK', 'ENDDATA']
                    is_skip = any(upper.startswith(kw) for kw in skip_keywords)
                    
                    if not is_skip:
                        all_lines_raw.append(line)  # Orijinal satırı koru
                        line_count += 1
                
                self.log1(f"        -> {line_count} data lines, {structure_include_count} structure includes, {common_include_count} common includes")
        
        self.log1(f"    Matched {matched_count} MASTER BDFs with {matched_subcases} total subcases")
        self.log1(f"    Total raw lines (including structure includes): {len(all_lines_raw)}")
        self.log1(f"    Total raw COMMON includes: {len(common_includes_raw)}")
        
        # Duplicate satırları çıkar
        self.log1("    Removing duplicate lines...")
        seen_lines = set()
        unique_lines = []
        for line in all_lines_raw:
            # Normalize et (boşlukları düzenle)
            normalized = ' '.join(line.split())
            if normalized not in seen_lines:
                seen_lines.add(normalized)
                unique_lines.append(line)
        
        # Duplicate COMMON INCLUDE'ları çıkar
        seen_includes = set()
        unique_common_includes = []
        for inc_path in common_includes_raw:
            normalized = inc_path.lower().replace('\\', '/')
            if normalized not in seen_includes:
                seen_includes.add(normalized)
                unique_common_includes.append(inc_path)
        
        self.log1(f"    After removing duplicates:")
        self.log1(f"      Unique lines: {len(unique_lines)} (removed {len(all_lines_raw) - len(unique_lines)} duplicates)")
        self.log1(f"      Unique COMMON includes: {len(unique_common_includes)}")
        
        return unique_lines, unique_common_includes, subcase_info_map
    
    def extract_param_cards(self, bulk_data):
        """
        Bulk data'dan PARAM kartlarını ayırır.
        Returns: (param_lines, remaining_bulk_data)
        """
        lines = bulk_data.split('\n')
        param_lines = []
        other_lines = []
        
        seen_params = set()  # Duplicate PARAM kontrolü
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            upper = stripped.upper()
            
            if upper.startswith('PARAM'):
                # PARAM kartı - continuation satırlarını da al
                param_block = [line]
                
                # PARAM name'ini çıkar (duplicate kontrolü için)
                param_name = None
                try:
                    if ',' in line:
                        parts = line.split(',')
                        if len(parts) > 1:
                            param_name = parts[1].strip().upper()
                    else:
                        if len(line) >= 16:
                            param_name = line[8:16].strip().upper()
                except:
                    pass
                
                i += 1
                # Continuation satırlarını kontrol et
                while i < len(lines):
                    next_line = lines[i]
                    next_stripped = next_line.strip()
                    is_cont = (next_line.startswith('+') or 
                              next_line.startswith('*') or
                              (next_line.startswith('        ') and next_stripped and 
                               not next_stripped.startswith('$') and
                               not any(next_stripped.upper().startswith(card) for card in 
                                      ['PARAM', 'GRID', 'CBAR', 'CBEAM', 'CQUAD', 'CTRIA', 
                                       'MAT', 'PBAR', 'PSHELL', 'FORCE', 'MOMENT', 'RBE',
                                       'CORD', 'SPC', 'MPC', 'INCLUDE', 'ENDDATA'])))
                    if is_cont and next_stripped:
                        param_block.append(next_line)
                        i += 1
                    else:
                        break
                
                # Duplicate kontrolü
                if param_name and param_name not in seen_params:
                    seen_params.add(param_name)
                    param_lines.extend(param_block)
            else:
                other_lines.append(line)
                i += 1
        
        return param_lines, '\n'.join(other_lines)
    
    def check_and_remove_duplicates(self, bulk_data):
        """
        Bulk data içindeki duplicate kartları tespit edip kaldırır.
        
        - Element/Property/Material kartları: ID bazlı kontrol (aynı ID → duplicate)
        - SPC/FORCE/MOMENT kartları: Tüm satır bazlı kontrol (birebir aynıysa → duplicate)
        """
        self.log1("    Checking for duplicate entries...")
        
        lines = bulk_data.split('\n')
        
        # ID bazlı kontrol yapılacak kartlar
        id_based_cards = {
            'GRID': set(),
            'CBAR': set(),
            'CBEAM': set(),
            'CROD': set(),
            'CONROD': set(),
            'CQUAD4': set(),
            'CQUAD8': set(),
            'CTRIA3': set(),
            'CTRIA6': set(),
            'CHEXA': set(),
            'CPENTA': set(),
            'CTETRA': set(),
            'CBUSH': set(),
            'CELAS1': set(),
            'CELAS2': set(),
            'CDAMP1': set(),
            'CDAMP2': set(),
            'CMASS1': set(),
            'CMASS2': set(),
            'RBE2': set(),
            'RBE3': set(),
            'PBAR': set(),
            'PBARL': set(),
            'PBEAM': set(),
            'PBEAML': set(),
            'PROD': set(),
            'PSHELL': set(),
            'PCOMP': set(),
            'PCOMPG': set(),
            'PSOLID': set(),
            'PBUSH': set(),
            'PELAS': set(),
            'PDAMP': set(),
            'PMASS': set(),
            'PTUBE': set(),
            'PVISC': set(),
            'PGAP': set(),
            'PWELD': set(),
            'MAT1': set(),
            'MAT2': set(),
            'MAT8': set(),
            'MAT9': set(),
            'MATS1': set(),
            'CORD1R': set(),
            'CORD2R': set(),
            'CORD1C': set(),
            'CORD2C': set(),
            'CORD1S': set(),
            'CORD2S': set(),
        }
        
        # Tüm satır bazlı kontrol yapılacak kartlar (SPC, FORCE, MOMENT, MPC)
        line_based_cards = ['SPC', 'SPC1', 'FORCE', 'MOMENT', 'MPC', 'LOAD', 'TEMP', 'TEMPD']
        seen_full_lines = set()  # Tüm satır için
        
        # İstatistikler
        duplicate_counts = {}
        
        result_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # Boş veya comment satırı
            if not stripped or stripped.startswith('$'):
                result_lines.append(line)
                i += 1
                continue
            
            upper = stripped.upper()
            
            # Hangi tip kart?
            card_type = None
            is_line_based = False
            
            # Önce line-based kartları kontrol et
            for ctype in line_based_cards:
                if upper.startswith(ctype) and (len(upper) == len(ctype) or 
                    upper[len(ctype)] in ' ,\t*'):
                    card_type = ctype
                    is_line_based = True
                    break
            
            # Sonra ID-based kartları kontrol et
            if not card_type:
                for ctype in id_based_cards.keys():
                    if upper.startswith(ctype) and (len(upper) == len(ctype) or 
                        upper[len(ctype)] in ' ,\t*'):
                        card_type = ctype
                        break
            
            if card_type:
                if is_line_based:
                    # LINE-BASED: Tüm satırı (ve continuation'ları) karşılaştır
                    full_card = [line]
                    j = i + 1
                    
                    # Continuation satırlarını topla
                    while j < len(lines):
                        next_line = lines[j]
                        next_stripped = next_line.strip()
                        is_cont = (next_line.startswith('+') or 
                                  next_line.startswith('*') or
                                  (next_line.startswith('        ') and next_stripped and 
                                   not next_stripped.startswith('$') and
                                   not any(next_stripped.upper().startswith(ct) for ct in 
                                          list(id_based_cards.keys()) + line_based_cards)))
                        if is_cont and next_stripped:
                            full_card.append(next_line)
                            j += 1
                        else:
                            break
                    
                    # Normalize edilmiş tam kart
                    normalized_card = '|'.join(' '.join(l.split()) for l in full_card)
                    
                    if normalized_card in seen_full_lines:
                        # Duplicate - atla
                        if card_type not in duplicate_counts:
                            duplicate_counts[card_type] = 0
                        duplicate_counts[card_type] += 1
                        i = j
                        continue
                    else:
                        seen_full_lines.add(normalized_card)
                        result_lines.extend(full_card)
                        i = j
                        continue
                else:
                    # ID-BASED: Sadece ID'ye bak
                    card_id = None
                    try:
                        if ',' in line:
                            parts = line.split(',')
                            if len(parts) > 1:
                                id_str = parts[1].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                        else:
                            if len(line) >= 16:
                                id_str = line[8:16].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                    except:
                        pass
                    
                    if card_id is not None:
                        if card_id in id_based_cards[card_type]:
                            # Duplicate - atla
                            if card_type not in duplicate_counts:
                                duplicate_counts[card_type] = 0
                            duplicate_counts[card_type] += 1
                            
                            # Continuation satırlarını da atla
                            i += 1
                            while i < len(lines):
                                next_line = lines[i]
                                next_stripped = next_line.strip()
                                is_cont = (next_line.startswith('+') or 
                                          next_line.startswith('*') or
                                          (next_line.startswith('        ') and next_stripped and 
                                           not next_stripped.startswith('$') and
                                           not any(next_stripped.upper().startswith(ct) for ct in 
                                                  list(id_based_cards.keys()) + line_based_cards)))
                                if is_cont and next_stripped:
                                    i += 1
                                else:
                                    break
                            continue
                        else:
                            id_based_cards[card_type].add(card_id)
            
            result_lines.append(line)
            i += 1
        
        # Rapor
        if duplicate_counts:
            self.log1("    Removed duplicates:")
            for ctype, count in sorted(duplicate_counts.items()):
                self.log1(f"      {ctype}: {count}")
            total = sum(duplicate_counts.values())
            self.log1(f"      TOTAL: {total} duplicate entries removed")
        else:
            self.log1("    No duplicates found")
        
        return '\n'.join(result_lines)
    
    def merge_lines_with_pynastran(self, lines):
        """
        Satırları temp BDF'e yazıp pyNastran ile merge eder.
        Duplicate ID hatası olursa, satırları direkt yazar.
        Returns: merged bulk data string
        """
        if not lines:
            self.log1("    WARNING: No lines to merge!")
            return ""
        
        temp_bdf_path = os.path.join(tempfile.gettempdir(), "_temp_lines_to_merge.bdf")
        
        self.log1(f"    Writing {len(lines)} lines to temp BDF...")
        
        try:
            # Temp BDF oluştur
            with open(temp_bdf_path, 'w', encoding='utf-8') as f:
                f.write("$ Temporary BDF for merging\n")
                f.write("SOL 101\n")
                f.write("CEND\n")
                f.write("BEGIN BULK\n")
                for line in lines:
                    f.write(line + "\n")
                f.write("ENDDATA\n")
            
            self.log1(f"    Reading temp BDF with pyNastran (following includes)...")
            
            try:
                bdf = BDF(debug=False)
                # allow_duplicate_ids ile dene
                bdf.read_bdf(temp_bdf_path, validate=False, xref=False, 
                            read_includes=True, save_file_structure=False)
                
                self.log1(f"      Loaded: {len(bdf.nodes)} nodes, {len(bdf.elements)} elements")
                self.log1(f"      Properties: {len(bdf.properties)}, Materials: {len(bdf.materials)}")
                self.log1(f"      Coords: {len(bdf.coords)}, MPCs: {len(bdf.mpcs)}, SPCs: {len(bdf.spcs)}")
                
                # Merge edilmiş BDF'i yaz
                merged_temp_path = os.path.join(tempfile.gettempdir(), "_temp_merged.bdf")
                bdf.write_bdf(merged_temp_path, size=8, is_double=False)
                
                with open(merged_temp_path, 'r', errors='ignore') as f:
                    merged_content = f.read()
                
                if os.path.exists(merged_temp_path): os.remove(merged_temp_path)
                
            except Exception as e:
                self.log1(f"    pyNastran failed: {str(e)[:100]}")
                self.log1(f"    Falling back to direct file reading...")
                
                # pyNastran başarısız oldu - INCLUDE'ları manuel aç
                merged_content = self.expand_includes_manually(temp_bdf_path)
            
            # Temizlik
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)
            
            # BEGIN BULK sonrasını al
            bulk_match = re.search(r'BEGIN\s*,?\s*BULK', merged_content, re.IGNORECASE)
            if bulk_match:
                bulk_data = merged_content[bulk_match.end():]
            else:
                bulk_data = merged_content
            
            # Gereksiz satırları temizle
            result_lines = []
            for l in bulk_data.split('\n'):
                if l.startswith('$pyNastran'): continue
                if l.strip().upper().startswith('ENDDATA'): continue
                if l.strip().upper().startswith('INCLUDE'): continue
                if l.strip().upper().startswith('SOL '): continue
                if l.strip().upper().startswith('CEND'): continue
                if l.strip().upper().startswith('BEGIN'): continue
                result_lines.append(l)
            
            result = '\n'.join(result_lines)
            self.log1(f"      Merged bulk data: {len(result)} characters")
            return result
            
        except Exception as e:
            self.log1(f"    ERROR merging: {e}")
            import traceback
            self.log1(traceback.format_exc())
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)
            return ""
    
    def expand_includes_manually(self, bdf_path):
        """
        INCLUDE'ları manuel olarak açar (pyNastran başarısız olduğunda).
        """
        self.log1("    Expanding includes manually...")
        
        content = self.read_file_safe(bdf_path)
        bdf_dir = os.path.dirname(os.path.abspath(bdf_path))
        
        # INCLUDE'ları bul ve aç
        all_includes = self.parse_multiline_includes(content, bdf_dir)
        
        # Satırları işle
        lines = content.split('\n')
        result_lines = []
        
        # INCLUDE satır indekslerini topla
        include_indices = {}
        for inc in all_includes:
            for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                include_indices[idx] = inc
        
        processed_includes = set()
        
        for idx, line in enumerate(lines):
            if idx in include_indices:
                inc = include_indices[idx]
                # Sadece start_idx'te işle (continuation satırlarını atla)
                if idx == inc['start_idx']:
                    inc_path = inc['abs_path']
                    if inc_path not in processed_includes:
                        processed_includes.add(inc_path)
                        # INCLUDE dosyasını oku
                        if os.path.exists(inc_path):
                            try:
                                inc_content = self.read_file_safe(inc_path)
                                result_lines.append(f"$ === EXPANDED: {os.path.basename(inc_path)} ===")
                                for inc_line in inc_content.split('\n'):
                                    # Recursive INCLUDE'ları atla (basit tutuyoruz)
                                    if not inc_line.strip().upper().startswith('INCLUDE'):
                                        result_lines.append(inc_line)
                            except Exception as e:
                                result_lines.append(f"$ ERROR reading {inc_path}: {e}")
                        else:
                            result_lines.append(f"$ FILE NOT FOUND: {inc_path}")
            else:
                result_lines.append(line)
        
        self.log1(f"      Expanded {len(processed_includes)} includes")
        return '\n'.join(result_lines)
    
    def start_processing(self):
        if not self.thermal_bdfs and not self.maneuver_bdfs:
            messagebox.showerror("Error","Add BDF files"); return
        if not self.excel_path.get():
            messagebox.showerror("Error","Select Excel"); return
        if not self.output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.process_btn.config(state=tk.DISABLED)
        self.progress1.start()
        threading.Thread(target=self.process_merge, daemon=True).start()
    
    def process_merge(self):
        try:
            self.log1("="*70)
            self.log1("BDF Merger Tool v8")
            self.log1("="*70)
            
            self.log1("\n[1] Reading Excel...")
            xl = pd.ExcelFile(self.excel_path.get())
            sheets = xl.sheet_names
            self.log1(f"    Available sheets: {sheets}")

            thermal_sh = maneuver_sh = element_sh = node_sh = None

            # First pass: look for exact or specific matches
            for s in sheets:
                sl = s.lower()
                # Element_Set exact match (priority)
                if sl == 'element_set' or sl == 'elementset':
                    element_sh = s
                # Node_Set exact match (priority)
                elif sl == 'node_set' or sl == 'nodeset':
                    node_sh = s
                # Thermal
                elif 'thermal' in sl and not thermal_sh:
                    thermal_sh = s
                # Maneuver
                elif ('maneuver' in sl or 'manevra' in sl) and not maneuver_sh:
                    maneuver_sh = s

            # Second pass: if element_sh/node_sh not found, look for partial matches
            if not element_sh:
                for s in sheets:
                    sl = s.lower()
                    if 'element' in sl and 'set' in sl:
                        element_sh = s
                        break
            if not node_sh:
                for s in sheets:
                    sl = s.lower()
                    if 'node' in sl and 'set' in sl:
                        node_sh = s
                        break

            # Fallback to index-based if still not found
            if not thermal_sh and len(sheets) > 0: thermal_sh = sheets[0]
            if not maneuver_sh and len(sheets) > 1: maneuver_sh = sheets[1]
            if not element_sh and len(sheets) > 2: element_sh = sheets[2]
            if not node_sh and len(sheets) > 3: node_sh = sheets[3]

            self.log1(f"    Using sheets -> Thermal: '{thermal_sh}', Maneuver: '{maneuver_sh}', Element_Set: '{element_sh}', Node_Set: '{node_sh}'")

            thermal_cases = pd.read_excel(xl, sheet_name=thermal_sh).iloc[:,0].dropna().astype(int).tolist() if thermal_sh else []
            maneuver_cases = pd.read_excel(xl, sheet_name=maneuver_sh).iloc[:,0].dropna().astype(int).tolist() if maneuver_sh else []
            element_ids = sorted(pd.read_excel(xl, sheet_name=element_sh).iloc[:,0].dropna().astype(int).tolist()) if element_sh else []
            node_ids = sorted(pd.read_excel(xl, sheet_name=node_sh).iloc[:,0].dropna().astype(int).tolist()) if node_sh else []

            self.log1(f"    Thermal cases: {len(thermal_cases)}")
            self.log1(f"    Maneuver cases: {len(maneuver_cases)}")
            self.log1(f"    Element IDs: {len(element_ids)}")
            self.log1(f"    Node IDs: {len(node_ids)}")

            set_id = int(self.set_id.get())
            node_set_id = int(self.node_set_id.get())
            temp_initial = self.temp_initial.get()
            out_dir = self.output_folder.get()
            os.makedirs(out_dir, exist_ok=True)

            if self.thermal_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[2] Processing THERMAL...")
                self.log1("="*70)
                self.process_thermal_bdf(self.thermal_bdfs, thermal_cases, element_ids, set_id,
                    node_ids, node_set_id, temp_initial, os.path.join(out_dir, self.output_thermal_name.get()))

            if self.maneuver_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[3] Processing MANEUVER...")
                self.log1("="*70)
                self.process_maneuver_bdf(self.maneuver_bdfs, maneuver_cases, element_ids, set_id,
                    node_ids, node_set_id, os.path.join(out_dir, self.output_maneuver_name.get()))
            
            self.log1("\n" + "="*70)
            self.log1("COMPLETED!")
            self.log1("="*70)
            self.root.after(0, lambda: messagebox.showinfo("Done","Merge completed!"))
        except Exception as e:
            self.log1(f"\nERROR: {e}")
            import traceback
            self.log1(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error",str(e)))
        finally:
            self.root.after(0, lambda: [self.progress1.stop(), self.process_btn.config(state=tk.NORMAL)])
    
    def process_thermal_bdf(self, bdf_files, load_cases, element_ids, set_id, node_ids, node_set_id, temp_initial, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        load_case_set = set(load_cases)
        
        # Step 1: Tüm satırları topla
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'THERMAL'
        )
        
        # Step 2: pyNastran ile merge et
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)
        
        # Step 2.5: Duplicate kontrolü
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)
        
        # Step 3: Output dosyası oluştur
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        out.append(f"$ {'='*60}")
        out.append(f"$ THERMAL - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = THERMAL ANALYSIS")
        out.append(f"TEMPERATURE(INITIAL) = {temp_initial}")
        out.append("$")
        
        # SET definition
        chunks = []
        current = ""
        for eid in element_ids:
            test = f"{current},{eid}" if current else str(eid)
            if len(test) > 60 and current:
                chunks.append(current)
                current = str(eid)
            else:
                current = test
        if current: chunks.append(current)

        for i, chunk in enumerate(chunks):
            if i == 0:
                out.append(f"SET {set_id} = {chunk}" + ("," if len(chunks) > 1 else ""))
            elif i == len(chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        # SET definition (nodes)
        node_chunks = []
        current = ""
        for nid in node_ids:
            test = f"{current},{nid}" if current else str(nid)
            if len(test) > 60 and current:
                node_chunks.append(current)
                current = str(nid)
            else:
                current = test
        if current: node_chunks.append(current)

        for i, chunk in enumerate(node_chunks):
            if i == 0:
                out.append(f"SET {node_set_id} = {chunk}" + ("," if len(node_chunks) > 1 else ""))
            elif i == len(node_chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append(f"GPFORCE(PLOT)={node_set_id}")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")

        # SUBCASE definitions
        for lc in load_cases:
            if lc in subcase_info_map:
                info = subcase_info_map[lc]
                temp_load = info['temp_load_id']
                subtitle = info['subtitle']
            else:
                temp_load = lc
                subtitle = f"Thermal Case {lc}"
            out.append(f"SUBCASE {lc}")
            out.append(f"SUBTITLE {subtitle}")
            out.append("SPC = 1")
            out.append(f"TEMPERATURE(LOAD) = {temp_load}")
            out.append("$")
        
        out.append("BEGIN BULK")
        
        # PARAM kartlarını ayır ve BEGIN BULK'tan hemen sonra yaz
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)
        
        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")
        
        # Merged bulk data
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)
        
        # Common Thermal INCLUDE'ları
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON THERMAL INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")
        
        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)
        
        out.append("$")
        out.append("ENDDATA")
        
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))
        
        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON THERMAL INCLUDES: {len(common_includes)}")
    
    def process_maneuver_bdf(self, bdf_files, load_cases, element_ids, set_id, node_ids, node_set_id, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        load_case_set = set(load_cases)
        
        # Step 1: Tüm satırları topla
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'LOAD'
        )
        
        # Step 2: pyNastran ile merge et
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)
        
        # Step 2.5: Duplicate kontrolü
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)
        
        # Step 3: Output dosyası oluştur
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        out.append(f"$ {'='*60}")
        out.append(f"$ MANEUVER - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = MANEUVER ANALYSIS")
        out.append("$")
        
        # SET definition
        chunks = []
        current = ""
        for eid in element_ids:
            test = f"{current},{eid}" if current else str(eid)
            if len(test) > 60 and current:
                chunks.append(current)
                current = str(eid)
            else:
                current = test
        if current: chunks.append(current)

        for i, chunk in enumerate(chunks):
            if i == 0:
                out.append(f"SET {set_id} = {chunk}" + ("," if len(chunks) > 1 else ""))
            elif i == len(chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        # SET definition (nodes)
        node_chunks = []
        current = ""
        for nid in node_ids:
            test = f"{current},{nid}" if current else str(nid)
            if len(test) > 60 and current:
                node_chunks.append(current)
                current = str(nid)
            else:
                current = test
        if current: node_chunks.append(current)

        for i, chunk in enumerate(node_chunks):
            if i == 0:
                out.append(f"SET {node_set_id} = {chunk}" + ("," if len(node_chunks) > 1 else ""))
            elif i == len(node_chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append(f"GPFORCE(PLOT)={node_set_id}")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")

        # SUBCASE definitions
        for lc in load_cases:
            if lc in subcase_info_map:
                info = subcase_info_map[lc]
                load_id = info['load_id']
                subtitle = info['subtitle']
            else:
                load_id = lc
                subtitle = f"Manoeuvre {lc}"
            out.append(f"SUBCASE {lc}")
            out.append(f"SUBTITLE {subtitle}")
            out.append("SPC = 1")
            out.append(f"LOAD = {load_id}")
            out.append("$")
        
        out.append("BEGIN BULK")
        
        # PARAM kartlarını ayır ve BEGIN BULK'tan hemen sonra yaz
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)
        
        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")
        
        # Merged bulk data
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)
        
        # Common Load INCLUDE'ları
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON LOAD INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")
        
        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)
        
        out.append("$")
        out.append("ENDDATA")
        
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))
        
        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON LOAD INCLUDES: {len(common_includes)}")


class BarPropertySolverTab:
    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame
        self.maneuver_bdfs = []
        # self.root.title("Joint Structure Type Understanding")
        # self.root.geometry("1300x950")

        # Input paths
        self.bdf_paths = []
        self.property_excel_path = tk.StringVar()
        self.element_excel_path = tk.StringVar()
        self.nastran_path = tk.StringVar()
        self.nastran_memory = tk.StringVar(value="")
        self.scratch_folder = tk.StringVar(value="")  # optional; blank = each run's own output folder (still collision-free)
        self.output_folder = tk.StringVar()

        # Default bar thickness used only when a property has no Excel/BDF value
        self.default_bar_thickness = 2.0

        # Data storage
        self.bdf_model = None
        self.bdf_models = []
        self.bar_properties = {}          # PID -> {'dim1': val, 'dim2': val} (base model values)
        self.skin_properties = {}         # PID -> {'thickness': val} from 'Skin Property' sheet
        self.bar_structure_map = {}       # PID -> Structure Name
        self.structure_groups = {}        # Structure Name -> [PID list]
        self.pbarl_dims = {}              # PID -> {'dim1': val, 'dim2': val} from BDF
        self.current_bar_thicknesses = {} # PID -> base thickness (from load_properties)
        self.current_skin_thicknesses = {}# PID -> effective thickness (Excel 'Skin Property' if given, else BDF), for base model + offsets
        self.original_skin_thicknesses = {} # PID -> raw PSHELL thickness from BDF (fallback only)
        self.original_bar_thicknesses = {} # PID -> original dim1 from BDF
        self.maneuver_base_path = None    # Excel-properties-applied maneuver (offset source) BDF, set by _build_base_model
        self.material_densities = {}      # MID -> density
        self.prop_to_material = {}        # PID -> MID
        self.element_areas = {}
        self.bar_lengths = {}
        self.prop_elements = {}
        self.elem_to_prop = {}
        self.element_centroids = {}
        self.bar_elements = []
        self.shell_elements = []
        self.landing_elem_ids = []
        self.bar_offset_elem_ids = []

        # Run state
        self.is_running = False
        self.base_stresses = None          # Raw stresses from the base model solve

        self.setup_ui()

    # ==================== GUI ====================
    def setup_ui(self):
        canvas = tk.Canvas(self.parent_frame)
        scrollbar = ttk.Scrollbar(self.parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main = scrollable_frame

        ttk.Label(main, text="Joint Structure Type Understanding", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="Apply Property Excel + offsets and solve the base model",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Input Files ----
        f1 = ttk.LabelFrame(main, text="\U0001F4C1  1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # All BDF Files
        bdf_frame = ttk.Frame(f1)
        bdf_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bdf_frame, text="All BDF Files:", width=18).pack(side=tk.LEFT, anchor=tk.N)

        bdf_list_frame = ttk.Frame(bdf_frame)
        bdf_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bdf_listbox = tk.Listbox(bdf_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                       font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                       selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                       relief='flat', highlightthickness=1,
                                       highlightbackground=_APP_COLORS['border'],
                                       highlightcolor=_APP_COLORS['primary'])
        self.bdf_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        bdf_scroll = ttk.Scrollbar(bdf_list_frame, orient="vertical", command=self.bdf_listbox.yview)
        bdf_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.bdf_listbox.configure(yscrollcommand=bdf_scroll.set)

        bdf_btn_frame = ttk.Frame(bdf_frame)
        bdf_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(bdf_btn_frame, text="Add", command=self.add_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(bdf_btn_frame, text="Remove", command=self.remove_bdf).pack(fill=tk.X, pady=1)

        self.bdf_status = ttk.Label(f1, text="No BDF files loaded", foreground="gray")
        self.bdf_status.pack(anchor=tk.W)

        # Maneuver BDF (Offset Source)
        man_frame = ttk.Frame(f1)
        man_frame.pack(fill=tk.X, pady=2)
        man_lbl = ttk.Frame(man_frame)
        man_lbl.pack(side=tk.LEFT, anchor=tk.N)
        ttk.Label(man_lbl, text="Maneuver BDF:", width=18).pack(anchor=tk.W)
        ttk.Label(man_lbl, text="  (Offset Source)", foreground="blue").pack(anchor=tk.W)

        man_list_frame = ttk.Frame(man_frame)
        man_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.maneuver_listbox = tk.Listbox(man_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                            font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                            selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                            relief='flat', highlightthickness=1,
                                            highlightbackground=_APP_COLORS['border'],
                                            highlightcolor=_APP_COLORS['primary'])
        self.maneuver_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        man_scroll = ttk.Scrollbar(man_list_frame, orient="vertical", command=self.maneuver_listbox.yview)
        man_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.maneuver_listbox.configure(yscrollcommand=man_scroll.set)

        man_btn_frame = ttk.Frame(man_frame)
        man_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(man_btn_frame, text="Add", command=self._add_maneuver_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(man_btn_frame, text="Remove", command=self._remove_maneuver_bdf).pack(fill=tk.X, pady=1)


        # Property Excel
        self._add_file_row(f1, "Property Excel:", self.property_excel_path, self.load_properties)
        self.prop_status = ttk.Label(f1, text="Not loaded", foreground="gray")
        self.prop_status.pack(anchor=tk.W)

        # Offset Element IDs
        self._add_file_row(f1, "Offset Element IDs:", self.element_excel_path, self.load_element_ids)
        self.elem_status = ttk.Label(f1, text="Not loaded", foreground="gray")
        self.elem_status.pack(anchor=tk.W)

        # Nastran Exe
        nast_frame = ttk.Frame(f1)
        nast_frame.pack(fill=tk.X, pady=2)
        ttk.Label(nast_frame, text="Nastran Exe:", width=18).pack(side=tk.LEFT)
        ttk.Entry(nast_frame, textvariable=self.nastran_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(nast_frame, text="Browse", command=lambda: self.nastran_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)

        # NX Nastran Requested Memory (optional user-imposed cap, e.g. "4gb", "500mb")
        mem_frame = ttk.Frame(f1)
        mem_frame.pack(fill=tk.X, pady=2)
        ttk.Label(mem_frame, text="Requested Memory:", width=18).pack(side=tk.LEFT)
        ttk.Entry(mem_frame, textvariable=self.nastran_memory, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(mem_frame, text="(optional, e.g. 4gb / 500mb - blank = Nastran default)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # NX Nastran Scratch Folder (optional; each run gets its own unique
        # subfolder under it, so parallel runs never collide on scratch files)
        scr_frame = ttk.Frame(f1)
        scr_frame.pack(fill=tk.X, pady=2)
        ttk.Label(scr_frame, text="Scratch Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(scr_frame, textvariable=self.scratch_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(scr_frame, text="Browse", command=lambda: self.scratch_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)
        scr_hint = ttk.Frame(f1)
        scr_hint.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(scr_hint, text="(recommended if C: is low on space - blank = Nastran uses its own default "
                                  "temp location. If set, each run gets its own subfolder here, set as that "
                                  "run's TEMP/TMP, so scratch never lands on C: or collides between runs)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Output Folder
        out_frame = ttk.Frame(f1)
        out_frame.pack(fill=tk.X, pady=2)
        ttk.Label(out_frame, text="Output Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="Browse", command=lambda: self.output_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

        # ---- Section 2: Actions ----
        f3 = ttk.LabelFrame(main, text="▶  2. Actions", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        btn_row = ttk.Frame(f3)
        btn_row.pack(fill=tk.X)
        self.btn_start = ttk.Button(btn_row, text="▶  SOLVE BASE MODEL", command=self.start_solve,
                                     state=tk.DISABLED, style='Accent.TButton')
        self.btn_start.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(f3, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ---- Section 3: Structure Groups ----
        f4 = ttk.LabelFrame(main, text="\U0001F9E9  3. Structure Groups (from Property Excel)", padding=10)
        f4.pack(fill=tk.X, pady=5, padx=10)

        self.group_text = scrolledtext.ScrolledText(f4, height=6, width=100, state=tk.DISABLED,
                                                      font=('Consolas', 10),
                                                      bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                                      relief='flat', borderwidth=0)
        self.group_text.pack(fill=tk.X)

        # ---- Section 4: Log ----
        f5 = ttk.LabelFrame(main, text="\U0001F4DC  4. Log", padding=10)
        f5.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f5, height=20, width=100, font=('Consolas', 10),
                                                    bg=_APP_COLORS['log_bg'], fg=_APP_COLORS['log_fg'],
                                                    insertbackground=_APP_COLORS['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _add_file_row(self, parent, label, var, load_cmd):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda: var.set(
            filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Button(frame, text="Load", command=load_cmd).pack(side=tk.LEFT, padx=2)

    def log(self, msg):
        def _do():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ==================== BDF FILE OPS ====================

    def _add_maneuver_bdf(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, os.path.basename(f))

    def _remove_maneuver_bdf(self):
        sel = self.maneuver_listbox.curselection()
        if sel:
            idx = sel[0]
            self.maneuver_listbox.delete(idx)
            del self.maneuver_bdfs[idx]

    def add_bdf(self):
        paths = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for p in paths:
            if p not in self.bdf_paths:
                self.bdf_paths.append(p)
                self.bdf_listbox.insert(tk.END, os.path.basename(p))
        self.bdf_status.config(text=f"{len(self.bdf_paths)} BDF files selected", foreground="blue")

    def remove_bdf(self):
        selection = self.bdf_listbox.curselection()
        if selection:
            idx = selection[0]
            self.bdf_listbox.delete(idx)
            del self.bdf_paths[idx]
        self.bdf_status.config(text=f"{len(self.bdf_paths)} BDF files selected", foreground="blue")

    # ==================== BDF LOADING ====================
    def load_bdf(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one BDF file")
            return

        self.log("\n" + "=" * 70)
        self.log("LOADING BDF FILES")
        self.log("=" * 70)
        self.log(f"  BDF files to load: {len(self.bdf_paths)}")
        for i, p in enumerate(self.bdf_paths):
            self.log(f"    {i + 1}. {os.path.basename(p)}")

        try:
            path = self.bdf_paths[0]
            self.log(f"\n  Loading main BDF: {os.path.basename(path)}")
            self.bdf_model = BDF(debug=False)
            self.bdf_model.read_bdf(path, validate=False, xref=True, read_includes=False, encoding='latin-1')

            self.bdf_models = []
            for bdf_path in self.bdf_paths:
                self.log(f"  Loading: {os.path.basename(bdf_path)}")
                model = BDF(debug=False)
                model.read_bdf(bdf_path, validate=False, xref=True, read_includes=False, encoding='latin-1')
                self.bdf_models.append({'path': bdf_path, 'model': model, 'name': os.path.basename(bdf_path)})

            self.log(f"  Nodes: {len(self.bdf_model.nodes)}")
            self.log(f"  Elements: {len(self.bdf_model.elements)}")
            self.log(f"  Properties: {len(self.bdf_model.properties)}")
            self.log(f"  Materials: {len(self.bdf_model.materials)}")

            # Extract material densities
            self.material_densities = {}
            for mid, mat in self.bdf_model.materials.items():
                rho = None
                if hasattr(mat, 'rho') and mat.rho is not None:
                    rho = mat.rho
                elif hasattr(mat, 'Rho') and mat.Rho is not None:
                    rho = mat.Rho()
                if rho:
                    self.material_densities[mid] = rho
                    self.log(f"    Material {mid} ({mat.type}): density = {rho}")

            # Property -> Material mapping
            self.prop_to_material = {}
            for pid, prop in self.bdf_model.properties.items():
                mid = None
                if hasattr(prop, 'mid') and prop.mid:
                    mid = prop.mid if isinstance(prop.mid, int) else prop.mid.mid
                elif hasattr(prop, 'mid1') and prop.mid1:
                    mid = prop.mid1 if isinstance(prop.mid1, int) else prop.mid1.mid
                elif hasattr(prop, 'mid_ref') and prop.mid_ref:
                    mid = prop.mid_ref.mid
                if mid:
                    self.prop_to_material[pid] = mid

            # Element geometry
            self.element_areas = {}
            self.bar_lengths = {}
            self.prop_elements = {}
            self.elem_to_prop = {}
            self.element_centroids = {}
            self.bar_elements = []
            self.shell_elements = []

            shell_count = bar_count = 0
            for eid, elem in self.bdf_model.elements.items():
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid:
                    self.elem_to_prop[eid] = pid
                    if pid not in self.prop_elements:
                        self.prop_elements[pid] = []
                    self.prop_elements[pid].append(eid)

                try:
                    centroid = elem.Centroid()
                    self.element_centroids[eid] = centroid
                except:
                    pass

                if elem.type in ['CQUAD4', 'CTRIA3', 'CQUAD8', 'CTRIA6']:
                    shell_count += 1
                    self.shell_elements.append(eid)
                    try:
                        self.element_areas[eid] = elem.Area()
                    except:
                        self.element_areas[eid] = 0
                elif elem.type in ['CBAR', 'CBEAM']:
                    bar_count += 1
                    self.bar_elements.append(eid)
                    try:
                        self.bar_lengths[eid] = elem.Length()
                    except:
                        self.bar_lengths[eid] = 0

            # Extract PBARL dimensions from BDF
            self.pbarl_dims = {}
            for pid, prop in self.bdf_model.properties.items():
                if prop.type == 'PBARL':
                    dims = prop.dim if hasattr(prop, 'dim') else []
                    if len(dims) >= 2:
                        self.pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[1]}
                    elif len(dims) == 1:
                        self.pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[0]}

            # Store original dim1 from BDF as the original thicknesses
            self.original_bar_thicknesses = {}
            if self.pbarl_dims:
                self.log(f"  PBARL dimensions extracted: {len(self.pbarl_dims)} properties")
                for pid, d in self.pbarl_dims.items():
                    self.original_bar_thicknesses[pid] = d['dim1']
                    self.log(f"    PID {pid}: dim1={d['dim1']}, dim2={d['dim2']}")

            # Extract PSHELL thicknesses from BDF (fallback for offset calc when
            # a PID isn't in the Excel 'Skin Property' sheet). Stored separately
            # from self.current_skin_thicknesses (the effective value actually
            # used) and merged via _refresh_effective_skin_thicknesses, so
            # whichever of load_bdf()/load_properties() the user happens to run
            # second never silently discards the other's data - there is no
            # "Load BDF" button, so this runs lazily from start_solve() and can
            # easily land AFTER load_properties() in the normal click order.
            self.original_skin_thicknesses = {}
            for pid, prop in self.bdf_model.properties.items():
                if prop.type == 'PSHELL':
                    t = prop.t if hasattr(prop, 't') and prop.t is not None else None
                    if t is not None:
                        self.original_skin_thicknesses[pid] = t
            if self.original_skin_thicknesses:
                self.log(f"  PSHELL thicknesses extracted: {len(self.original_skin_thicknesses)} properties (fallback for offset calc)")
            self._refresh_effective_skin_thicknesses()

            self.log(f"  Shells: {shell_count}, Bars: {bar_count}")
            self.log(f"  Centroids calculated: {len(self.element_centroids)}")
            self.log(f"\n  Total BDF models loaded: {len(self.bdf_models)}")

            self.bdf_status.config(
                text=f"Loaded: {len(self.bdf_models)} BDFs, {len(self.bdf_model.elements)} elements",
                foreground="green"
            )

            if not self.output_folder.get():
                self.output_folder.set(os.path.dirname(path))

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.bdf_status.config(text="Error", foreground="red")

    # ==================== PROPERTY LOADING ====================
    def load_properties(self):
        """Load structure grouping ('Bar Property Structure Type' sheet) plus the
        Excel-defined base Dim1/Dim2 ('Bar Property' sheet) and skin thickness
        ('Skin Property' sheet). These base values are what the base model gets
        written with (see _build_base_model)."""
        path = self.property_excel_path.get()
        if not path:
            messagebox.showerror("Error", "Select Property Excel")
            return

        self.log("\n" + "=" * 70)
        self.log("LOADING PROPERTIES")
        self.log("=" * 70)

        try:
            xl = pd.ExcelFile(path)
            self.log(f"  Sheets: {xl.sheet_names}")

            bar_min = self.default_bar_thickness

            # ---- Locate the 3 sheets we care about ----
            struct_sheet = bar_dim_sheet = skin_sheet = None
            for sheet in xl.sheet_names:
                sl = sheet.lower().replace('_', '').replace(' ', '').replace('-', '')
                if 'bar' in sl and 'prop' in sl and 'structure' in sl:
                    struct_sheet = sheet
                elif sl == 'barproperty':
                    bar_dim_sheet = sheet
                elif sl == 'skinproperty':
                    skin_sheet = sheet
            # Fallback partial match for Bar/Skin Property sheets (skip the structure sheet)
            for sheet in xl.sheet_names:
                if sheet == struct_sheet:
                    continue
                sl = sheet.lower().replace('_', '').replace(' ', '').replace('-', '')
                if bar_dim_sheet is None and 'bar' in sl and 'prop' in sl:
                    bar_dim_sheet = sheet
                elif skin_sheet is None and 'skin' in sl and 'prop' in sl:
                    skin_sheet = sheet

            if struct_sheet is None:
                raise ValueError("No 'Bar Property Structure Type' sheet found "
                                  "(expects sheet/column names containing 'bar', 'prop', 'structure')")

            # ---- 1. Excel base Dim1/Dim2 ('Bar Property' sheet: PID, Dim1, Dim2) ----
            excel_bar_dims = {}
            if bar_dim_sheet:
                self.log(f"\n  Reading base bar Dim1/Dim2 from '{bar_dim_sheet}'...")
                df_bar = pd.read_excel(xl, sheet_name=bar_dim_sheet)
                for _, row in df_bar.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        d1 = float(row.iloc[1])
                        d2 = float(row.iloc[2]) if len(df_bar.columns) > 2 else d1
                        excel_bar_dims[pid] = {'dim1': d1, 'dim2': d2}
                    except Exception:
                        pass
                self.log(f"    Base bar dimensions loaded: {len(excel_bar_dims)} properties")
            else:
                self.log("  WARNING: No 'Bar Property' sheet found - bar base thickness "
                          "will fall back to BDF PBARL values where possible.")

            # ---- 2. Excel base skin thickness ('Skin Property' sheet: PID, Thickness) ----
            self.skin_properties = {}
            if skin_sheet:
                self.log(f"\n  Reading base skin thickness from '{skin_sheet}'...")
                df_skin = pd.read_excel(xl, sheet_name=skin_sheet)
                for _, row in df_skin.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        t = float(row.iloc[1])
                        self.skin_properties[pid] = {'thickness': t}
                    except Exception:
                        pass
                self.log(f"    Base skin thicknesses loaded: {len(self.skin_properties)} properties")
            else:
                self.log("  WARNING: No 'Skin Property' sheet found - skin thickness "
                          "will fall back to BDF PSHELL values (unchanged).")
            # Recompute the effective thickness (Excel first, BDF fallback) now,
            # regardless of whether load_bdf() has run yet or not.
            self._refresh_effective_skin_thicknesses()

            # ---- 3. Structure grouping ('Bar Property Structure Type' sheet) ----
            self.log(f"\n  Reading structure groups from '{struct_sheet}'...")
            df = pd.read_excel(xl, sheet_name=struct_sheet)
            self.log(f"    Columns: {list(df.columns)}")

            # Find columns by name
            pid_col = None
            struct_col = None
            for col in df.columns:
                col_clean = str(col).lower().replace('_', '').replace(' ', '')
                if 'barproperty' in col_clean or 'propertyid' in col_clean or col_clean == 'barpropid':
                    pid_col = col
                elif 'structure' in col_clean or 'structurename' in col_clean:
                    struct_col = col

            # Fallback to positional
            if pid_col is None:
                pid_col = df.columns[0]
                self.log(f"    Using first column as PID: {pid_col}")
            if struct_col is None and len(df.columns) > 1:
                struct_col = df.columns[1]
                self.log(f"    Using second column as Structure Name: {struct_col}")

            self.log(f"    PID column: {pid_col}")
            self.log(f"    Structure column: {struct_col}")

            self.bar_properties = {}
            self.bar_structure_map = {}
            self.structure_groups = {}
            self.current_bar_thicknesses = {}

            for _, row in df.iterrows():
                pid_val = row[pid_col]
                if pd.isna(pid_val):
                    continue
                pid = int(pid_val)
                struct_name = str(row[struct_col]).strip() if struct_col and pd.notna(row[struct_col]) else "DEFAULT"

                # Base Dim1/Dim2: prefer the Excel 'Bar Property' sheet, else fall back to BDF PBARL
                if pid in excel_bar_dims:
                    dim1 = excel_bar_dims[pid]['dim1']
                    dim2 = excel_bar_dims[pid]['dim2']
                else:
                    dim1 = self.original_bar_thicknesses.get(pid, bar_min)
                    dim2 = self.pbarl_dims[pid]['dim2'] if pid in self.pbarl_dims else bar_min

                self.bar_properties[pid] = {
                    'dim1': dim1,
                    'dim2': dim2,
                }
                # Initialize to the base value (Excel-derived where available)
                self.current_bar_thicknesses[pid] = dim1
                self.bar_structure_map[pid] = struct_name

                if struct_name not in self.structure_groups:
                    self.structure_groups[struct_name] = []
                self.structure_groups[struct_name].append(pid)

            self.log(f"  Loaded {len(self.bar_properties)} bar properties")
            self.log(f"  Structure groups: {len(self.structure_groups)}")
            for name, pids in sorted(self.structure_groups.items()):
                self.log(f"    {name}: {len(pids)} properties")

            from_excel = sum(1 for pid in self.bar_properties if pid in excel_bar_dims)
            self.log(f"  Base Dim1/Dim2 source: {from_excel} from 'Bar Property' sheet, "
                      f"{len(self.bar_properties) - from_excel} from BDF (fallback)")

            # Update group display
            self._update_group_display()

            total = len(self.bar_properties)
            groups = len(self.structure_groups)
            self.prop_status.config(
                text=f"Loaded: {total} bar props / {groups} groups, {len(self.skin_properties)} skin props",
                foreground="green"
            )
            # Excel (structure groups) is what the solve actually needs - only
            # enable the solve button once it's successfully loaded.
            self.btn_start.config(state=tk.NORMAL if self.structure_groups else tk.DISABLED)

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.prop_status.config(text="Error", foreground="red")
            self.btn_start.config(state=tk.DISABLED)

    def _update_group_display(self):
        self.group_text.config(state=tk.NORMAL)
        self.group_text.delete(1.0, tk.END)
        for name in sorted(self.structure_groups.keys()):
            pids = self.structure_groups[name]
            pid_str = ", ".join(str(p) for p in sorted(pids)[:15])
            if len(pids) > 15:
                pid_str += f" ... (+{len(pids) - 15} more)"
            self.group_text.insert(tk.END, f"{name} ({len(pids)} props): {pid_str}\n")
        self.group_text.config(state=tk.DISABLED)

    def _refresh_effective_skin_thicknesses(self):
        """Recompute self.current_skin_thicknesses - the effective value used
        for offset calc - as: Excel 'Skin Property' sheet value if given,
        else the BDF's own PSHELL thickness (self.original_skin_thicknesses).
        There is no dedicated "Load BDF" button - load_bdf() only runs lazily
        from start_solve() - so load_bdf() and load_properties() can each run
        in either order relative to the other. Calling this from both after
        they update their own half of the data means neither one's data ever
        gets silently discarded by the other running afterward."""
        merged = dict(self.original_skin_thicknesses)
        for pid, d in self.skin_properties.items():
            merged[pid] = d['thickness']
        self.current_skin_thicknesses = merged

    # ==================== ELEMENT IDs (OFFSET) ====================
    def load_element_ids(self):
        path = self.element_excel_path.get()
        if not path:
            return

        self.log("\n" + "=" * 70)
        self.log("LOADING ELEMENT IDs FOR OFFSET")
        self.log("=" * 70)

        try:
            xl = pd.ExcelFile(path)
            self.landing_elem_ids = []
            self.bar_offset_elem_ids = []

            for s in xl.sheet_names:
                sl = s.lower().replace('_', '').replace(' ', '')
                df = pd.read_excel(xl, sheet_name=s)
                if 'landing' in sl:
                    self.landing_elem_ids = df.iloc[:, 0].dropna().astype(int).tolist()
                    self.log(f"  Landing: {len(self.landing_elem_ids)}")
                elif 'bar' in sl and 'offset' in sl:
                    self.bar_offset_elem_ids = df.iloc[:, 0].dropna().astype(int).tolist()
                    self.log(f"  Bar offset: {len(self.bar_offset_elem_ids)}")

            self.elem_status.config(
                text=f"Landing: {len(self.landing_elem_ids)}, Bar offset: {len(self.bar_offset_elem_ids)}",
                foreground="green"
            )

        except Exception as e:
            self.log(f"ERROR: {e}")
            self.elem_status.config(text="Error", foreground="red")

    # ==================== BDF WRITING ====================
    def _write_bdf_for_model(self, folder, model, original_path, thickness_overrides):
        """thickness_overrides: {pid: new_dim1}. Only these PBARL PIDs are
        rewritten; every other card in original_path (the base model) is left
        untouched. Taking the overrides as an explicit argument - instead of
        reading a shared self.current_bar_thicknesses/active_group_pids state -
        is what makes this safe to call from multiple threads at once."""
        output_bdf = os.path.join(folder, os.path.basename(original_path))

        with open(original_path, 'r', encoding='latin-1') as f:
            lines = f.readlines()

        new_lines = []
        i = 0
        while i < len(lines):
            line = lines[i]

            if line.startswith('PBARL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in thickness_overrides:
                        t = max(0.1, thickness_overrides[pid])
                        new_lines.append(line)
                        i += 1
                        while i < len(lines) and (
                            lines[i].startswith('+') or lines[i].startswith('*') or
                            (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))
                        ):
                            cont = lines[i]
                            if cont.strip() and not cont.strip().startswith('$'):
                                try:
                                    original_dim2 = cont[16:24].strip()
                                    dim2_val = float(original_dim2) if original_dim2 else t
                                except:
                                    dim2_val = t
                                cont_name = cont[:8]
                                rest = cont[24:] if len(cont) > 24 else '\n'
                                new_cont = f"{cont_name}{t:8.4f}{dim2_val:8.4f}{rest}"
                                new_lines.append(new_cont)
                            else:
                                new_lines.append(cont)
                            i += 1
                        continue
                except:
                    pass

            new_lines.append(line)
            i += 1

        with open(output_bdf, 'w', encoding='latin-1') as f:
            f.writelines(new_lines)

        self.log(f"    BDF written: {os.path.basename(output_bdf)}")
        return output_bdf

    # ==================== BASE MODEL (Excel Bar/Skin properties) ====================
    def _apply_excel_properties(self, source_path, folder):
        """Write a copy of source_path with ALL bar Dim1/Dim2 ('Bar Property' sheet)
        and ALL skin thicknesses ('Skin Property' sheet) applied - i.e. every PID,
        not just the active sweep group. This is the base model that the sweep's
        per-group thickness changes are layered on top of.

        Named "..._prop_updated" (rather than keeping source_path's own name)
        so it's never confused with the later, actually-offset file - a
        source file already named "..._offseted.bdf" (e.g. carried over from
        an earlier processing stage) would otherwise make this property-only
        step look like it had already applied offsets, and the real offset
        step's own "_offseted" suffix on top of that produced a misleading
        "..._offseted_offseted" name. Any existing trailing "_offseted" on
        the source's own name is stripped first, so e.g.
        "merged_maneuver_offseted.bdf" becomes "merged_maneuver_prop_updated.bdf"
        here and "merged_maneuver_prop_updated_offseted.bdf" once actually
        offset - not "merged_maneuver_offseted_prop_updated...", which just
        carries the source's stale suffix along for no reason."""
        base, ext = os.path.splitext(os.path.basename(source_path))
        if base.endswith('_offseted'):
            base = base[:-len('_offseted')]
        output_bdf = os.path.join(folder, base + "_prop_updated" + ext)

        with open(source_path, 'r', encoding='latin-1') as f:
            lines = f.readlines()

        new_lines = []
        i = 0
        bar_updated = 0
        skin_updated = 0
        while i < len(lines):
            line = lines[i]

            if line.startswith('PBARL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in self.bar_properties:
                        d1 = self.bar_properties[pid]['dim1']
                        d2 = self.bar_properties[pid]['dim2']
                        new_lines.append(line)
                        i += 1
                        if i < len(lines) and (
                            lines[i].startswith('+') or lines[i].startswith('*') or
                            (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))
                        ):
                            cont = lines[i]
                            cont_name = cont[:8]
                            rest = cont[24:] if len(cont) > 24 else '\n'
                            new_lines.append(f"{cont_name}{d1:8.4f}{d2:8.4f}{rest}")
                            i += 1
                        bar_updated += 1
                        continue
                except:
                    pass

            elif line.startswith('PSHELL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in self.skin_properties:
                        t = self.skin_properties[pid]['thickness']
                        rest = line[32:] if len(line) > 32 else '\n'
                        new_lines.append(line[:24] + f"{t:8.4f}" + rest)
                        skin_updated += 1
                        i += 1
                        continue
                except:
                    pass

            new_lines.append(line)
            i += 1

        with open(output_bdf, 'w', encoding='latin-1') as f:
            f.writelines(new_lines)

        self.log(f"    {os.path.basename(source_path)}: {bar_updated} bar props, "
                  f"{skin_updated} skin props set to Excel base values")
        return output_bdf

    def _build_base_model(self, run_folder):
        """Apply the Excel base Bar/Skin properties to every loaded BDF and to
        the maneuver (offset source) BDF. These base files carry ONLY the
        Excel properties - no offsets. Offsets are computed and applied fresh
        for every iteration (baseline included) by _prepare_iteration_bdfs,
        since they depend on the maneuver BDF's geometry at that iteration's
        thickness, not on some value frozen once up front."""
        base_folder = os.path.join(run_folder, "_base_model")
        os.makedirs(base_folder, exist_ok=True)

        self.log("\n" + "=" * 70)
        self.log("BUILDING BASE MODEL (Excel Bar/Skin Property thicknesses)")
        self.log("=" * 70)

        self.maneuver_base_path = None
        if self.maneuver_bdfs:
            man_folder = os.path.join(base_folder, "maneuver")
            os.makedirs(man_folder, exist_ok=True)
            self.maneuver_base_path = self._apply_excel_properties(self.maneuver_bdfs[0], man_folder)
            self.log(f"  Base maneuver (offset source) BDF ready: {os.path.basename(self.maneuver_base_path)}")
        else:
            self.log("  No maneuver BDF provided - no offsets will be computed.")

        for bdf_idx, bdf_info in enumerate(self.bdf_models):
            bdf_folder = os.path.join(base_folder, f"bdf_{bdf_idx + 1}_{os.path.splitext(bdf_info['name'])[0]}")
            os.makedirs(bdf_folder, exist_ok=True)
            base_path = self._apply_excel_properties(bdf_info['path'], bdf_folder)
            bdf_info['base_path'] = base_path
            self.log(f"  Base model ready for {bdf_info['name']}: {os.path.basename(base_path)}")

        # Run Nastran + extract stresses on the base model. thickness_overrides={}
        # means every PID uses its self.bar_properties base value - offsets get
        # computed fresh from the maneuver BDF at that same base thickness.
        self.log("\n  Running Nastran + stress extraction on the BASE MODEL...")
        base_stresses = self._run_single_iteration(base_folder, {}, label="[BASE] ")
        self.base_stresses = base_stresses
        if base_stresses:
            bar_s = [s for s in base_stresses if s['type'] == 'bar']
            shell_s = [s for s in base_stresses if s['type'] == 'shell']
            self.log(f"  Base model stresses: {len(bar_s)} bar, {len(shell_s)} shell")
        else:
            self.log("  WARNING: no stresses extracted from the base model (check Nastran path/run).")
        self.log(f"  Base model results saved under: {base_folder}")

        self.log("BASE MODEL SOLVE COMPLETE.\n")

    # ==================== OFFSET CALCULATION / APPLICATION ====================
    def _calculate_offset_csv(self, thickness_overrides, folder, label=""):
        """Write the maneuver (offset source) BDF at THIS iteration's
        thickness_overrides (layered on its Excel-properties base, same as any
        other BDF), then compute landing/bar offsets from that freshly-written
        geometry and save them to offsets.csv in `folder`. Called once per
        iteration (sweep step or baseline) - not once for the whole sweep -
        since the maneuver geometry (and therefore the offsets) changes with
        whichever group's thickness is being tested. Returns the CSV path, or
        None if there's no maneuver BDF or no offset elements configured."""
        if not self.maneuver_base_path:
            return None
        if not self.landing_elem_ids and not self.bar_offset_elem_ids:
            return None

        man_folder = os.path.join(folder, "maneuver_for_offset")
        os.makedirs(man_folder, exist_ok=True)
        maneuver_path = self._write_bdf_for_model(man_folder, None, self.maneuver_base_path, thickness_overrides)

        try:
            bdf = BDF(debug=False)
            bdf.read_bdf(maneuver_path, validate=False, xref=True, read_includes=True, encoding='latin-1')
        except Exception as e:
            self.log(f"    {label}Offset calc error reading maneuver BDF: {e}")
            return None

        try:
            # Landing (shell) offsets: zoffset = -t/2
            landing_normals = {}
            rows = []

            for eid in self.landing_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                t = self.current_skin_thicknesses.get(pid, 3.0)
                t = max(0.1, t)
                zoffset = -t / 2.0

                normal = None
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    try:
                        nids = elem.node_ids[:3]
                        nodes = [bdf.nodes[n] for n in nids if n in bdf.nodes]
                        if len(nodes) >= 3:
                            p1, p2, p3 = [np.array(n.get_position()) for n in nodes]
                            nvec = np.cross(p2 - p1, p3 - p1)
                            nl = np.linalg.norm(nvec)
                            if nl > 1e-10:
                                normal = nvec / nl
                    except Exception:
                        pass
                landing_normals[eid] = normal
                rows.append({'Type': 'Landing', 'Element': eid, 'ZOffset': zoffset,
                             'WA_X': '', 'WA_Y': '', 'WA_Z': ''})

            # Node -> shell mapping for bar offset
            node_to_shells = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    for nid in elem.node_ids:
                        node_to_shells.setdefault(nid, []).append(eid)

            # Bar offsets: WA = WB = -normal * (landing_t + bar_t/2)
            bar_min = self.default_bar_thickness
            for eid in self.bar_offset_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                if elem.type not in ['CBAR', 'CBEAM']:
                    continue
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                bar_t = thickness_overrides.get(pid, self.bar_properties.get(pid, {}).get('dim1', bar_min))
                bar_t = max(0.1, bar_t)

                bar_nodes = elem.node_ids[:2]
                if bar_nodes[0] in node_to_shells and bar_nodes[1] in node_to_shells:
                    common = set(node_to_shells[bar_nodes[0]]) & set(node_to_shells[bar_nodes[1]])
                    max_t = 0
                    best_normal = None
                    for shell_eid in common:
                        if shell_eid in landing_normals:
                            shell_elem = bdf.elements[shell_eid]
                            shell_pid = shell_elem.pid if hasattr(shell_elem, 'pid') else None
                            if shell_pid:
                                shell_t = self.current_skin_thicknesses.get(shell_pid, 0)
                                if shell_t > max_t:
                                    max_t = shell_t
                                    best_normal = landing_normals.get(shell_eid)

                    if best_normal is not None and max_t > 0:
                        offset_mag = max_t + bar_t / 2.0
                        vec = -best_normal * offset_mag
                        rows.append({'Type': 'Bar', 'Element': eid, 'ZOffset': '',
                                     'WA_X': vec[0], 'WA_Y': vec[1], 'WA_Z': vec[2]})

            if not rows:
                return None

            n_landing = sum(1 for r in rows if r['Type'] == 'Landing')
            n_bar = sum(1 for r in rows if r['Type'] == 'Bar')
            csv_path = os.path.join(folder, 'offsets.csv')
            with open(csv_path, 'w', newline='') as f:
                w = csv.DictWriter(f, fieldnames=['Type', 'Element', 'ZOffset', 'WA_X', 'WA_Y', 'WA_Z'])
                w.writeheader()
                w.writerows(rows)
            self.log(f"    {label}Offset CSV: {n_landing} landing, {n_bar} bar -> {csv_path}")
            return csv_path

        except Exception as e:
            self.log(f"    {label}Offset calc error: {e}")
            return None

    def _apply_offsets_from_csv(self, bdf_path, csv_path, folder, label=""):
        """Read an offsets.csv (from _calculate_offset_csv) and write a copy of
        bdf_path with those ZOFFSET/WA-WB values applied, saved as
        <name>_offseted.bdf in `folder`."""
        landing_offsets = {}
        bar_offsets = {}
        try:
            with open(csv_path, 'r', newline='') as f:
                for row in csv.DictReader(f):
                    eid = int(row['Element'])
                    if row['Type'] == 'Landing':
                        landing_offsets[eid] = float(row['ZOffset'])
                    elif row['Type'] == 'Bar':
                        bar_offsets[eid] = (float(row['WA_X']), float(row['WA_Y']), float(row['WA_Z']))
        except Exception as e:
            self.log(f"    {label}Could not read offsets CSV: {e}")
            return bdf_path

        try:
            with open(bdf_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()

            def fmt(v, w=8):
                s = f"{v:.4f}"
                return s[:w].ljust(w) if len(s) <= w else f"{v:.2E}"[:w].ljust(w)

            new_lines = []
            i = 0
            applied_landing = 0
            applied_bar = 0

            while i < len(lines):
                line = lines[i]

                if line.startswith('CQUAD4'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in landing_offsets:
                            padded = line.rstrip().ljust(72)
                            new_lines.append(padded[:64] + fmt(landing_offsets[eid]) + '\n')
                            applied_landing += 1
                            i += 1
                            continue
                    except:
                        pass

                elif line.startswith('CTRIA3'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in landing_offsets:
                            padded = line.rstrip().ljust(56)
                            new_lines.append(padded[:48] + fmt(landing_offsets[eid]) + '\n')
                            applied_landing += 1
                            i += 1
                            continue
                    except:
                        pass

                elif line.startswith('CBAR'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in bar_offsets:
                            offset_vec = bar_offsets[eid]

                            if i + 1 < len(lines) and (
                                lines[i + 1].startswith('+') or lines[i + 1].startswith('*') or
                                (lines[i + 1][0] == ' ' and lines[i + 1].strip())
                            ):
                                cont_line = lines[i + 1]
                                if len(cont_line) < 24:
                                    cont_line = cont_line.rstrip().ljust(24)
                                new_cont = cont_line[:24]
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += '\n'

                                new_lines.append(line)
                                new_lines.append(new_cont)
                                applied_bar += 1
                                i += 2
                                continue
                            else:
                                cont_name = '+CB' + str(eid)[-4:]
                                new_lines.append(line.rstrip() + cont_name + '\n')

                                new_cont = cont_name.ljust(8)
                                new_cont += '        '   # PA
                                new_cont += '        '   # PB
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += '\n'
                                new_lines.append(new_cont)

                                applied_bar += 1
                                i += 1
                                continue
                    except:
                        pass

                new_lines.append(line)
                i += 1

            base, ext = os.path.splitext(os.path.basename(bdf_path))
            output_bdf = os.path.join(folder, base + "_offseted" + ext)
            with open(output_bdf, 'w', encoding='latin-1') as f:
                f.writelines(new_lines)

            self.log(f"    {label}Offsets applied from CSV: {applied_landing} landing, {applied_bar} bar")
            return output_bdf

        except Exception as e:
            self.log(f"    {label}Offset apply error: {e}")
            return bdf_path

    def _prepare_iteration_bdfs(self, iter_folder, thickness_overrides, label=""):
        """Write every bdf_model's BDF at thickness_overrides (layered on its
        Excel-properties base), compute a FRESH offsets.csv from the maneuver
        (source) BDF at this same thickness_overrides, and apply those offsets
        to every written BDF. Used for both the baseline run and every sweep
        iteration, so offsets are always recomputed from the current thickness
        - never frozen from a one-time base-model calculation. Returns a list
        of {'name', 'subfolder', 'bdf_path', 'run_ok'} dicts, one per bdf_model."""
        n_bdfs = len(self.bdf_models) if self.bdf_models else 1

        offset_csv = self._calculate_offset_csv(thickness_overrides, iter_folder, label=label)

        per_bdf = []
        for bdf_idx, bdf_info in enumerate(self.bdf_models):
            bdf_name = bdf_info['name']
            bdf_subfolder = os.path.join(iter_folder, f"bdf_{bdf_idx + 1}_{os.path.splitext(bdf_name)[0]}") if n_bdfs > 1 else iter_folder
            os.makedirs(bdf_subfolder, exist_ok=True)
            base_source = bdf_info.get('base_path', bdf_info['path'])
            bdf_path = self._write_bdf_for_model(bdf_subfolder, bdf_info['model'], base_source, thickness_overrides)
            if offset_csv:
                bdf_path = self._apply_offsets_from_csv(bdf_path, offset_csv, bdf_subfolder, label=label)
            self._write_structure_only_bdf(bdf_path, bdf_subfolder, label=label)
            per_bdf.append({'name': bdf_name, 'subfolder': bdf_subfolder, 'bdf_path': bdf_path, 'run_ok': False})
        return per_bdf

    # ==================== STRUCTURE-ONLY BDF (NO INCLUDES) ====================
    def _write_structure_only_bdf(self, bdf_path, folder, label=""):
        """Write a copy of the given (prop-updated + offseted) BDF with INCLUDE
        cards removed, so only the structure itself remains, without needing
        the include files. Reference-only output - never solved."""
        try:
            with open(bdf_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()

            new_lines = [l for l in lines if not l.strip().upper().startswith('INCLUDE')]
            removed = len(lines) - len(new_lines)

            base, ext = os.path.splitext(bdf_path)
            output_bdf = base + "_stru" + ext
            with open(output_bdf, 'w', encoding='latin-1') as f:
                f.writelines(new_lines)

            self.log(f"    {label}Structure-only BDF written: {os.path.basename(output_bdf)} ({removed} include lines removed)")
            return output_bdf

        except Exception as e:
            self.log(f"    {label}Structure-only BDF error: {e}")
            return None

    # ==================== NASTRAN EXECUTION ====================
    def _run_nastran(self, bdf_path, folder, label=""):
        nastran = self.nastran_path.get()
        if not nastran or not os.path.exists(nastran):
            self.log(f"    {label}WARNING: Nastran exe not found!")
            return False

        # Neither of the two earlier approaches to scratch isolation actually
        # redirected where Nastran puts its scratch files. Confirmed on real
        # hardware: with a Scratch Folder configured, a run FATAL-errored with
        # "There is not enough space on the disk", and the offending scratch
        # file was sitting in C:\Users\<user>\AppData\Local\Temp\<n>\ - the
        # plain Windows per-process temp folder, completely ignoring both the
        # sdirectory= argument and cwd= that were tried before. That path
        # shape is exactly what Windows' GetTempPath() returns, which reads
        # the TMP/TEMP environment variables - so Nastran's scratch placement
        # here follows those variables, not argv or the process cwd. Override
        # them for this subprocess only (not the whole session) so its
        # scratch lands in the chosen Scratch Folder instead of the system
        # C: temp drive, without touching any other process on the machine.
        scratch_base = self.scratch_folder.get().strip()
        run_scratch_dir = None
        made_temp_scratch = False
        try:
            env = None
            if scratch_base:
                os.makedirs(scratch_base, exist_ok=True)
                run_scratch_dir = tempfile.mkdtemp(prefix="nastran_scr_", dir=scratch_base)
                made_temp_scratch = True
                env = os.environ.copy()
                env['TEMP'] = run_scratch_dir
                env['TMP'] = run_scratch_dir
                self.log(f"    {label}Scratch dir (via TEMP/TMP): {run_scratch_dir}")

            cwd_dir = run_scratch_dir or folder

            cmd = [nastran, bdf_path, f"out={folder}", "scratch=yes", "batch=no"]
            mem = self.nastran_memory.get().strip()
            if mem:
                cmd.append(f"mem={mem}")

            result = subprocess.run(cmd, cwd=cwd_dir, env=env, capture_output=True, text=True)
            if result.returncode != 0:
                self.log(f"    {label}WARNING: Nastran exited with code {result.returncode}")
                if result.stderr:
                    self.log(f"    {label}stderr: {result.stderr.strip()[:500]}")

            if made_temp_scratch:
                scratch_files = os.listdir(run_scratch_dir)
                if scratch_files:
                    self.log(f"    {label}Scratch folder in use: {len(scratch_files)} file(s) in {run_scratch_dir}")
                else:
                    self.log(f"    {label}WARNING: Scratch folder is still empty - Nastran likely used its "
                              f"own default temp location instead of {run_scratch_dir}")

            # subprocess.run() (no shell=True) blocks until Nastran itself has
            # actually exited, unlike the earlier shell=True + Popen().wait()
            # combination. Still confirm the .op2 as a safety net in case some
            # install's launcher detaches a child process anyway.
            if not self._wait_for_op2(bdf_path, folder, scratch_dir=run_scratch_dir, label=label):
                self.log(f"    {label}WARNING: .op2 never appeared/finished writing - treating run as failed")
                return False
            return True
        except Exception as e:
            self.log(f"    {label}Nastran run error: {e}")
            return False
        finally:
            # Scratch files are pure intermediates, never needed once the run is
            # done - clean up our own temp scratch dir so a long sweep doesn't
            # fill the scratch disk.
            if made_temp_scratch and run_scratch_dir and os.path.isdir(run_scratch_dir):
                shutil.rmtree(run_scratch_dir, ignore_errors=True)

    def _describe_run_progress(self, dirs, base):
        """Diagnostic string for the heartbeat below: which of the usual Nastran
        output files exist yet in the watched folder(s), and their size."""
        parts = []
        seen = set()
        for d in dirs:
            if not d or d in seen or not os.path.isdir(d):
                continue
            seen.add(d)
            for ext in ('.f04', '.f06', '.log', '.op2'):
                p = os.path.join(d, base + ext)
                if os.path.exists(p):
                    try:
                        parts.append(f"{os.path.basename(p)}={os.path.getsize(p)}B")
                    except OSError:
                        pass
        return ", ".join(parts) if parts else "no output files found yet in any watched folder"

    def _wait_for_op2(self, bdf_path, folder, scratch_dir=None, timeout=3600, poll_interval=5,
                       heartbeat_every=30, label=""):
        """Poll until Nastran has genuinely finished this run.

        This used to declare "done" once the .op2 file's size stopped
        changing for 3 polls (15s). On a large model that's wrong: Nastran
        can go minutes with no bytes written to the .op2 during assembly/
        decomposition, long before the run is actually done - so that
        heuristic returned True while the real solve (and the detached
        "analysis" process behind it) kept running in the background. The
        tool then read a still-incomplete .op2 (empty/garbage results) AND
        freed its worker slot early, letting more jobs start than Max
        Parallel Runs was supposed to allow.

        The reliable signal is Nastran's own .f06 log: it contains "FATAL"
        on failure and "END OF JOB" once a run has genuinely completed,
        written only when the solve itself is done - not when some
        launcher process happens to exit. Only if no .f06 ever appears at
        all (a different install's behavior) does this fall back to a much
        longer op2-size-stability window (minutes, not seconds) as a last
        resort.
        """
        base = os.path.splitext(os.path.basename(bdf_path))[0]
        f06_path = os.path.join(folder, base + '.f06')
        op2_search_dirs = [folder] + ([scratch_dir] if scratch_dir and scratch_dir != folder else [])
        start = time.time()
        last_heartbeat = 0.0
        last_size = -1
        stable_since = None
        fallback_stable_seconds = 180

        def _find_op2():
            for d in op2_search_dirs:
                candidate = os.path.join(d, base + '.op2')
                if os.path.exists(candidate) and os.path.getsize(candidate) > 0:
                    return d, candidate
            return None, None

        while time.time() - start < timeout:
            elapsed = time.time() - start

            if os.path.exists(f06_path):
                try:
                    with open(f06_path, 'r', errors='ignore') as f:
                        tail = f.read()[-8000:].upper()
                except OSError:
                    tail = ''

                if 'FATAL' in tail:
                    self.log(f"    {label}Nastran reported a FATAL error - see {f06_path}")
                    return False

                if 'END OF JOB' in tail:
                    op2_dir, op2_path = _find_op2()
                    if op2_path:
                        if op2_dir != folder:
                            self.log(f"    {label}NOTE: .op2 appeared in {op2_dir} (not the out= folder) - copying it to {folder}")
                            try:
                                shutil.copy2(op2_path, os.path.join(folder, base + '.op2'))
                            except Exception as e:
                                self.log(f"    {label}WARNING: could not copy .op2 into {folder}: {e}")
                        return True
                    # F06 says the job is over but the .op2 hasn't shown up
                    # yet - keep polling a little longer rather than failing.
            else:
                # No .f06 at all for this install/config - fall back to a
                # long op2 size-stability window instead of failing outright.
                _, op2_path = _find_op2()
                if op2_path:
                    size = os.path.getsize(op2_path)
                    if size == last_size:
                        if stable_since is None:
                            stable_since = time.time()
                        elif time.time() - stable_since >= fallback_stable_seconds:
                            return True
                    else:
                        stable_since = None
                    last_size = size

            if elapsed - last_heartbeat >= heartbeat_every:
                last_heartbeat = elapsed
                self.log(f"    {label}...still waiting for Nastran ({int(elapsed)}s elapsed): "
                          f"{self._describe_run_progress([folder] + op2_search_dirs, base)}")

            time.sleep(poll_interval)

        self.log(f"    {label}WARNING: timed out after {timeout}s waiting for {base} to finish")
        _, op2_path = _find_op2()
        return op2_path is not None

    # ==================== STRESS EXTRACTION ====================
    def _extract_stresses(self, folder, thickness_overrides):
        """thickness_overrides: {pid: dim1} for the group actively swept in this
        iteration (empty for the base model run). PIDs not in it use their base
        self.bar_properties dim1 - same explicit-argument approach as
        _write_bdf_for_model, so concurrent iterations don't race on shared state."""
        results = []
        bar_stress_rows = []

        for f in os.listdir(folder):
            if f.lower().endswith('.op2'):
                op2_name = f
                try:
                    op2 = OP2(debug=False)
                    op2.read_op2(os.path.join(folder, f))

                    # BAR STRESS from cbar_force
                    if hasattr(op2, 'cbar_force') and op2.cbar_force:
                        for sc_id, force in op2.cbar_force.items():
                            for i, eid in enumerate(force.element):
                                axial = force.data[0, i, 6] if len(force.data.shape) == 3 else force.data[i, 6]
                                pid = self.elem_to_prop.get(int(eid))
                                d1 = d2 = area = stress = None

                                if pid and pid in self.bar_properties:
                                    d1 = thickness_overrides.get(pid, self.bar_properties[pid].get('dim1', 0))
                                    # dim2 must come from self.bar_properties (the Excel 'Bar
                                    # Property' sheet, with BDF as its own fallback - see
                                    # load_properties) - NOT self.pbarl_dims directly, which is
                                    # always the raw original BDF value and previously overrode
                                    # the Excel dim2 here even when Excel provided one, making
                                    # the reported stress use the wrong area.
                                    d2 = self.bar_properties[pid].get('dim2', d1)
                                    area = d1 * d2
                                    if area > 0:
                                        stress = axial / area

                                results.append({
                                    'eid': int(eid), 'type': 'bar',
                                    'stress': float(stress) if stress else 0,
                                    'subcase': int(sc_id)
                                })

                                struct_name = self.bar_structure_map.get(pid, '') if pid else ''
                                bar_stress_rows.append({
                                    'OP2': op2_name, 'Subcase': int(sc_id), 'Element': int(eid),
                                    'Property': pid, 'Structure': struct_name,
                                    'Axial': float(axial) if axial else 0,
                                    'Dim1': d1, 'Dim2': d2, 'Area': area,
                                    'Stress': float(stress) if stress else None
                                })

                    # SHELL STRESS
                    shell_stress_attrs = [
                        ('cquad4_stress', 'CQUAD4'),
                        ('ctria3_stress', 'CTRIA3'),
                        ('cquad8_stress', 'CQUAD8'),
                        ('ctria6_stress', 'CTRIA6'),
                        ('cquad4_composite_stress', 'CQUAD4_COMP'),
                        ('ctria3_composite_stress', 'CTRIA3_COMP'),
                    ]

                    for attr_name, stress_type in shell_stress_attrs:
                        if hasattr(op2, attr_name):
                            stress_data = getattr(op2, attr_name)
                            if stress_data:
                                for sc_id, data in stress_data.items():
                                    for i, eid in enumerate(data.element):
                                        try:
                                            if len(data.data.shape) == 3:
                                                stress = data.data[0, i, -1]
                                            else:
                                                stress = data.data[i, -1]
                                            results.append({
                                                'eid': int(eid), 'type': 'shell',
                                                'stress': float(abs(stress)),
                                                'subcase': int(sc_id)
                                            })
                                        except:
                                            pass

                except Exception as e:
                    self.log(f"    OP2 read error: {e}")

        # Save bar stress CSV
        if bar_stress_rows:
            csv_path = os.path.join(folder, 'bar_stress_results.csv')
            with open(csv_path, 'w', newline='') as f:
                w = csv.DictWriter(f, fieldnames=[
                    'OP2', 'Subcase', 'Element', 'Property', 'Structure', 'Axial', 'Dim1', 'Dim2', 'Area', 'Stress'
                ])
                w.writeheader()
                w.writerows(bar_stress_rows)
            self.log(f"    Saved: bar_stress_results.csv ({len(bar_stress_rows)} rows)")

        return results

    # ==================== SOLVE BASE MODEL ====================
    def _run_single_iteration(self, folder, thickness_overrides, label=""):
        """Write every BDF + fresh offsets for thickness_overrides, run Nastran,
        and extract stresses. thickness_overrides is {} for the base model's
        own solve. Offsets are recomputed here every call (see
        _prepare_iteration_bdfs / _calculate_offset_csv) from the maneuver BDF
        at this same thickness_overrides."""
        try:
            per_bdf = self._prepare_iteration_bdfs(folder, thickness_overrides, label=label)
            n_bdfs = len(per_bdf)
            all_stresses = []

            for pb in per_bdf:
                if n_bdfs > 1:
                    self.log(f"    {label}--- BDF: {pb['name']} ---")

                # Run Nastran and WAIT for it to actually finish (incl. the
                # .op2 being fully written)
                self.log(f"    {label}Running Nastran...")
                ok = self._run_nastran(pb['bdf_path'], pb['subfolder'], label=label)
                if not ok:
                    self.log(f"    {label}Nastran run failed/incomplete for {pb['name']} - skipping stress extraction for it")
                    continue

                # Extract stresses (only once the run above is confirmed done)
                stresses = self._extract_stresses(pb['subfolder'], thickness_overrides)
                all_stresses.extend(stresses)

                if stresses:
                    bar_s = [s for s in stresses if s['type'] == 'bar']
                    shell_s = [s for s in stresses if s['type'] == 'shell']
                    self.log(f"    {label}Extracted: {len(bar_s)} bar, {len(shell_s)} shell stresses")

            return all_stresses

        except Exception as e:
            self.log(f"  {label}Iteration ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            return []

    def start_solve(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one BDF file")
            return
        if not self.bdf_model:
            self.load_bdf()
        if not self.bdf_model:
            return
        if not self.structure_groups:
            messagebox.showerror("Error", "Load Property Excel first")
            return

        self.is_running = True
        self.btn_start.config(state=tk.DISABLED)

        threading.Thread(target=self._run_solve, daemon=True).start()

    def _run_solve(self):
        """Process the property excel, apply offsets, and solve the base model once."""
        try:
            output_base = self.output_folder.get()
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            run_folder = os.path.join(output_base, f"BaseModel_{ts}")
            os.makedirs(run_folder, exist_ok=True)

            self.log("\n" + "=" * 70)
            self.log("SOLVING BASE MODEL")
            self.log("=" * 70)
            self.log(f"  Structure groups: {len(self.structure_groups)}")
            self.log(f"  Output: {run_folder}")

            self.root.after(0, lambda: self.progress_var.set(10))
            self._build_base_model(run_folder)
            self.root.after(0, lambda: self.progress_var.set(100))

            self.log(f"\n{'=' * 70}")
            self.log("BASE MODEL SOLVE COMPLETE")
            self.log(f"  Output folder: {run_folder}")
            self.log(f"{'=' * 70}")

        except Exception as e:
            self.log(f"\nSOLVE ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))




def main():
    root = tk.Tk()
    app = IntegratedBDFRFTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
