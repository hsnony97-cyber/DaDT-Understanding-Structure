#!/usr/bin/env python3
"""
Integrated BDF Tool v16.0
=========================
Tab 1: BDF Merge Preparation
Tab 2: BDF Post-Process (with integrated offset calculation & application)
Tab 3: Understanding Structure Type (maneuver→thermal offset, 'Bar Property Structure Type' sheet)
Tab 4: Structure Optimization (thickness iteration with maneuver→thermal offset)
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
import random
import copy
from datetime import datetime
from itertools import combinations
import pandas as pd
from pyNastran.bdf.bdf import BDF
from pyNastran.op2.op2 import OP2
import numpy as np
from scipy.optimize import minimize, curve_fit


class IntegratedBDFRFTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Integrated BDF Tool v16.0")
        self.root.geometry("1100x950")
        
        # Tab 1 variables
        self.thermal_bdfs = []
        self.maneuver_bdfs = []
        self.excel_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_thermal_name = tk.StringVar(value="merged_thermal.bdf")
        self.output_maneuver_name = tk.StringVar(value="merged_maneuver.bdf")
        self.set_id = tk.StringVar(value="99")
        self.temp_initial = tk.StringVar(value="10")
        
        # Tab 2 variables
        self.run_bdfs = []
        self.property_excel_path = tk.StringVar()
        self.nastran_path = tk.StringVar()
        self.run_output_folder = tk.StringVar()
        self.csv_output_name = tk.StringVar(value="bar_stress_results.csv")
        self.combined_csv_name = tk.StringVar(value="combined_stress_results.csv")
        
        self.bar_properties = {}
        self.skin_properties = {}
        self.residual_strength_df = None
        
        # Offset variables
        self.offset_element_excel = tk.StringVar()

        self.setup_ui()
    
    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="BDF Merge Preparation")
        self.setup_tab1()
        
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="BDF Post-Process")
        self.setup_tab2()

        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="Understanding Structure Type")
        self.tab3_tool = BarPropertySolverTab(self.tab3, self.root)

        self.tab4 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab4, text="Structure Optimization")
        self.tab4_tool = StructureOptimizationTab(self.tab4, self.root)
    
    def setup_tab1(self):
        main = ttk.Frame(self.tab1, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main, text="BDF Merge Preparation v8", font=('Helvetica', 14, 'bold')).pack(pady=(0,10))
        
        # === THERMAL SECTION ===
        thermal_main = ttk.LabelFrame(main, text="THERMAL", padding="5")
        thermal_main.pack(fill=tk.X, pady=5)
        
        thm_master_f = ttk.Frame(thermal_main)
        thm_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(thm_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(thm_master_f, text="Add...", command=self.add_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(thm_master_f, text="Clear", command=self.clear_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        self.thermal_count = tk.StringVar(value="0 files")
        ttk.Label(thm_master_f, textvariable=self.thermal_count).pack(side=tk.LEFT, padx=5)
        
        self.thermal_listbox = tk.Listbox(thermal_main, height=3, width=100)
        self.thermal_listbox.pack(fill=tk.X, pady=2)
        
        # === MANEUVER SECTION ===
        maneuver_main = ttk.LabelFrame(main, text="MANEUVER", padding="5")
        maneuver_main.pack(fill=tk.X, pady=5)
        
        man_master_f = ttk.Frame(maneuver_main)
        man_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(man_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(man_master_f, text="Add...", command=self.add_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(man_master_f, text="Clear", command=self.clear_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        self.maneuver_count = tk.StringVar(value="0 files")
        ttk.Label(man_master_f, textvariable=self.maneuver_count).pack(side=tk.LEFT, padx=5)
        
        self.maneuver_listbox = tk.Listbox(maneuver_main, height=3, width=100)
        self.maneuver_listbox.pack(fill=tk.X, pady=2)
        
        # === SETTINGS ===
        sf = ttk.LabelFrame(main, text="Settings", padding="10")
        sf.pack(fill=tk.X, pady=5)
        ttk.Label(sf, text="Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.excel_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        ttk.Label(sf, text="Output:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.output_folder, width=70).grid(row=1, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_output).grid(row=1, column=2)
        ttk.Label(sf, text="SET ID:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.set_id, width=10).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Label(sf, text="TEMP(INIT):").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.temp_initial, width=10).grid(row=3, column=1, sticky=tk.W, padx=5)
        
        bf = ttk.Frame(main)
        bf.pack(fill=tk.X, pady=10)
        self.process_btn = ttk.Button(bf, text=">>> PROCESS & MERGE <<<", command=self.start_processing)
        self.process_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Clear Log", command=self.clear_log1).pack(side=tk.LEFT)
        
        self.progress1 = ttk.Progressbar(main, mode='indeterminate')
        self.progress1.pack(fill=tk.X, pady=5)
        
        lf = ttk.LabelFrame(main, text="Log", padding="10")
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text1 = scrolledtext.ScrolledText(lf, height=15)
        self.log_text1.pack(fill=tk.BOTH, expand=True)
    
    def setup_tab2(self):
        main = ttk.Frame(self.tab2, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main, text="BDF Post-Process", font=('Helvetica', 14, 'bold')).pack(pady=(0,10))
        
        bf = ttk.LabelFrame(main, text="BDF Files", padding="10")
        bf.pack(fill=tk.X, pady=5)
        bb = ttk.Frame(bf)
        bb.pack(fill=tk.X)
        ttk.Button(bb, text="Add...", command=self.add_run_bdfs).pack(side=tk.LEFT, padx=5)
        ttk.Button(bb, text="Clear", command=self.clear_run_bdfs).pack(side=tk.LEFT)
        self.run_listbox = tk.Listbox(bf, height=3, width=100)
        self.run_listbox.pack(fill=tk.X, pady=5)
        self.run_count = tk.StringVar(value="0 files")
        ttk.Label(bf, textvariable=self.run_count).pack(anchor=tk.W)
        
        pf = ttk.LabelFrame(main, text="Property Excel", padding="10")
        pf.pack(fill=tk.X, pady=5)
        ttk.Label(pf, text="Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(pf, textvariable=self.property_excel_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(pf, text="Browse", command=self.browse_property_excel).grid(row=0, column=2)
        ttk.Button(pf, text="Load Properties", command=self.load_properties).grid(row=1, column=0, pady=5)
        
        pvf = ttk.Frame(pf)
        pvf.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=5)
        self.bar_prop_text = tk.Text(pvf, height=2, width=35)
        self.bar_prop_text.pack(side=tk.LEFT, padx=3)
        self.bar_prop_text.insert(tk.END, "Bar: Not loaded")
        self.skin_prop_text = tk.Text(pvf, height=2, width=35)
        self.skin_prop_text.pack(side=tk.LEFT, padx=3)
        self.skin_prop_text.insert(tk.END, "Skin: Not loaded")
        self.resid_text = tk.Text(pvf, height=2, width=30)
        self.resid_text.pack(side=tk.LEFT, padx=3)
        self.resid_text.insert(tk.END, "Residual: Not loaded")
        
        nf = ttk.LabelFrame(main, text="Nastran", padding="10")
        nf.pack(fill=tk.X, pady=5)
        ttk.Label(nf, text="Path:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(nf, textvariable=self.nastran_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(nf, text="Browse", command=self.browse_nastran).grid(row=0, column=2)
        
        of = ttk.LabelFrame(main, text="Output", padding="10")
        of.pack(fill=tk.X, pady=5)
        ttk.Label(of, text="Folder:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(of, textvariable=self.run_output_folder, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(of, text="Browse", command=self.browse_run_output).grid(row=0, column=2)
        ttk.Label(of, text="CSV:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(of, textvariable=self.csv_output_name, width=25).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # === OFFSET SETTINGS ===
        off = ttk.LabelFrame(main, text="Offset Element IDs (Optional)", padding="10")
        off.pack(fill=tk.X, pady=5)

        off_r1 = ttk.Frame(off)
        off_r1.pack(fill=tk.X, pady=2)
        ttk.Label(off_r1, text="Element Excel:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(off_r1, textvariable=self.offset_element_excel, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(off_r1, text="Browse", command=self.browse_offset_element_excel).pack(side=tk.LEFT, padx=5)
        ttk.Label(off_r1, text="Sheets: 'Landing_Offset', 'Bar_Offset'",
                 font=('Helvetica', 8, 'italic')).pack(side=tk.LEFT, padx=5)

        # === 3 STEP BUTTONS + FULL ===
        af = ttk.Frame(main)
        af.pack(fill=tk.X, pady=10)
        self.btn1 = ttk.Button(af, text="1.Update+Offset", command=self.start_update_and_offset, width=14)
        self.btn1.pack(side=tk.LEFT, padx=2)
        self.btn2 = ttk.Button(af, text="2.Run Nastran", command=self.start_run_nastran, width=14)
        self.btn2.pack(side=tk.LEFT, padx=2)
        self.btn3 = ttk.Button(af, text="3.Post+Combine", command=self.start_postprocess_and_combine, width=14)
        self.btn3.pack(side=tk.LEFT, padx=2)
        self.btn_full = ttk.Button(af, text=">>> FULL <<<", command=self.start_full_run, width=12)
        self.btn_full.pack(side=tk.LEFT, padx=2)
        ttk.Button(af, text="Clear", command=self.clear_log2).pack(side=tk.LEFT, padx=2)
        
        self.progress2 = ttk.Progressbar(main, mode='indeterminate')
        self.progress2.pack(fill=tk.X, pady=5)
        
        lf = ttk.LabelFrame(main, text="Log", padding="10")
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text2 = scrolledtext.ScrolledText(lf, height=12)
        self.log_text2.pack(fill=tk.BOTH, expand=True)


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

            thermal_sh = maneuver_sh = element_sh = None

            # First pass: look for exact or specific matches
            for s in sheets:
                sl = s.lower()
                # Element_Set exact match (priority)
                if sl == 'element_set' or sl == 'elementset':
                    element_sh = s
                # Thermal
                elif 'thermal' in sl and not thermal_sh:
                    thermal_sh = s
                # Maneuver
                elif ('maneuver' in sl or 'manevra' in sl) and not maneuver_sh:
                    maneuver_sh = s

            # Second pass: if element_sh not found, look for partial matches
            if not element_sh:
                for s in sheets:
                    sl = s.lower()
                    if 'element' in sl and 'set' in sl:
                        element_sh = s
                        break

            # Fallback to index-based if still not found
            if not thermal_sh and len(sheets) > 0: thermal_sh = sheets[0]
            if not maneuver_sh and len(sheets) > 1: maneuver_sh = sheets[1]
            if not element_sh and len(sheets) > 2: element_sh = sheets[2]

            self.log1(f"    Using sheets -> Thermal: '{thermal_sh}', Maneuver: '{maneuver_sh}', Element_Set: '{element_sh}'")
            
            thermal_cases = pd.read_excel(xl, sheet_name=thermal_sh).iloc[:,0].dropna().astype(int).tolist() if thermal_sh else []
            maneuver_cases = pd.read_excel(xl, sheet_name=maneuver_sh).iloc[:,0].dropna().astype(int).tolist() if maneuver_sh else []
            element_ids = sorted(pd.read_excel(xl, sheet_name=element_sh).iloc[:,0].dropna().astype(int).tolist()) if element_sh else []
            
            self.log1(f"    Thermal cases: {len(thermal_cases)}")
            self.log1(f"    Maneuver cases: {len(maneuver_cases)}")
            self.log1(f"    Element IDs: {len(element_ids)}")
            
            set_id = int(self.set_id.get())
            temp_initial = self.temp_initial.get()
            out_dir = self.output_folder.get()
            os.makedirs(out_dir, exist_ok=True)
            
            if self.thermal_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[2] Processing THERMAL...")
                self.log1("="*70)
                self.process_thermal_bdf(self.thermal_bdfs, thermal_cases, element_ids, set_id,
                    temp_initial, os.path.join(out_dir, self.output_thermal_name.get()))
            
            if self.maneuver_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[3] Processing MANEUVER...")
                self.log1("="*70)
                self.process_maneuver_bdf(self.maneuver_bdfs, maneuver_cases, element_ids, set_id,
                    os.path.join(out_dir, self.output_maneuver_name.get()))
            
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
    
    def process_thermal_bdf(self, bdf_files, load_cases, element_ids, set_id, temp_initial, output_path):
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
        
        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append("GPFORCE(PLOT)=ALL")
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
    
    def process_maneuver_bdf(self, bdf_files, load_cases, element_ids, set_id, output_path):
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
        
        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append("GPFORCE(PLOT)=ALL")
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

    # ============= TAB 2 HELPERS =============
    def add_run_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.run_bdfs:
                self.run_bdfs.append(f)
                self.run_listbox.insert(tk.END, f)
        self.run_count.set(f"{len(self.run_bdfs)} files")
    
    def clear_run_bdfs(self):
        self.run_bdfs.clear()
        self.run_listbox.delete(0, tk.END)
        self.run_count.set("0 files")
    
    def browse_property_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f: self.property_excel_path.set(f)
    
    def browse_nastran(self):
        f = filedialog.askopenfilename(filetypes=[("All","*.*")])
        if f: self.nastran_path.set(f)
    
    def browse_run_output(self):
        f = filedialog.askdirectory()
        if f: self.run_output_folder.set(f)
    
    def log2(self, msg):
        self.log_text2.insert(tk.END, msg + "\n")
        self.log_text2.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log2(self):
        self.log_text2.delete(1.0, tk.END)
    
    def load_properties(self):
        if not self.property_excel_path.get():
            messagebox.showerror("Error","Select Excel"); return
        try:
            xl = pd.ExcelFile(self.property_excel_path.get())
            sheets = xl.sheet_names
            bar_sh = skin_sh = res_sh = None

            # First pass: exact matches (priority)
            for s in sheets:
                sl = s.lower().replace('_',' ').replace('-',' ')
                if sl == 'skin property' or sl == 'skinproperty':
                    skin_sh = s
                elif sl == 'bar property' or sl == 'barproperty':
                    bar_sh = s
                elif sl == 'residual strength' or sl == 'residualstrength':
                    res_sh = s

            # Second pass: partial matches
            for s in sheets:
                sl = s.lower().replace('_',' ')
                if not bar_sh and 'bar' in sl and 'prop' in sl:
                    bar_sh = s
                elif not skin_sh and 'skin' in sl and 'prop' in sl:
                    skin_sh = s
                elif not res_sh and ('residual' in sl or 'strength' in sl):
                    res_sh = s

            self.bar_properties.clear()
            self.skin_properties.clear()

            if bar_sh:
                df = pd.read_excel(xl, sheet_name=bar_sh)
                for _, row in df.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        d1 = float(row.iloc[1]) if len(df.columns)>1 else 0
                        d2 = float(row.iloc[2]) if len(df.columns)>2 else 0
                        self.bar_properties[pid] = {'dim1':d1, 'dim2':d2}
                    except: pass

            if skin_sh:
                df = pd.read_excel(xl, sheet_name=skin_sh)
                for _, row in df.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        t = float(row.iloc[1])
                        self.skin_properties[pid] = {'thickness':t}
                    except: pass

            if res_sh:
                self.residual_strength_df = pd.read_excel(xl, sheet_name=res_sh)

            self.bar_prop_text.delete(1.0, tk.END)
            self.bar_prop_text.insert(tk.END, f"Bar: {len(self.bar_properties)} loaded")
            self.skin_prop_text.delete(1.0, tk.END)
            self.skin_prop_text.insert(tk.END, f"Skin: {len(self.skin_properties)} loaded")
            self.resid_text.delete(1.0, tk.END)
            self.resid_text.insert(tk.END, f"Residual: {'Yes' if self.residual_strength_df is not None else 'No'}")

            # Log which sheets were used
            print(f"[Load Properties] Bar sheet: {bar_sh}, Skin sheet: {skin_sh}")
            print(f"[Load Properties] Bar PIDs: {len(self.bar_properties)}, Skin PIDs: {len(self.skin_properties)}")
            if self.skin_properties:
                sample_pids = list(self.skin_properties.keys())[:5]
                print(f"[Load Properties] Sample Skin PIDs: {sample_pids}")

            messagebox.showinfo("OK", f"Bar:{len(self.bar_properties)} Skin:{len(self.skin_properties)}\n\nSheets used:\nBar: {bar_sh}\nSkin: {skin_sh}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", str(e))
    
    def read_file(self, path):
        for enc in ['utf-8','latin-1','cp1252']:
            try:
                with open(path, 'r', encoding=enc) as f:
                    return f.read()
            except: pass
        return ""
    
    def copy_bdf_to_output(self, bdf_path, output_folder):
        bdf_name = os.path.basename(bdf_path)
        out_bdf = os.path.join(output_folder, bdf_name)
        shutil.copy2(bdf_path, out_bdf)
        return out_bdf
    
    def count_pcomp_plies(self, lines, start_idx):
        ply_count = 0
        j = start_idx + 1
        while j < len(lines):
            line = lines[j]
            stripped = line.strip()
            is_continuation = (
                line.startswith('+') or line.startswith('*') or
                (line.startswith(' ') and stripped and not stripped.startswith('$') and
                 not any(stripped.upper().startswith(card) for card in 
                        ['PSHELL','PCOMP','PBAR','PBARL','CBAR','CQUAD','CTRIA',
                         'GRID','MAT','FORCE','INCLUDE','END','SOL','CEND','BEGIN']))
            )
            if not is_continuation:
                break
            ply_count += 1
            j += 1
        return ply_count, j
    
    def update_properties_in_file(self, filepath):
        content = self.read_file(filepath)
        lines = content.split('\n')
        new_lines = []
        i = 0
        stats = {'pbarl': 0, 'pbar': 0, 'pshell': 0, 'pcomp': 0}
        warnings = []
        pshell_found = []
        pcomp_found = []
        
        while i < len(lines):
            line = lines[i]
            upper = line.upper().strip()
            
            if upper.startswith('PBARL'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    if i+1 < len(lines) and pid in self.bar_properties:
                        d1 = self.bar_properties[pid]['dim1']
                        d2 = self.bar_properties[pid]['dim2']
                        new_lines.append(line)
                        next_line = lines[i+1]
                        if ',' in next_line:
                            parts = next_line.split(',')
                            start_idx = 1 if parts[0].strip().startswith('+') else 0
                            if len(parts) > start_idx: parts[start_idx] = f"{d1}."
                            if len(parts) > start_idx + 1: parts[start_idx + 1] = f"{d2}."
                            new_lines.append(','.join(parts))
                        else:
                            cont = next_line[:8]
                            rest = next_line[24:] if len(next_line) > 24 else ""
                            d1_str = f"{d1:<8.6g}".rstrip()
                            if '.' not in d1_str and 'E' not in d1_str.upper(): d1_str += '.'
                            d2_str = f"{d2:<8.6g}".rstrip()
                            if '.' not in d2_str and 'E' not in d2_str.upper(): d2_str += '.'
                            new_lines.append(f"{cont}{d1_str:>8}{d2_str:>8}{rest}")
                        stats['pbarl'] += 1
                        i += 2
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            
            elif upper.startswith('PBAR') and not upper.startswith('PBARL'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    if pid in self.bar_properties:
                        d1 = self.bar_properties[pid]['dim1']
                        d2 = self.bar_properties[pid]['dim2']
                        area = d1 * d2
                        if ',' in line:
                            parts = line.split(',')
                            parts[3] = str(area)
                            new_lines.append(','.join(parts))
                        else:
                            new_lines.append(line[:24] + f"{area:8.4g}" + line[32:])
                        stats['pbar'] += 1
                        i += 1
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            
            elif upper.startswith('PSHELL'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    pshell_found.append(pid)
                    if pid in self.skin_properties:
                        t = self.skin_properties[pid]['thickness']
                        if ',' in line:
                            parts = line.split(',')
                            t_str = f"{t}"
                            if '.' not in t_str and 'E' not in t_str.upper():
                                t_str += '.'
                            parts[3] = t_str
                            new_lines.append(','.join(parts))
                        else:
                            t_str = f"{t:<8.6g}".rstrip()
                            if '.' not in t_str and 'E' not in t_str.upper():
                                t_str += '.'
                            new_lines.append(line[:24] + f"{t_str:>8}" + line[32:])
                        stats['pshell'] += 1
                        i += 1
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            
            elif upper.startswith('PCOMP'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    pcomp_found.append(pid)
                    ply_count, end_idx = self.count_pcomp_plies(lines, i)
                    if pid in self.skin_properties:
                        if ply_count > 1:
                            warnings.append(f"PCOMP {pid}: {ply_count} plies - SKIPPED")
                            for k in range(i, end_idx):
                                new_lines.append(lines[k])
                            i = end_idx
                            continue
                        else:
                            t = self.skin_properties[pid]['thickness']
                            new_lines.append(line)
                            if i + 1 < end_idx:
                                cont_line = lines[i + 1]
                                if ',' in cont_line:
                                    parts = cont_line.split(',')
                                    if len(parts) >= 3:
                                        t_str = f"{t}"
                                        if '.' not in t_str and 'E' not in t_str.upper():
                                            t_str += '.'
                                        parts[2] = t_str
                                    new_lines.append(','.join(parts))
                                else:
                                    cont = cont_line[:8]
                                    mid = cont_line[8:16] if len(cont_line) > 8 else "        "
                                    rest = cont_line[24:] if len(cont_line) > 24 else ""
                                    t_str = f"{t:<8.6g}".rstrip()
                                    if '.' not in t_str and 'E' not in t_str.upper():
                                        t_str += '.'
                                    new_lines.append(f"{cont}{mid}{t_str:>8}{rest}")
                                for k in range(i + 2, end_idx):
                                    new_lines.append(lines[k])
                            stats['pcomp'] += 1
                            i = end_idx
                            continue
                    else:
                        for k in range(i, end_idx):
                            new_lines.append(lines[k])
                        i = end_idx
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            else:
                new_lines.append(line)
                i += 1
        
        # Debug logging
        print(f"[Update Props] File: {os.path.basename(filepath)}")
        print(f"[Update Props] PSHELL found in BDF: {len(pshell_found)}, PIDs: {pshell_found[:10]}...")
        print(f"[Update Props] PCOMP found in BDF: {len(pcomp_found)}, PIDs: {pcomp_found[:10]}...")
        print(f"[Update Props] Skin properties loaded: {len(self.skin_properties)}")
        if pshell_found:
            matched = [p for p in pshell_found if p in self.skin_properties]
            not_matched = [p for p in pshell_found if p not in self.skin_properties]
            print(f"[Update Props] PSHELL matched: {len(matched)}, not matched: {len(not_matched)}")
            if not_matched:
                print(f"[Update Props] Sample unmatched PSHELL PIDs: {not_matched[:5]}")

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(new_lines))
        return stats, warnings
    
    def start_update_and_offset(self):
        if not self.run_bdfs:
            messagebox.showerror("Error", "Add BDF files"); return
        if not self.run_output_folder.get():
            messagebox.showerror("Error", "Select output folder"); return
        self.btn1.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_update_and_offset, daemon=True).start()

    def do_update_and_offset(self):
        """Step 1: Copy BDFs → Update Properties → Calculate & Apply Offsets (all in memory)"""
        try:
            self.log2("="*60)
            self.log2("STEP 1: Update Properties + Apply Offsets")
            self.log2("="*60)
            out_folder = self.run_output_folder.get()
            os.makedirs(out_folder, exist_ok=True)

            # --- 1a: Copy BDFs and update properties ---
            if self.bar_properties or self.skin_properties:
                self.log2(f"\n  Bar properties: {len(self.bar_properties)}")
                self.log2(f"  Skin properties: {len(self.skin_properties)}")
                all_warnings = []
                for bdf_path in self.run_bdfs:
                    self.log2(f"\n  Processing: {os.path.basename(bdf_path)}")
                    self.log2("    Copying BDF to output folder...")
                    out_bdf = self.copy_bdf_to_output(bdf_path, out_folder)
                    self.log2("    Updating properties...")
                    stats, warnings = self.update_properties_in_file(out_bdf)
                    self.log2(f"    Updated: PBARL={stats['pbarl']} PBAR={stats['pbar']} PSHELL={stats['pshell']} PCOMP={stats['pcomp']}")
                    all_warnings.extend(warnings)
                if all_warnings:
                    self.log2("\n  Warnings:")
                    for w in all_warnings[:10]:
                        self.log2(f"    {w}")
            else:
                self.log2("\n  No properties loaded - copying BDFs without property update")
                for bdf_path in self.run_bdfs:
                    self.copy_bdf_to_output(bdf_path, out_folder)
                    self.log2(f"  Copied: {os.path.basename(bdf_path)}")

            # --- 1b: Calculate & Apply Offsets (if Element Excel provided) ---
            if self.offset_element_excel.get():
                self.log2("\n" + "="*60)
                self.log2("CALCULATING & APPLYING OFFSETS")
                self.log2("="*60)

                # Read element IDs from Excel
                self.log2("\n  Reading element IDs from Excel...")
                xl = pd.ExcelFile(self.offset_element_excel.get())
                sheets = xl.sheet_names

                landing_sheet = bar_sheet = None
                for s in sheets:
                    s_lower = s.lower().replace('_', '').replace(' ', '')
                    if 'landing' in s_lower and 'offset' in s_lower:
                        landing_sheet = s
                    elif 'bar' in s_lower and 'offset' in s_lower:
                        bar_sheet = s

                landing_elem_ids = []
                bar_elem_ids = []

                if landing_sheet:
                    df = pd.read_excel(xl, sheet_name=landing_sheet)
                    landing_elem_ids = df.iloc[:,0].dropna().astype(int).tolist()
                    self.log2(f"  Landing elements: {len(landing_elem_ids)} (from '{landing_sheet}')")

                if bar_sheet:
                    df = pd.read_excel(xl, sheet_name=bar_sheet)
                    bar_elem_ids = df.iloc[:,0].dropna().astype(int).tolist()
                    self.log2(f"  Bar elements: {len(bar_elem_ids)} (from '{bar_sheet}')")

                if not landing_elem_ids and not bar_elem_ids:
                    self.log2("  No element IDs found - skipping offsets")
                else:
                    # Read UPDATED BDF from output folder (not original) for correct thickness values
                    updated_bdf_path = os.path.join(out_folder, os.path.basename(self.run_bdfs[0]))
                    bdf_path = updated_bdf_path
                    self.log2(f"\n  Reading UPDATED BDF with pyNastran: {os.path.basename(bdf_path)}")

                    bdf_model = BDF(debug=False)
                    try:
                        bdf_model.read_bdf(bdf_path, validate=False, xref=False,
                                    read_includes=True, encoding='latin-1')
                    except Exception:
                        bdf_model = BDF(debug=False)
                        bdf_model.read_bdf(bdf_path, validate=False, xref=False,
                                    read_includes=True, encoding='latin-1', punch=True)

                    self.log2(f"  Nodes: {len(bdf_model.nodes)}, Elements: {len(bdf_model.elements)}")

                    # Calculate landing offsets
                    landing_offsets = {}  # {eid: zoffset}
                    landing_thickness = {}
                    landing_normals = {}

                    for eid in landing_elem_ids:
                        if eid in bdf_model.elements:
                            elem = bdf_model.elements[eid]
                            if hasattr(elem, 'pid') and elem.pid in bdf_model.properties:
                                prop = bdf_model.properties[elem.pid]
                                thickness = None
                                if hasattr(prop, 't'):
                                    thickness = prop.t
                                elif hasattr(prop, 'total_thickness'):
                                    thickness = prop.total_thickness()
                                if thickness:
                                    landing_offsets[eid] = -thickness / 2.0
                                    landing_thickness[eid] = thickness

                                    if elem.type in ['CQUAD4', 'CTRIA3', 'CQUAD8', 'CTRIA6']:
                                        node_ids = elem.node_ids[:4] if elem.type.startswith('CQUAD') else elem.node_ids[:3]
                                        nodes = [bdf_model.nodes[nid] for nid in node_ids if nid in bdf_model.nodes]
                                        if len(nodes) >= 3:
                                            p1 = np.array(nodes[0].xyz)
                                            p2 = np.array(nodes[1].xyz)
                                            p3 = np.array(nodes[2].xyz)
                                            normal = np.cross(p2 - p1, p3 - p1)
                                            normal_len = np.linalg.norm(normal)
                                            if normal_len > 1e-10:
                                                landing_normals[eid] = normal / normal_len

                    self.log2(f"  Landing offsets calculated: {len(landing_offsets)}")

                    # Build node-to-shell mapping for bar calculations
                    node_to_shells = {}
                    for eid, elem in bdf_model.elements.items():
                        if elem.type in ['CQUAD4', 'CTRIA3', 'CQUAD8', 'CTRIA6']:
                            for nid in elem.node_ids:
                                if nid not in node_to_shells:
                                    node_to_shells[nid] = []
                                node_to_shells[nid].append(eid)

                    # Calculate bar offsets
                    bar_offsets = {}  # {eid: (ox, oy, oz)}
                    bar_no_landing = 0

                    for eid in bar_elem_ids:
                        if eid in bdf_model.elements:
                            elem = bdf_model.elements[eid]
                            if elem.type == 'CBAR' and hasattr(elem, 'pid') and elem.pid in bdf_model.properties:
                                prop = bdf_model.properties[elem.pid]
                                thickness = None
                                if prop.type == 'PBARL':
                                    if hasattr(prop, 'dim') and len(prop.dim) > 0:
                                        thickness = prop.dim[0]
                                elif prop.type == 'PBAR':
                                    if hasattr(prop, 'A') and prop.A > 0:
                                        thickness = np.sqrt(prop.A)
                                if thickness:
                                    bar_nodes = elem.node_ids[:2]
                                    if bar_nodes[0] in node_to_shells and bar_nodes[1] in node_to_shells:
                                        connected = set(node_to_shells[bar_nodes[0]]).intersection(
                                            set(node_to_shells[bar_nodes[1]]))
                                        max_t = 0
                                        best_normal = None
                                        for shell_eid in connected:
                                            if shell_eid in landing_thickness:
                                                t = landing_thickness[shell_eid]
                                                if t > max_t:
                                                    max_t = t
                                                    if shell_eid in landing_normals:
                                                        best_normal = landing_normals[shell_eid]
                                        if best_normal is not None and max_t > 0:
                                            mag = max_t + thickness / 2.0
                                            vec = -best_normal * mag
                                            bar_offsets[eid] = (vec[0], vec[1], vec[2])
                                        else:
                                            bar_no_landing += 1
                                    else:
                                        bar_no_landing += 1

                    self.log2(f"  Bar offsets calculated: {len(bar_offsets)}")
                    if bar_no_landing > 0:
                        self.log2(f"  Bars skipped (no landing): {bar_no_landing}")

                    # Apply offsets to output BDF files (text-based)
                    def fmt_field(value, width=8):
                        if isinstance(value, float):
                            s = f"{value:.4f}"
                            if len(s) > width:
                                s = f"{value:.2E}"
                            return s[:width].ljust(width)
                        return str(value)[:width].ljust(width)

                    out_bdfs = [os.path.join(out_folder, f) for f in os.listdir(out_folder)
                               if f.lower().endswith(('.bdf', '.dat', '.nas'))]

                    for out_bdf in out_bdfs:
                        self.log2(f"\n  Applying offsets to: {os.path.basename(out_bdf)}")
                        with open(out_bdf, 'r', encoding='latin-1') as f:
                            lines = f.readlines()

                        new_lines = []
                        i = 0
                        landing_mod = 0
                        bar_mod = 0

                        while i < len(lines):
                            line = lines[i]

                            if line.startswith('CQUAD4'):
                                try:
                                    eid = int(line[8:16].strip())
                                    if eid in landing_offsets:
                                        zoff = landing_offsets[eid]
                                        if len(line) >= 64:
                                            new_line = line[:64] + fmt_field(zoff) + (line[72:] if len(line) > 72 else '\n')
                                            new_lines.append(new_line)
                                            landing_mod += 1
                                            i += 1
                                            continue
                                except:
                                    pass
                                new_lines.append(line)
                                i += 1

                            elif line.startswith('CBAR'):
                                try:
                                    eid = int(line[8:16].strip())
                                    if eid in bar_offsets:
                                        vec = bar_offsets[eid]
                                        if i + 1 < len(lines) and (lines[i+1].startswith('+') or lines[i+1].startswith('*') or lines[i+1].startswith(' ')):
                                            cont_line = lines[i+1]
                                            new_cont = cont_line[:24]
                                            new_cont += fmt_field(vec[0]) + fmt_field(vec[1]) + fmt_field(vec[2])
                                            new_cont += fmt_field(vec[0]) + fmt_field(vec[1]) + fmt_field(vec[2])
                                            new_cont += '\n'
                                            new_lines.append(line)
                                            new_lines.append(new_cont)
                                            bar_mod += 1
                                            i += 2
                                            continue
                                        else:
                                            cont_name = '+CB' + str(eid)[-4:]
                                            new_lines.append(line.rstrip() + cont_name + '\n')
                                            new_cont = cont_name.ljust(8) + '        ' + '        '
                                            new_cont += fmt_field(vec[0]) + fmt_field(vec[1]) + fmt_field(vec[2])
                                            new_cont += fmt_field(vec[0]) + fmt_field(vec[1]) + fmt_field(vec[2])
                                            new_cont += '\n'
                                            new_lines.append(new_cont)
                                            bar_mod += 1
                                            i += 1
                                            continue
                                except:
                                    pass
                                new_lines.append(line)
                                i += 1

                            else:
                                new_lines.append(line)
                                i += 1

                        # Write to NEW file with _offseted suffix (keep updated BDF unchanged)
                        base, ext = os.path.splitext(out_bdf)
                        offseted_bdf = base + "_offseted" + ext
                        with open(offseted_bdf, 'w', encoding='latin-1') as f:
                            f.writelines(new_lines)
                        self.log2(f"    Landing (ZOFFS): {landing_mod}, Bar (WA/WB): {bar_mod}")
                        self.log2(f"    Written: {os.path.basename(offseted_bdf)}")
            else:
                self.log2("\n  No Element Excel selected - skipping offsets")

            self.log2("\n" + "="*60)
            self.log2("STEP 1 COMPLETED!")
            self.log2("="*60)
            self.root.after(0, lambda: messagebox.showinfo("Done", "Update + Offset completed!"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn1.config(state=tk.NORMAL)])
    
    def start_run_nastran(self):
        if not self.nastran_path.get():
            messagebox.showerror("Error","Select Nastran path"); return
        if not self.run_output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.btn2.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_run_nastran, daemon=True).start()

    def do_run_nastran(self):
        try:
            self.log2("="*60)
            self.log2("STEP 2: Run Nastran")
            self.log2("="*60)
            out_folder = self.run_output_folder.get()
            nastran = self.nastran_path.get()
            bdfs = [os.path.join(out_folder, f) for f in os.listdir(out_folder) if f.lower().endswith(('.bdf','.dat','.nas'))]
            for bdf in bdfs:
                self.log2(f"\n  Running: {os.path.basename(bdf)}")
                cmd = f'"{nastran}" "{bdf}" out="{out_folder}" scratch=yes batch=no'
                os.popen(cmd)
            self.log2("\nJobs submitted!")
            self.root.after(0, lambda: messagebox.showinfo("Done","Nastran jobs submitted!"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn2.config(state=tk.NORMAL)])
    
    def start_postprocess_and_combine(self):
        if not self.run_output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.btn3.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_postprocess_and_combine, daemon=True).start()

    def do_postprocess_and_combine(self):
        """Run post-process and combine stress in one step"""
        try:
            self.do_postprocess_inner()
            if self.residual_strength_df is not None:
                self.do_combine_stress_inner()
            else:
                self.log2("\n  Combine SKIPPED (No Residual Strength data loaded)")
            self.log2("\n" + "="*60)
            self.log2("POST-PROCESS + COMBINE COMPLETED!")
            self.log2("="*60)
            self.root.after(0, lambda: messagebox.showinfo("Done", "Post-process + Combine completed!"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn3.config(state=tk.NORMAL)])

    def do_postprocess_inner(self):
        """Post-process OP2 files (shared logic for step 5 and full run)"""
        self.log2("="*60)
        self.log2("STEP 5a: Post-Process OP2")
        self.log2("="*60)
        out_folder = self.run_output_folder.get()
        op2_files = [os.path.join(out_folder, f) for f in os.listdir(out_folder) if f.lower().endswith('.op2')]
        if not op2_files:
            self.log2("No OP2 files found!")
            return

        elem_prop = {}
        pbarl_dims = {}

        bdf_files_to_read = list(self.run_bdfs)
        for f in os.listdir(out_folder):
            if f.lower().endswith(('.bdf', '.dat', '.nas')):
                bdf_files_to_read.append(os.path.join(out_folder, f))

        for bdf_path in bdf_files_to_read:
            try:
                self.log2(f"  Reading BDF: {os.path.basename(bdf_path)}")
                bdf = BDF(debug=False)
                bdf.read_bdf(bdf_path, validate=False, xref=False, read_includes=True)

                for eid, el in bdf.elements.items():
                    if hasattr(el, 'pid'):
                        elem_prop[eid] = el.pid

                for pid, prop in bdf.properties.items():
                    prop_type = prop.type
                    if prop_type == 'PBARL':
                        dims = prop.dim if hasattr(prop, 'dim') else []
                        bar_type = prop.bar_type if hasattr(prop, 'bar_type') else 'UNKNOWN'
                        if len(dims) >= 2:
                            pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[1], 'type': bar_type}
                        elif len(dims) == 1:
                            pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[0], 'type': bar_type}
                    elif prop_type == 'PBAR':
                        area = prop.A if hasattr(prop, 'A') else None
                        if area:
                            import math
                            side = math.sqrt(area) if area > 0 else 0
                            pbarl_dims[pid] = {'dim1': side, 'dim2': side, 'type': 'PBAR', 'area': area}

                self.log2(f"    Elements: {len(elem_prop)}, PBARL/PBAR props: {len(pbarl_dims)}")
            except Exception as e:
                self.log2(f"    Warning reading BDF: {e}")

        self.log2(f"\n  Total: {len(elem_prop)} elements, {len(pbarl_dims)} bar properties from BDF")

        results = []
        for op2_path in op2_files:
            self.log2(f"\n  Processing: {os.path.basename(op2_path)}")
            try:
                op2 = OP2(debug=False)
                op2.read_op2(op2_path)
                if hasattr(op2, 'cbar_force') and op2.cbar_force:
                    for sc_id, force in op2.cbar_force.items():
                        for i, eid in enumerate(force.element):
                            axial = force.data[0,i,6] if len(force.data.shape)==3 else force.data[i,6]
                            pid = elem_prop.get(eid)
                            d1 = d2 = area = stress = None

                            if pid and pid in self.bar_properties:
                                d1 = self.bar_properties[pid]['dim1']
                                d2 = self.bar_properties[pid]['dim2']
                                area = d1 * d2
                                if area > 0: stress = axial / area
                            elif pid and pid in pbarl_dims:
                                prop_info = pbarl_dims[pid]
                                d1 = prop_info['dim1']
                                d2 = prop_info['dim2']
                                if 'area' in prop_info:
                                    area = prop_info['area']
                                else:
                                    area = d1 * d2
                                if area > 0: stress = axial / area

                            results.append({'OP2': os.path.basename(op2_path), 'Subcase': sc_id, 'Element': eid,
                                'Property': pid, 'Axial': axial, 'Dim1': d1, 'Dim2': d2, 'Area': area, 'Stress': stress})
            except Exception as e:
                self.log2(f"    ERROR: {e}")

        csv_path = os.path.join(out_folder, self.csv_output_name.get())
        with open(csv_path, 'w', newline='') as f:
            w = csv.DictWriter(f, fieldnames=['OP2','Subcase','Element','Property','Axial','Dim1','Dim2','Area','Stress'])
            w.writeheader()
            w.writerows(results)
        self.log2(f"\n  Saved: {csv_path} ({len(results)} rows)")

    def do_combine_stress_inner(self):
        """Combine stress using residual strength (shared logic for step 5 and full run)"""
        self.log2("\n" + "="*60)
        self.log2("STEP 5b: Combine Stress")
        self.log2("="*60)
        out_folder = self.run_output_folder.get()
        stress_csv = os.path.join(out_folder, self.csv_output_name.get())
        if not os.path.exists(stress_csv):
            self.log2("Stress CSV not found")
            return

        stress_df = pd.read_csv(stress_csv)
        lookup = {}
        for _, row in stress_df.iterrows():
            key = (int(row['Subcase']), int(row['Element']))
            lookup[key] = row['Stress'] if pd.notna(row['Stress']) else 0
        elements = stress_df['Element'].unique()

        rs_df = self.residual_strength_df
        cols = rs_df.columns.tolist()
        comb_col = cols[0]

        comp_cols = []
        i = 1
        while i < len(cols) - 1:
            col_name = str(cols[i]).upper()
            next_col_name = str(cols[i+1]).upper()
            if ('CASE' in col_name or 'ID' in col_name) and 'MULT' in next_col_name:
                comp_cols.append((cols[i], cols[i+1]))
                i += 2
            else:
                i += 1

        results = []
        for _, rs_row in rs_df.iterrows():
            comb_lc = rs_row[comb_col]
            if pd.isna(comb_lc): continue
            comb_lc = int(comb_lc)

            for eid in elements:
                total_stress = 0.0
                components = []
                for case_col, mult_col in comp_cols:
                    case_id = rs_row[case_col]
                    multiplier = rs_row[mult_col]
                    if pd.isna(case_id) or pd.isna(multiplier): continue
                    case_id = int(case_id)
                    multiplier = float(multiplier)
                    key = (case_id, int(eid))
                    if key in lookup:
                        stress = lookup[key]
                        if stress is not None:
                            total_stress += stress * multiplier
                            components.append(f"{case_id}*{multiplier}")
                if components:
                    results.append({'Combined_LC': comb_lc, 'Element': eid,
                        'Combined_Stress': total_stress, 'Components': ' + '.join(components)})

        comb_csv = os.path.join(out_folder, self.combined_csv_name.get())
        with open(comb_csv, 'w', newline='') as f:
            w = csv.DictWriter(f, fieldnames=['Combined_LC','Element','Combined_Stress','Components'])
            w.writeheader()
            w.writerows(results)
        self.log2(f"\n  Saved: {comb_csv} ({len(results)} rows)")


    def browse_offset_element_excel(self):
        f = filedialog.askopenfilename(
            title="Select Element IDs Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if f:
            self.offset_element_excel.set(f)
            self.log2(f"Selected Element Excel: {os.path.basename(f)}")

    def start_full_run(self):
        if not self.run_bdfs:
            messagebox.showerror("Error","Add BDF files"); return
        self.btn_full.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_full_run, daemon=True).start()
    
    def do_full_run(self):
        try:
            self.log2("="*60)
            self.log2("FULL RUN (All 3 Steps)")
            self.log2("="*60)
            out_folder = self.run_output_folder.get()
            os.makedirs(out_folder, exist_ok=True)

            # === STEP 1: Update Properties + Apply Offsets ===
            self.log2("\n>>> STEP 1: Update Properties + Apply Offsets")
            self.do_update_and_offset()

            # === STEP 2: Run Nastran ===
            if self.nastran_path.get():
                self.log2("\n>>> STEP 2: Run Nastran")
                nastran = self.nastran_path.get()
                import subprocess
                import time

                bdf_files_in_output = [f for f in os.listdir(out_folder) if f.lower().endswith(('.bdf','.dat','.nas'))]

                for f in bdf_files_in_output:
                    bdf_full_path = os.path.join(out_folder, f)
                    self.log2(f"  Running: {f}")
                    try:
                        cmd = f'"{nastran}" "{bdf_full_path}" out="{out_folder}" scratch=yes batch=no'
                        process = subprocess.Popen(cmd, shell=True)
                        process.wait()
                        self.log2(f"    Completed: {f}")
                    except Exception as e:
                        self.log2(f"    Error running {f}: {e}")

                self.log2("  Waiting for OP2 files...")
                time.sleep(2)
            else:
                self.log2("\n>>> STEP 2: SKIPPED (No Nastran path)")

            # === STEP 3: Post-Process + Combine ===
            self.log2("\n>>> STEP 3: Post-Process + Combine")
            self.do_postprocess_inner()
            if self.residual_strength_df is not None:
                self.do_combine_stress_inner()
            else:
                self.log2("  Combine SKIPPED (No Residual Strength data)")

            self.log2("\n" + "="*60)
            self.log2("FULL RUN COMPLETED!")
            self.log2("="*60)
            self.root.after(0, lambda: messagebox.showinfo("Done","Full run completed!"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn_full.config(state=tk.NORMAL)])



class BarPropertySolverTab:
    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame
        self.maneuver_bdfs = []
        # self.root.title("Bar Property Solver v1.0")
        # self.root.geometry("1300x950")

        # Input paths
        self.bdf_paths = []
        self.property_excel_path = tk.StringVar()
        self.element_excel_path = tk.StringVar()
        self.residual_strength_path = tk.StringVar()
        self.nastran_path = tk.StringVar()
        self.output_folder = tk.StringVar()

        # Thickness ranges
        self.bar_min_thickness = tk.StringVar(value="2.0")
        self.bar_max_thickness = tk.StringVar(value="12.0")
        self.thickness_step = tk.StringVar(value="0.5")

        # Data storage
        self.bdf_model = None
        self.bdf_models = []
        self.bar_properties = {}          # PID -> {'dim1': val, 'dim2': val}
        self.bar_structure_map = {}       # PID -> Structure Name
        self.structure_groups = {}        # Structure Name -> [PID list]
        self.pbarl_dims = {}              # PID -> {'dim1': val, 'dim2': val} from BDF
        self.current_bar_thicknesses = {} # PID -> thickness
        self.current_skin_thicknesses = {}# PID -> thickness (from BDF, for offsets only)
        self.original_bar_thicknesses = {} # PID -> original dim1 from BDF
        self.active_group_pids = None      # Set of PIDs for active sweep group (None = all)
        self.material_densities = {}      # MID -> density
        self.prop_to_material = {}        # PID -> MID
        self.element_areas = {}
        self.bar_lengths = {}
        self.prop_elements = {}
        self.elem_to_prop = {}
        self.element_centroids = {}
        self.bar_elements = []
        self.shell_elements = []
        self.residual_strength_df = None
        self.combination_table = []
        self.landing_elem_ids = []
        self.bar_offset_elem_ids = []

        # Run state
        self.is_running = False
        self.sweep_results = {}  # Structure Name -> [{thickness, stresses, ...}]

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

        ttk.Label(main, text="Bar Property Solver v1.0", font=('Helvetica', 16, 'bold')).pack(pady=10)
        ttk.Label(main, text="Per-structure thickness sweep with combined stress output", foreground='gray').pack()

        # ---- Section 1: Input Files ----
        f1 = ttk.LabelFrame(main, text="1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # All BDF Files
        bdf_frame = ttk.Frame(f1)
        bdf_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bdf_frame, text="All BDF Files:", width=18).pack(side=tk.LEFT, anchor=tk.N)

        bdf_list_frame = ttk.Frame(bdf_frame)
        bdf_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bdf_listbox = tk.Listbox(bdf_list_frame, height=3, width=55, selectmode=tk.SINGLE)
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
        self.maneuver_listbox = tk.Listbox(man_list_frame, height=3, width=55, selectmode=tk.SINGLE)
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

        # Residual Strength
        self._add_file_row(f1, "Residual Strength:", self.residual_strength_path, self.load_residual_strength)
        self.resid_status = ttk.Label(f1, text="Not loaded", foreground="gray")
        self.resid_status.pack(anchor=tk.W)

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

        # Output Folder
        out_frame = ttk.Frame(f1)
        out_frame.pack(fill=tk.X, pady=2)
        ttk.Label(out_frame, text="Output Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="Browse", command=lambda: self.output_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

        # ---- Section 2: Thickness Ranges ----
        f2 = ttk.LabelFrame(main, text="2. Thickness Ranges (mm)", padding=10)
        f2.pack(fill=tk.X, pady=5, padx=10)

        row1 = ttk.Frame(f2)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Bar:  Min:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.bar_min_thickness, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="Max:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.bar_max_thickness, width=8).pack(side=tk.LEFT, padx=5)

        row2 = ttk.Frame(f2)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Step:").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.thickness_step, width=8).pack(side=tk.LEFT, padx=5)

        # ---- Section 3: Actions ----
        f3 = ttk.LabelFrame(main, text="3. Actions", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        btn_row = ttk.Frame(f3)
        btn_row.pack(fill=tk.X)
        self.btn_start = ttk.Button(btn_row, text=">>> START SWEEP", command=self.start_sweep)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_stop = ttk.Button(btn_row, text="STOP", command=self.stop_sweep, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Export", command=self.export_results).pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(f3, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ---- Section 4: Structure Groups ----
        f4 = ttk.LabelFrame(main, text="4. Structure Groups (from Property Excel)", padding=10)
        f4.pack(fill=tk.X, pady=5, padx=10)

        self.group_text = scrolledtext.ScrolledText(f4, height=6, width=100, state=tk.DISABLED)
        self.group_text.pack(fill=tk.X)

        # ---- Section 5: Log ----
        f5 = ttk.LabelFrame(main, text="5. Log", padding=10)
        f5.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f5, height=20, width=100)
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
            self.bdf_model.read_bdf(path, validate=False, xref=True, read_includes=True, encoding='latin-1')

            self.bdf_models = []
            for bdf_path in self.bdf_paths:
                self.log(f"  Loading: {os.path.basename(bdf_path)}")
                model = BDF(debug=False)
                model.read_bdf(bdf_path, validate=False, xref=True, read_includes=True, encoding='latin-1')
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

            # Extract PSHELL thicknesses from BDF (for offset calculations only - not modified)
            self.current_skin_thicknesses = {}
            for pid, prop in self.bdf_model.properties.items():
                if prop.type == 'PSHELL':
                    t = prop.t if hasattr(prop, 't') and prop.t is not None else None
                    if t is not None:
                        self.current_skin_thicknesses[pid] = t
            if self.current_skin_thicknesses:
                self.log(f"  PSHELL thicknesses extracted: {len(self.current_skin_thicknesses)} properties (for offsets only)")

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
        """Load Bar Property sheet with Bar Property ID and Structure Name columns."""
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

            bar_min = float(self.bar_min_thickness.get())

            self.bar_properties = {}
            self.bar_structure_map = {}
            self.structure_groups = {}
            self.current_bar_thicknesses = {}

            for sheet in xl.sheet_names:
                sl = sheet.lower().replace('_', '').replace(' ', '')

                if 'bar' in sl and 'prop' in sl and 'structure' in sl:
                    self.log(f"\n  Reading bar properties from '{sheet}'...")
                    df = pd.read_excel(xl, sheet_name=sheet)
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

                    for _, row in df.iterrows():
                        pid_val = row[pid_col]
                        if pd.isna(pid_val):
                            continue
                        pid = int(pid_val)
                        struct_name = str(row[struct_col]).strip() if struct_col and pd.notna(row[struct_col]) else "DEFAULT"

                        # Use original BDF dim1 if available, otherwise bar_min
                        orig_dim1 = self.original_bar_thicknesses.get(pid, bar_min)
                        orig_dim2 = self.pbarl_dims[pid]['dim2'] if pid in self.pbarl_dims else bar_min

                        self.bar_properties[pid] = {
                            'dim1': orig_dim1,
                            'dim2': orig_dim2,
                        }
                        # Initialize to original BDF value (not bar_min)
                        self.current_bar_thicknesses[pid] = orig_dim1
                        self.bar_structure_map[pid] = struct_name

                        if struct_name not in self.structure_groups:
                            self.structure_groups[struct_name] = []
                        self.structure_groups[struct_name].append(pid)

                    self.log(f"  Loaded {len(self.bar_properties)} bar properties")
                    self.log(f"  Structure groups: {len(self.structure_groups)}")
                    for name, pids in sorted(self.structure_groups.items()):
                        self.log(f"    {name}: {len(pids)} properties")

                    # Log original thickness info
                    bdf_count = sum(1 for pid in self.bar_properties if pid in self.original_bar_thicknesses)
                    self.log(f"  Properties with BDF original dim1: {bdf_count}/{len(self.bar_properties)}")
                    self.log(f"  NOTE: Only active group thicknesses will change during sweep.")

            # Update group display
            self._update_group_display()

            total = len(self.bar_properties)
            groups = len(self.structure_groups)
            self.prop_status.config(
                text=f"Loaded: {total} bar props in {groups} groups",
                foreground="green"
            )

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.prop_status.config(text="Error", foreground="red")

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

    # ==================== RESIDUAL STRENGTH ====================
    def load_residual_strength(self):
        path = self.residual_strength_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Select Residual Strength Excel file")
            return

        self.log("\n" + "=" * 70)
        self.log("LOADING RESIDUAL STRENGTH DATA")
        self.log("=" * 70)

        try:
            xl = pd.ExcelFile(path)
            self.log(f"  Sheets: {xl.sheet_names}")

            res_sh = None
            for sheet in xl.sheet_names:
                sl = sheet.lower().replace(' ', '').replace('_', '')
                if 'residual' in sl or 'strength' in sl:
                    res_sh = sheet
                    break

            if not res_sh:
                res_sh = xl.sheet_names[0]
                self.log(f"  No 'Residual Strength' sheet found, using: {res_sh}")
            else:
                self.log(f"  Found sheet: {res_sh}")

            self.residual_strength_df = pd.read_excel(xl, sheet_name=res_sh)
            self.log(f"  Loaded {len(self.residual_strength_df)} rows")
            self.log(f"  Columns: {list(self.residual_strength_df.columns)}")

            cols = self.residual_strength_df.columns.tolist()
            self.combination_table = []
            i = 1
            while i < len(cols) - 1:
                col_name = str(cols[i]).upper()
                next_col_name = str(cols[i + 1]).upper()
                if ('CASE' in col_name or 'ID' in col_name) and 'MULT' in next_col_name:
                    self.combination_table.append((cols[i], cols[i + 1]))
                    self.log(f"    Found pair: {cols[i]} + {cols[i + 1]}")
                    i += 2
                else:
                    i += 1

            self.log(f"  Total combination pairs: {len(self.combination_table)}")
            self.resid_status.config(
                text=f"{len(self.residual_strength_df)} rows, {len(self.combination_table)} pairs",
                foreground="green"
            )

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.resid_status.config(text="Error", foreground="red")

    # ==================== BDF WRITING ====================
    def _write_bdf_for_model(self, folder, model, original_path):
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
                    # Only update PIDs in the active group; others stay as original BDF
                    active_pids = getattr(self, 'active_group_pids', None)
                    should_update = (pid in self.current_bar_thicknesses and
                                     (active_pids is None or pid in active_pids))
                    if should_update:
                        t = self.current_bar_thicknesses[pid]
                        t = max(0.1, t)
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

    # ==================== OFFSET APPLICATION ====================
    def _apply_offsets(self, bdf_path, folder, maneuver_bdf_path=None):
        if not self.landing_elem_ids and not self.bar_offset_elem_ids:
            self.log("    No offset elements defined, skipping")
            return bdf_path

        try:
            # Use UPDATED maneuver BDF for offset geometry if provided
            if maneuver_bdf_path:
                geom_path = maneuver_bdf_path
                self.log(f"    Using UPDATED maneuver BDF for offsets: {os.path.basename(geom_path)}")
            else:
                geom_path = bdf_path
                self.log(f"    Using working BDF for offset geometry")
            self.log(f"    Applying offsets...")
            bdf = BDF(debug=False)
            bdf.read_bdf(geom_path, validate=False, xref=True, read_includes=True, encoding='latin-1')

            # Landing (shell) offsets: zoffset = -t/2
            landing_offsets = {}
            landing_normals = {}

            for eid in self.landing_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                t = self.current_skin_thicknesses.get(pid, 3.0)
                t = max(0.1, t)
                landing_offsets[eid] = -t / 2.0

                if elem.type in ['CQUAD4', 'CTRIA3']:
                    try:
                        nids = elem.node_ids[:3]
                        nodes = [bdf.nodes[n] for n in nids if n in bdf.nodes]
                        if len(nodes) >= 3:
                            p1, p2, p3 = [np.array(n.get_position()) for n in nodes]
                            normal = np.cross(p2 - p1, p3 - p1)
                            nl = np.linalg.norm(normal)
                            if nl > 1e-10:
                                landing_normals[eid] = normal / nl
                    except:
                        pass

            self.log(f"    Landing offsets: {len(landing_offsets)} elements")

            # Node -> shell mapping for bar offset
            node_to_shells = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    for nid in elem.node_ids:
                        if nid not in node_to_shells:
                            node_to_shells[nid] = []
                        node_to_shells[nid].append(eid)

            # Bar offsets: WA = WB = -normal * (landing_t + bar_t/2)
            bar_offsets = {}
            for eid in self.bar_offset_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                if elem.type not in ['CBAR', 'CBEAM']:
                    continue

                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                bar_t = self.current_bar_thicknesses.get(pid, float(self.bar_min_thickness.get()))
                bar_t = max(0.1, bar_t)

                bar_nodes = elem.node_ids[:2]
                if bar_nodes[0] in node_to_shells and bar_nodes[1] in node_to_shells:
                    common = set(node_to_shells[bar_nodes[0]]) & set(node_to_shells[bar_nodes[1]])
                    max_t = 0
                    best_normal = None

                    for shell_eid in common:
                        if shell_eid in landing_offsets:
                            shell_elem = bdf.elements[shell_eid]
                            shell_pid = shell_elem.pid if hasattr(shell_elem, 'pid') else None
                            if shell_pid:
                                shell_t = self.current_skin_thicknesses.get(shell_pid, 0)
                                if shell_t > max_t:
                                    max_t = shell_t
                                    if shell_eid in landing_normals:
                                        best_normal = landing_normals[shell_eid]

                    if best_normal is not None and max_t > 0:
                        offset_mag = max_t + bar_t / 2.0
                        bar_offsets[eid] = tuple(-best_normal * offset_mag)

            self.log(f"    Bar offsets: {len(bar_offsets)} elements")

            # Apply to BDF file
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
                            new_line = padded[:64] + fmt(landing_offsets[eid]) + '\n'
                            new_lines.append(new_line)
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
                            new_line = padded[:48] + fmt(landing_offsets[eid]) + '\n'
                            new_lines.append(new_line)
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

            # Write as _offseted.bdf
            base, ext = os.path.splitext(bdf_path)
            output_bdf = base + "_offseted" + ext
            with open(output_bdf, 'w', encoding='latin-1') as f:
                f.writelines(new_lines)

            self.log(f"    Offsets applied: {applied_landing} landing, {applied_bar} bar")
            return output_bdf

        except Exception as e:
            self.log(f"    Offset error: {e}")
            return bdf_path

    # ==================== NASTRAN EXECUTION ====================
    def _run_nastran(self, bdf_path, folder):
        nastran = self.nastran_path.get()
        if not nastran or not os.path.exists(nastran):
            self.log("    WARNING: Nastran exe not found!")
            return False
        try:
            cmd = f'"{nastran}" "{bdf_path}" out="{folder}" scratch=yes batch=no'
            proc = subprocess.Popen(cmd, shell=True)
            proc.wait(timeout=600)
            return True
        except:
            return False

    # ==================== STRESS EXTRACTION ====================
    def _extract_stresses(self, folder):
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
                                    d1 = self.current_bar_thicknesses.get(pid, self.bar_properties[pid].get('dim1', 0))
                                    if pid in self.pbarl_dims:
                                        d2 = self.pbarl_dims[pid]['dim2']
                                    else:
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

    # ==================== STRESS COMBINATION ====================
    def _combine_stresses_multi(self, folder, n_bdfs):
        if self.residual_strength_df is None or len(self.combination_table) == 0:
            return None

        try:
            all_stress_data = []
            if n_bdfs > 1:
                for item in os.listdir(folder):
                    subfolder = os.path.join(folder, item)
                    if os.path.isdir(subfolder) and item.startswith('bdf_'):
                        stress_csv = os.path.join(subfolder, 'bar_stress_results.csv')
                        if os.path.exists(stress_csv):
                            all_stress_data.append(pd.read_csv(stress_csv))
            else:
                stress_csv = os.path.join(folder, 'bar_stress_results.csv')
                if os.path.exists(stress_csv):
                    all_stress_data.append(pd.read_csv(stress_csv))

            if not all_stress_data:
                return None

            stress_df = pd.concat(all_stress_data, ignore_index=True)

            lookup = {}
            for _, row in stress_df.iterrows():
                key = (int(row['Subcase']), int(row['Element']))
                stress_val = row['Stress'] if pd.notna(row['Stress']) else 0
                if key not in lookup or abs(stress_val) > abs(lookup[key]):
                    lookup[key] = stress_val

            elements = stress_df['Element'].unique()
            rs_df = self.residual_strength_df
            cols = rs_df.columns.tolist()
            comb_col = cols[0]

            results = []
            for _, rs_row in rs_df.iterrows():
                comb_lc = rs_row[comb_col]
                if pd.isna(comb_lc):
                    continue
                comb_lc = int(comb_lc)

                for eid in elements:
                    total_stress = 0.0
                    components = []

                    for case_col, mult_col in self.combination_table:
                        case_id = rs_row[case_col]
                        multiplier = rs_row[mult_col]
                        if pd.isna(case_id) or pd.isna(multiplier):
                            continue
                        case_id = int(case_id)
                        multiplier = float(multiplier)

                        key = (case_id, int(eid))
                        if key in lookup:
                            stress = lookup[key]
                            if stress is not None:
                                total_stress += stress * multiplier
                                components.append(f"{case_id}*{multiplier}")

                    if components:
                        pid = self.elem_to_prop.get(int(eid))
                        struct_name = self.bar_structure_map.get(pid, '') if pid else ''
                        results.append({
                            'Combined_LC': comb_lc, 'Element': int(eid),
                            'Property': pid, 'Structure': struct_name,
                            'Combined_Stress': total_stress,
                            'Components': ' + '.join(components)
                        })

            if results:
                comb_csv = os.path.join(folder, 'combined_stress_results.csv')
                with open(comb_csv, 'w', newline='') as f:
                    w = csv.DictWriter(f, fieldnames=['Combined_LC', 'Element', 'Property', 'Structure', 'Combined_Stress', 'Components'])
                    w.writeheader()
                    w.writerows(results)
                self.log(f"    Combined stress CSV saved ({len(results)} rows)")
                return results

        except Exception as e:
            self.log(f"    Combine error: {e}")

        return None

    # ==================== SINGLE THICKNESS ITERATION ====================
    def _run_single_iteration(self, folder):
        """Run Nastran for current thicknesses and extract stresses."""
        try:
            n_bdfs = len(self.bdf_models) if self.bdf_models else 1
            all_stresses = []

            # Step 0: Update maneuver BDF with current thicknesses for offset geometry
            updated_maneuver = None
            if self.maneuver_bdfs:
                maneuver_folder = os.path.join(folder, "maneuver_updated")
                os.makedirs(maneuver_folder, exist_ok=True)
                updated_maneuver = self._write_bdf_for_model(
                    maneuver_folder, None, self.maneuver_bdfs[0])
                self.log(f"    Maneuver BDF updated: {os.path.basename(updated_maneuver)}")

            for bdf_idx, bdf_info in enumerate(self.bdf_models):
                bdf_name = bdf_info['name']
                bdf_subfolder = os.path.join(folder, f"bdf_{bdf_idx + 1}_{os.path.splitext(bdf_name)[0]}") if n_bdfs > 1 else folder
                os.makedirs(bdf_subfolder, exist_ok=True)

                if n_bdfs > 1:
                    self.log(f"    --- BDF {bdf_idx + 1}/{n_bdfs}: {bdf_name} ---")

                # 1. Write thermal BDF with current thicknesses
                bdf_path = self._write_bdf_for_model(bdf_subfolder, bdf_info['model'], bdf_info['path'])

                # 2. Apply offsets (from updated maneuver BDF)
                offset_bdf = self._apply_offsets(bdf_path, bdf_subfolder, updated_maneuver)

                # 3. Run Nastran
                self.log("    Running Nastran...")
                self._run_nastran(offset_bdf or bdf_path, bdf_subfolder)

                # 4. Extract stresses
                stresses = self._extract_stresses(bdf_subfolder)
                all_stresses.extend(stresses)

                if stresses:
                    bar_s = [s for s in stresses if s['type'] == 'bar']
                    shell_s = [s for s in stresses if s['type'] == 'shell']
                    self.log(f"    Extracted: {len(bar_s)} bar, {len(shell_s)} shell stresses")

            stresses = all_stresses

            # Combine stresses using Residual Strength
            combined_stresses = None
            if self.residual_strength_df is not None:
                self.log("    Combining stresses...")
                combined_stresses = self._combine_stresses_multi(folder, n_bdfs)
                if combined_stresses:
                    self.log(f"    Combined: {len(combined_stresses)} stress combinations")

            return stresses, combined_stresses

        except Exception as e:
            self.log(f"  Iteration ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            return [], None

    # ==================== MAIN SWEEP LOGIC ====================
    def start_sweep(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one BDF file")
            return
        if not self.bdf_model:
            self.load_bdf()
        if not self.bdf_model:
            return
        if not self.structure_groups:
            messagebox.showerror("Error", "Load Property Excel with structure groups first")
            return

        self.is_running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.sweep_results = {}

        threading.Thread(target=self._run_sweep, daemon=True).start()

    def stop_sweep(self):
        self.is_running = False
        self.log("\n*** STOPPING ***")

    def _run_sweep(self):
        """Main sweep: for each structure group, iterate through thicknesses."""
        try:
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            step = float(self.thickness_step.get())
            output_base = self.output_folder.get()
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            run_folder = os.path.join(output_base, f"BarSweep_{ts}")
            os.makedirs(run_folder, exist_ok=True)

            # Build thickness list
            thicknesses = []
            t = bar_min
            while t <= bar_max + 1e-9:
                thicknesses.append(round(t, 4))
                t += step

            total_groups = len(self.structure_groups)
            total_steps = total_groups * len(thicknesses)
            current_step = 0

            self.log("\n" + "=" * 70)
            self.log("BAR PROPERTY SWEEP")
            self.log("=" * 70)
            self.log(f"  Structure groups: {total_groups}")
            self.log(f"  Thickness range: {bar_min} to {bar_max} mm, step {step}")
            self.log(f"  Thickness values: {thicknesses}")
            self.log(f"  Total iterations: {total_steps}")
            self.log(f"  Output: {run_folder}")

            for group_idx, (struct_name, pids) in enumerate(sorted(self.structure_groups.items())):
                if not self.is_running:
                    break

                self.log(f"\n{'=' * 70}")
                self.log(f"STRUCTURE GROUP: {struct_name} ({group_idx + 1}/{total_groups})")
                self.log(f"  Properties: {len(pids)} -> {sorted(pids)[:20]}{'...' if len(pids) > 20 else ''}")
                self.log(f"  Only {struct_name} thicknesses will change. All other groups keep original BDF values.")
                self.log(f"{'=' * 70}")

                # Create folder for this structure
                group_folder = os.path.join(run_folder, struct_name)
                os.makedirs(group_folder, exist_ok=True)

                group_results = []

                for t_idx, thickness in enumerate(thicknesses):
                    if not self.is_running:
                        break

                    current_step += 1
                    progress = (current_step / total_steps) * 100
                    self.root.after(0, lambda p=progress: self.progress_var.set(p))

                    self.log(f"\n  --- {struct_name}: t = {thickness:.2f} mm ({t_idx + 1}/{len(thicknesses)}) ---")

                    # Set thickness for this group only, mark active group
                    self.active_group_pids = set(pids)
                    for pid in pids:
                        self.current_bar_thicknesses[pid] = thickness

                    # Create iteration folder
                    iter_folder = os.path.join(group_folder, f"t_{thickness:.2f}")
                    os.makedirs(iter_folder, exist_ok=True)

                    # Run iteration
                    stresses, combined_stresses = self._run_single_iteration(iter_folder)

                    # Collect stress results for this group's properties
                    group_stress_data = self._collect_group_stresses(
                        pids, stresses, combined_stresses
                    )

                    result_entry = {
                        'thickness': thickness,
                        'stresses': group_stress_data,
                        'raw_stress_count': len(stresses),
                        'combined_count': len(combined_stresses) if combined_stresses else 0,
                    }
                    group_results.append(result_entry)

                    # Log summary for this thickness
                    if group_stress_data:
                        max_stress = max(s['stress'] for s in group_stress_data) if group_stress_data else 0
                        avg_stress = np.mean([s['stress'] for s in group_stress_data]) if group_stress_data else 0
                        self.log(f"    Max stress: {max_stress:.2f}, Avg stress: {avg_stress:.2f}")
                        self.log(f"    Elements in group: {len(group_stress_data)}")

                self.sweep_results[struct_name] = group_results

                # Save group summary
                self._save_group_summary(group_folder, struct_name, group_results, pids)

                # Reset thicknesses for this group back to original BDF values
                self.active_group_pids = None
                for pid in pids:
                    self.current_bar_thicknesses[pid] = self.original_bar_thicknesses.get(pid, bar_min)

            # Save overall summary
            if self.is_running:
                self._save_overall_summary(run_folder)

            self.log(f"\n{'=' * 70}")
            self.log("SWEEP COMPLETE")
            self.log(f"  Output folder: {run_folder}")
            self.log(f"{'=' * 70}")

        except Exception as e:
            self.log(f"\nSWEEP ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.btn_stop.config(state=tk.DISABLED))
            self.root.after(0, lambda: self.progress_var.set(100))

    # ==================== STRESS COLLECTION PER GROUP ====================
    def _collect_group_stresses(self, group_pids, stresses, combined_stresses):
        """Collect stress values for elements belonging to the given property group."""
        group_pid_set = set(group_pids)
        result = []

        # Build combined stress lookup (max positive stress per element)
        combined_lookup = {}
        if combined_stresses:
            for cs in combined_stresses:
                eid = cs['Element']
                pid = self.elem_to_prop.get(eid)
                if pid and pid in group_pid_set:
                    comb_stress = cs['Combined_Stress']
                    if eid not in combined_lookup or comb_stress > combined_lookup[eid]:
                        combined_lookup[eid] = comb_stress

        # Prefer combined stress; fall back to raw stress
        seen_elements = set()
        for s in stresses:
            eid = s['eid']
            pid = self.elem_to_prop.get(eid)
            if pid not in group_pid_set:
                continue
            if eid in seen_elements:
                continue
            seen_elements.add(eid)

            if eid in combined_lookup:
                stress_val = combined_lookup[eid]
                source = 'combined'
            else:
                stress_val = s['stress']
                source = 'raw'

            result.append({
                'eid': eid,
                'pid': pid,
                'stress': stress_val,
                'source': source,
            })

        return result

    # ==================== SUMMARY GENERATION ====================
    def _save_group_summary(self, group_folder, struct_name, group_results, pids):
        """Save a CSV summary: thickness vs stress for this structure group."""
        try:
            # Per-property summary: each row = thickness, columns = per-PID max stress
            summary_rows = []
            all_pids_in_results = set()

            for gr in group_results:
                for s in gr['stresses']:
                    all_pids_in_results.add(s['pid'])

            sorted_pids = sorted(all_pids_in_results)

            for gr in group_results:
                row = {'Thickness_mm': gr['thickness']}

                # Per-PID max positive stress
                pid_stresses = {}
                for s in gr['stresses']:
                    pid = s['pid']
                    stress_val = s['stress']
                    if pid not in pid_stresses or stress_val > pid_stresses[pid]:
                        pid_stresses[pid] = stress_val

                for pid in sorted_pids:
                    row[f'PID_{pid}_Stress'] = pid_stresses.get(pid, '')

                # Aggregates (max positive)
                all_stress_vals = [s['stress'] for s in gr['stresses']] if gr['stresses'] else []
                row['Max_Stress'] = max(all_stress_vals) if all_stress_vals else ''
                row['Avg_Stress'] = np.mean(all_stress_vals) if all_stress_vals else ''
                row['Min_Stress'] = min(all_stress_vals) if all_stress_vals else ''
                row['Element_Count'] = len(gr['stresses'])

                summary_rows.append(row)

            if summary_rows:
                # Build column order
                columns = ['Thickness_mm']
                columns += [f'PID_{pid}_Stress' for pid in sorted_pids]
                columns += ['Max_Stress', 'Avg_Stress', 'Min_Stress', 'Element_Count']

                df = pd.DataFrame(summary_rows, columns=columns)
                csv_path = os.path.join(group_folder, f"{struct_name}_summary.csv")
                df.to_csv(csv_path, index=False)
                self.log(f"\n  Summary saved: {csv_path}")

                # Also save element-level detail (max positive stress, not abs)
                detail_rows = []
                for gr in group_results:
                    for s in gr['stresses']:
                        detail_rows.append({
                            'Thickness_mm': gr['thickness'],
                            'Element_ID': s['eid'],
                            'Property_ID': s['pid'],
                            'Stress': s['stress'],
                            'Source': s['source'],
                        })

                if detail_rows:
                    detail_df = pd.DataFrame(detail_rows)
                    detail_csv = os.path.join(group_folder, f"{struct_name}_detail.csv")
                    detail_df.to_csv(detail_csv, index=False)
                    self.log(f"  Detail saved: {detail_csv}")

                # Power law fit per element: stress = a * thickness^b
                self._save_powerlaw_fit(group_folder, struct_name, group_results)

                # Element-based load/strain driven summary
                self._save_load_strain_summary(group_folder, struct_name, group_results)

        except Exception as e:
            self.log(f"  Summary save error: {e}")

    def _save_powerlaw_fit(self, group_folder, struct_name, group_results):
        """Fit power law (stress = a * thickness^b) per element and save summary."""
        try:
            # Build per-element data: eid -> [(thickness, stress), ...]
            elem_data = {}
            elem_pid = {}
            for gr in group_results:
                t = gr['thickness']
                for s in gr['stresses']:
                    eid = s['eid']
                    if eid not in elem_data:
                        elem_data[eid] = []
                        elem_pid[eid] = s['pid']
                    elem_data[eid].append((t, s['stress']))

            if not elem_data:
                return

            def power_law(x, a, b):
                return a * np.power(x, b)

            fit_rows = []
            for eid in sorted(elem_data.keys()):
                points = elem_data[eid]
                pid = elem_pid[eid]

                t_arr = np.array([p[0] for p in points])
                s_arr = np.array([p[1] for p in points])

                # Need at least 2 positive points for power law fit
                valid = (t_arr > 0) & (s_arr > 0)
                t_valid = t_arr[valid]
                s_valid = s_arr[valid]

                if len(t_valid) < 2:
                    fit_rows.append({
                        'Element_ID': eid,
                        'Property_ID': pid,
                        'Structure': struct_name,
                        'a': '',
                        'b': '',
                        'R2': '',
                        'N_Points': len(t_valid),
                        'Status': 'insufficient_data',
                    })
                    continue

                try:
                    # Log-space linear regression for power law: log(s) = log(a) + b*log(t)
                    log_t = np.log(t_valid)
                    log_s = np.log(s_valid)
                    coeffs = np.polyfit(log_t, log_s, 1)
                    b_val = coeffs[0]
                    a_val = np.exp(coeffs[1])

                    # R^2 calculation
                    s_pred = a_val * np.power(t_valid, b_val)
                    ss_res = np.sum((s_valid - s_pred) ** 2)
                    ss_tot = np.sum((s_valid - np.mean(s_valid)) ** 2)
                    r2 = 1.0 - (ss_res / ss_tot) if ss_tot > 0 else 0.0

                    fit_rows.append({
                        'Element_ID': eid,
                        'Property_ID': pid,
                        'Structure': struct_name,
                        'a': round(a_val, 4),
                        'b': round(b_val, 4),
                        'R2': round(r2, 6),
                        'N_Points': len(t_valid),
                        'Status': 'ok',
                    })
                except Exception:
                    fit_rows.append({
                        'Element_ID': eid,
                        'Property_ID': pid,
                        'Structure': struct_name,
                        'a': '',
                        'b': '',
                        'R2': '',
                        'N_Points': len(t_valid),
                        'Status': 'fit_failed',
                    })

            if fit_rows:
                fit_df = pd.DataFrame(fit_rows)
                fit_csv = os.path.join(group_folder, f"{struct_name}_powerlaw_fit.csv")
                fit_df.to_csv(fit_csv, index=False)
                self.log(f"  Power law fit saved: {fit_csv}")

                # Log summary stats
                ok_fits = [r for r in fit_rows if r['Status'] == 'ok']
                if ok_fits:
                    r2_vals = [r['R2'] for r in ok_fits]
                    self.log(f"    Fitted {len(ok_fits)}/{len(fit_rows)} elements")
                    self.log(f"    R2 range: {min(r2_vals):.4f} - {max(r2_vals):.4f}")
                    self.log(f"    R2 mean:  {np.mean(r2_vals):.4f}")

        except Exception as e:
            self.log(f"  Power law fit error: {e}")

    def _save_load_strain_summary(self, group_folder, struct_name, group_results):
        """Element-based load/strain driven summary.
        Compares stress at min vs max thickness for each element.
        If stress drops significantly with thickness increase -> LOAD driven (stress ~ 1/A)
        If stress stays similar despite thickness increase -> STRAIN driven (stress ~ geometry)
        """
        try:
            if len(group_results) < 2:
                return

            t_min = group_results[0]['thickness']
            t_max = group_results[-1]['thickness']

            # Build stress at min and max thickness per element
            stress_at_min = {}  # eid -> stress
            stress_at_max = {}  # eid -> stress
            elem_pid_map = {}

            for s in group_results[0]['stresses']:
                stress_at_min[s['eid']] = s['stress']
                elem_pid_map[s['eid']] = s['pid']

            for s in group_results[-1]['stresses']:
                stress_at_max[s['eid']] = s['stress']
                if s['eid'] not in elem_pid_map:
                    elem_pid_map[s['eid']] = s['pid']

            # Also get power law fit data for each element
            elem_fit = {}
            elem_data = {}
            for gr in group_results:
                t = gr['thickness']
                for s in gr['stresses']:
                    eid = s['eid']
                    if eid not in elem_data:
                        elem_data[eid] = []
                    elem_data[eid].append((t, s['stress']))

            for eid, points in elem_data.items():
                t_arr = np.array([p[0] for p in points])
                s_arr = np.array([p[1] for p in points])
                valid = (t_arr > 0) & (s_arr > 0)
                t_v = t_arr[valid]
                s_v = s_arr[valid]
                if len(t_v) >= 2:
                    try:
                        log_t = np.log(t_v)
                        log_s = np.log(s_v)
                        coeffs = np.polyfit(log_t, log_s, 1)
                        b_val = coeffs[0]
                        a_val = np.exp(coeffs[1])
                        s_pred = a_val * np.power(t_v, b_val)
                        ss_res = np.sum((s_v - s_pred) ** 2)
                        ss_tot = np.sum((s_v - np.mean(s_v)) ** 2)
                        r2 = 1.0 - (ss_res / ss_tot) if ss_tot > 0 else 0.0
                        elem_fit[eid] = {'a': a_val, 'b': b_val, 'r2': r2}
                    except:
                        pass

            all_eids = sorted(set(stress_at_min.keys()) | set(stress_at_max.keys()))

            rows = []
            for eid in all_eids:
                pid = elem_pid_map.get(eid, '')
                s_min_t = stress_at_min.get(eid)
                s_max_t = stress_at_max.get(eid)

                if s_min_t is None or s_max_t is None:
                    continue

                stress_diff = s_min_t - s_max_t
                stress_pct = (stress_diff / s_min_t * 100) if s_min_t != 0 else 0
                thickness_ratio = t_max / t_min if t_min > 0 else 0

                # Classification based on power law exponent b
                # b ~ -1: pure load driven (stress = F/A, inversely proportional)
                # b ~ 0: strain driven (stress doesn't change much with thickness)
                fit = elem_fit.get(eid)
                if fit:
                    b_val = fit['b']
                    a_val = fit['a']
                    r2 = fit['r2']
                    if b_val <= -0.5:
                        classification = 'LOAD_DRIVEN'
                    elif b_val <= -0.1:
                        classification = 'PARTIALLY_LOAD'
                    elif b_val <= 0.1:
                        classification = 'STRAIN_DRIVEN'
                    else:
                        classification = 'INCREASING'
                else:
                    a_val = ''
                    b_val = ''
                    r2 = ''
                    classification = 'NO_FIT'

                rows.append({
                    'Element_ID': eid,
                    'Property_ID': pid,
                    'Structure': struct_name,
                    'T_Min_mm': t_min,
                    'Stress_at_T_Min': round(s_min_t, 4),
                    'T_Max_mm': t_max,
                    'Stress_at_T_Max': round(s_max_t, 4),
                    'Stress_Drop': round(stress_diff, 4),
                    'Stress_Drop_Pct': round(stress_pct, 2),
                    'Thickness_Ratio': round(thickness_ratio, 4),
                    'Power_a': round(a_val, 4) if isinstance(a_val, float) else '',
                    'Power_b': round(b_val, 4) if isinstance(b_val, float) else '',
                    'R2': round(r2, 6) if isinstance(r2, float) else '',
                    'Classification': classification,
                })

            if rows:
                df = pd.DataFrame(rows)
                csv_path = os.path.join(group_folder, f"{struct_name}_load_strain_summary.csv")
                df.to_csv(csv_path, index=False)
                self.log(f"  Load/Strain summary saved: {csv_path}")

                # Log classification counts
                cls_counts = {}
                for r in rows:
                    c = r['Classification']
                    cls_counts[c] = cls_counts.get(c, 0) + 1
                for c, cnt in sorted(cls_counts.items()):
                    self.log(f"    {c}: {cnt} elements")

        except Exception as e:
            self.log(f"  Load/Strain summary error: {e}")

    def _save_overall_summary(self, run_folder):
        """Save overall summary across all structure groups."""
        try:
            self.log("\n" + "-" * 70)
            self.log("OVERALL SUMMARY")
            self.log("-" * 70)

            overall_rows = []
            for struct_name, results in sorted(self.sweep_results.items()):
                for gr in results:
                    all_stress_vals = [s['stress'] for s in gr['stresses']] if gr['stresses'] else []
                    overall_rows.append({
                        'Structure': struct_name,
                        'Thickness_mm': gr['thickness'],
                        'Max_Stress': max(all_stress_vals) if all_stress_vals else '',
                        'Avg_Stress': np.mean(all_stress_vals) if all_stress_vals else '',
                        'Min_Stress': min(all_stress_vals) if all_stress_vals else '',
                        'Element_Count': len(gr['stresses']),
                    })

                    if all_stress_vals:
                        self.log(f"  {struct_name} | t={gr['thickness']:.2f}mm | "
                                 f"Max={max(all_stress_vals):.2f} | Avg={np.mean(all_stress_vals):.2f}")

            if overall_rows:
                df = pd.DataFrame(overall_rows)
                csv_path = os.path.join(run_folder, "overall_summary.csv")
                df.to_csv(csv_path, index=False)
                self.log(f"\n  Overall summary saved: {csv_path}")

            # Save run config
            config_path = os.path.join(run_folder, "run_config.txt")
            with open(config_path, 'w') as f:
                f.write(f"Bar Property Solver v1.0\n")
                f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Bar Min: {self.bar_min_thickness.get()} mm\n")
                f.write(f"Bar Max: {self.bar_max_thickness.get()} mm\n")
                f.write(f"Step: {self.thickness_step.get()} mm\n")
                f.write(f"BDF Files: {len(self.bdf_paths)}\n")
                for p in self.bdf_paths:
                    f.write(f"  {p}\n")
                f.write(f"Structure Groups: {len(self.structure_groups)}\n")
                for name, pids in sorted(self.structure_groups.items()):
                    f.write(f"  {name}: {len(pids)} properties\n")

        except Exception as e:
            self.log(f"  Overall summary error: {e}")

    # ==================== EXPORT ====================
    def export_results(self):
        if not self.sweep_results:
            messagebox.showerror("Error", "No results to export. Run sweep first.")
            return

        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = os.path.join(self.output_folder.get(), f"bar_sweep_results_{ts}.xlsx")

            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                # Overall sheet
                overall_rows = []
                for struct_name, results in sorted(self.sweep_results.items()):
                    for gr in results:
                        all_s = [abs(s['stress']) for s in gr['stresses']] if gr['stresses'] else []
                        overall_rows.append({
                            'Structure': struct_name,
                            'Thickness_mm': gr['thickness'],
                            'Max_Stress': max(all_s) if all_s else None,
                            'Avg_Stress': np.mean(all_s) if all_s else None,
                            'Min_Stress': min(all_s) if all_s else None,
                            'Element_Count': len(gr['stresses']),
                        })
                if overall_rows:
                    pd.DataFrame(overall_rows).to_excel(writer, sheet_name='Overall', index=False)

                # Per-group sheets
                for struct_name, results in sorted(self.sweep_results.items()):
                    sheet_name = struct_name[:31]  # Excel 31 char limit
                    detail_rows = []
                    for gr in results:
                        for s in gr['stresses']:
                            detail_rows.append({
                                'Thickness_mm': gr['thickness'],
                                'Element_ID': s['eid'],
                                'Property_ID': s['pid'],
                                'Stress': s['stress'],
                                'Abs_Stress': abs(s['stress']),
                                'Source': s['source'],
                            })
                    if detail_rows:
                        pd.DataFrame(detail_rows).to_excel(writer, sheet_name=sheet_name, index=False)

            self.log(f"\nExported: {path}")
            messagebox.showinfo("Export", f"Saved to:\n{path}")

        except Exception as e:
            self.log(f"Export error: {e}")
            messagebox.showerror("Error", str(e))




class StructureOptimizationTab:
    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame
        self.maneuver_bdfs = []
        # self.root.title("Thickness Iteration Tool v3.0")
        # self.root.geometry("1300x1000")

        # Input paths
        self.input_bdf_path = tk.StringVar()
        self.bdf_paths = []  # List of BDF paths for multi-BDF support
        self.allowable_excel_path = tk.StringVar()
        self.property_excel_path = tk.StringVar()
        self.element_excel_path = tk.StringVar()
        self.residual_strength_path = tk.StringVar()  # NEW: Residual Strength Excel
        self.nastran_path = tk.StringVar()
        self.output_folder = tk.StringVar()

        # Thickness ranges
        self.bar_min_thickness = tk.StringVar(value="2.0")
        self.bar_max_thickness = tk.StringVar(value="12.0")
        self.skin_min_thickness = tk.StringVar(value="3.0")
        self.skin_max_thickness = tk.StringVar(value="18.0")
        self.thickness_step = tk.StringVar(value="0.5")
        self.bar_skin_search_distance = tk.StringVar(value="150.0")  # mm - for proximity-based optimization
        self.use_gfem_thickness = tk.BooleanVar(value=False)  # Use GFEM thickness as initial/min
        self.keep_bars_constant = tk.BooleanVar(value=False)  # Keep bar thicknesses constant (from Excel)
        self.keep_skins_constant = tk.BooleanVar(value=False)  # Keep skin thicknesses constant (from Excel)

        # Store GFEM thicknesses from Excel (original design values)
        self.gfem_bar_thicknesses = {}  # PID -> thickness from Excel
        self.gfem_skin_thicknesses = {}  # PID -> thickness from Excel

        # RF settings
        self.target_rf = tk.StringVar(value="1.0")
        self.rf_tolerance = tk.StringVar(value="0.05")
        self.r2_threshold_var = tk.StringVar(value="0.95")
        self.min_data_points_var = tk.StringVar(value="3")

        # Optimization settings
        self.max_iterations = tk.StringVar(value="50")
        self.algorithm_var = tk.StringVar(value="Bottom-Up (Min to Target)")

        # GA Parameters
        self.ga_population = tk.StringVar(value="50")
        self.ga_generations = tk.StringVar(value="100")
        self.ga_mutation_rate = tk.StringVar(value="0.1")
        self.ga_crossover_rate = tk.StringVar(value="0.8")

        # Data storage
        self.bdf_model = None
        self.bar_properties = {}
        self.skin_properties = {}
        self.pbarl_dims = {}  # PID -> {'dim1': val, 'dim2': val} from BDF PBARL

        # Per-property current thicknesses
        self.current_bar_thicknesses = {}  # PID -> thickness
        self.current_skin_thicknesses = {}  # PID -> thickness

        # Material densities from BDF (MAT1, MAT8, MAT9, etc.)
        self.material_densities = {}  # MID -> density
        self.prop_to_material = {}  # PID -> MID

        # Geometry
        self.element_areas = {}
        self.bar_lengths = {}
        self.prop_elements = {}
        self.elem_to_prop = {}
        self.element_centroids = {}  # EID -> (x, y, z) centroid coordinates
        self.bar_to_nearby_skins = {}  # Bar PID -> list of nearby Skin PIDs
        self.skin_to_nearby_bars = {}  # Skin PID -> list of nearby Bar PIDs (reverse mapping)

        # Allowable fits (RF Check v2.1 format)
        self.allowable_interp = {}  # PID -> {a, b, r2, excluded, n_pts}
        self.allowable_elem_interp = {}  # EID -> {a, b, r2, excluded, property}
        self.allowable_df = None

        # Residual Strength data for stress combination
        self.residual_strength_df = None
        self.combination_table = []  # [(case_col, mult_col), ...]

        # Reference stresses for surrogate model
        self.reference_stresses = {}  # PID -> stress at reference thickness
        self.reference_thickness = {}  # PID -> reference thickness

        # Offset elements
        self.landing_elem_ids = []
        self.bar_offset_elem_ids = []

        # Results
        self.iteration_results = []
        self.best_solution = None
        self.is_running = False

        self.setup_ui()

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

        ttk.Label(main, text="Thickness Iteration Tool v2.1", font=('Helvetica', 16, 'bold')).pack(pady=10)
        ttk.Label(main, text="Per-property optimization with RF Check v2.1 allowable logic", foreground='gray').pack()

        # Section 1: Input Files
        f1 = ttk.LabelFrame(main, text="1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # All BDF Files
        bdf_frame = ttk.Frame(f1)
        bdf_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bdf_frame, text="All BDF Files:", width=18).pack(side=tk.LEFT, anchor=tk.N)

        bdf_list_frame = ttk.Frame(bdf_frame)
        bdf_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.bdf_listbox = tk.Listbox(bdf_list_frame, height=3, width=55, selectmode=tk.SINGLE)
        self.bdf_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        bdf_scroll = ttk.Scrollbar(bdf_list_frame, orient=tk.VERTICAL, command=self.bdf_listbox.yview)
        bdf_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.bdf_listbox.config(yscrollcommand=bdf_scroll.set)

        bdf_btn_frame = ttk.Frame(bdf_frame)
        bdf_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(bdf_btn_frame, text="Add", command=self.add_bdf, width=8).pack(pady=1)
        ttk.Button(bdf_btn_frame, text="Remove", command=self.remove_bdf, width=8).pack(pady=1)

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
        self.maneuver_listbox = tk.Listbox(man_list_frame, height=3, width=55, selectmode=tk.SINGLE)
        self.maneuver_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        man_scroll = ttk.Scrollbar(man_list_frame, orient=tk.VERTICAL, command=self.maneuver_listbox.yview)
        man_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.maneuver_listbox.configure(yscrollcommand=man_scroll.set)

        man_btn_frame = ttk.Frame(man_frame)
        man_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(man_btn_frame, text="Add", command=self._add_maneuver_bdf, width=8).pack(pady=1)
        ttk.Button(man_btn_frame, text="Remove", command=self._remove_maneuver_bdf, width=8).pack(pady=1)


        for label, var, cmd, status_name in [
            ("Property Excel:", self.property_excel_path, self.load_properties, "prop_status"),
            ("Allowable Excel:", self.allowable_excel_path, self.rf_load_allowable, "allow_status"),
            ("Residual Strength:", self.residual_strength_path, self.load_residual_strength, "resid_status"),
            ("Offset Element IDs:", self.element_excel_path, self.load_element_ids, "elem_status"),
        ]:
            row = ttk.Frame(f1)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=label, width=18).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=var, width=50).pack(side=tk.LEFT, padx=5)
            ttk.Button(row, text="Browse", command=lambda v=var: self.browse_file(v)).pack(side=tk.LEFT)
            ttk.Button(row, text="Load", command=cmd).pack(side=tk.LEFT, padx=2)
            setattr(self, status_name, ttk.Label(f1, text="Not loaded", foreground="gray"))
            getattr(self, status_name).pack(anchor=tk.W)

        row = ttk.Frame(f1)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="Nastran Exe:", width=18).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.nastran_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Browse", command=lambda: self.browse_file(self.nastran_path)).pack(side=tk.LEFT)

        row = ttk.Frame(f1)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="Output Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.output_folder, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Browse", command=self.browse_folder).pack(side=tk.LEFT)

        # Section 2: Thickness Ranges
        f2 = ttk.LabelFrame(main, text="2. Thickness Ranges (mm)", padding=10)
        f2.pack(fill=tk.X, pady=5, padx=10)

        row = ttk.Frame(f2)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Bar:").pack(side=tk.LEFT)
        ttk.Label(row, text="Min:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(row, textvariable=self.bar_min_thickness, width=8).pack(side=tk.LEFT)
        ttk.Label(row, text="Max:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(row, textvariable=self.bar_max_thickness, width=8).pack(side=tk.LEFT)

        row = ttk.Frame(f2)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Skin:").pack(side=tk.LEFT)
        ttk.Label(row, text="Min:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(row, textvariable=self.skin_min_thickness, width=8).pack(side=tk.LEFT)
        ttk.Label(row, text="Max:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(row, textvariable=self.skin_max_thickness, width=8).pack(side=tk.LEFT)

        row = ttk.Frame(f2)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Step:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.thickness_step, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Bar-Skin Distance (mm):").pack(side=tk.LEFT, padx=10)
        ttk.Entry(row, textvariable=self.bar_skin_search_distance, width=8).pack(side=tk.LEFT)

        row = ttk.Frame(f2)
        row.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(row, text="Use GFEM thickness for initial/minimum",
                        variable=self.use_gfem_thickness).pack(side=tk.LEFT)
        ttk.Label(row, text="(Each property uses its Excel thickness as start & min)",
                  foreground="gray").pack(side=tk.LEFT, padx=10)

        row = ttk.Frame(f2)
        row.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(row, text="Keep Bar thicknesses constant (from Excel)",
                        variable=self.keep_bars_constant).pack(side=tk.LEFT)
        ttk.Label(row, text="(Bar props stay at Excel values, only skins optimized)",
                  foreground="gray").pack(side=tk.LEFT, padx=10)

        row = ttk.Frame(f2)
        row.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(row, text="Keep Skin thicknesses constant (from Excel)",
                        variable=self.keep_skins_constant).pack(side=tk.LEFT)
        ttk.Label(row, text="(Skin props stay at Excel values, only bars optimized)",
                  foreground="gray").pack(side=tk.LEFT, padx=10)

        # Section 3: RF Settings
        f3 = ttk.LabelFrame(main, text="3. RF Settings", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        row = ttk.Frame(f3)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Target RF:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.target_rf, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Tolerance:").pack(side=tk.LEFT, padx=10)
        ttk.Entry(row, textvariable=self.rf_tolerance, width=8).pack(side=tk.LEFT)

        row = ttk.Frame(f3)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="R² Threshold:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.r2_threshold_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Min Points:").pack(side=tk.LEFT, padx=10)
        ttk.Entry(row, textvariable=self.min_data_points_var, width=8).pack(side=tk.LEFT)

        # Section 4: Optimization Algorithm
        f4 = ttk.LabelFrame(main, text="4. Optimization Algorithm", padding=10)
        f4.pack(fill=tk.X, pady=5, padx=10)

        row = ttk.Frame(f4)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Algorithm:").pack(side=tk.LEFT)
        algo_combo = ttk.Combobox(row, textvariable=self.algorithm_var, width=25, state='readonly')
        algo_combo['values'] = ('Bottom-Up (Min to Target)', 'Decoupled Min Weight', 'Coupled Efficiency Analysis')
        algo_combo.pack(side=tk.LEFT, padx=5)
        algo_combo.bind('<<ComboboxSelected>>', self._on_algorithm_change)

        ttk.Label(row, text="Max Iterations:").pack(side=tk.LEFT, padx=10)
        ttk.Entry(row, textvariable=self.max_iterations, width=8).pack(side=tk.LEFT, padx=5)

        # GA Parameters Frame
        self.ga_frame = ttk.LabelFrame(f4, text="GA Parameters", padding=5)
        self.ga_frame.pack(fill=tk.X, pady=5)

        row = ttk.Frame(self.ga_frame)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="Population:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.ga_population, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Generations:").pack(side=tk.LEFT, padx=10)
        ttk.Entry(row, textvariable=self.ga_generations, width=8).pack(side=tk.LEFT, padx=5)

        row = ttk.Frame(self.ga_frame)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="Mutation Rate:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.ga_mutation_rate, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Crossover Rate:").pack(side=tk.LEFT, padx=10)
        ttk.Entry(row, textvariable=self.ga_crossover_rate, width=8).pack(side=tk.LEFT, padx=5)

        # Initially hide GA parameters
        self.ga_frame.pack_forget()

        # Section 5: Actions
        f5 = ttk.LabelFrame(main, text="5. Actions", padding=10)
        f5.pack(fill=tk.X, pady=5, padx=10)

        row = ttk.Frame(f5)
        row.pack(fill=tk.X, pady=5)
        self.btn_start = ttk.Button(row, text=">>> START OPTIMIZATION <<<", command=self.start_optimization, width=25)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_stop = ttk.Button(row, text="STOP", command=self.stop_optimization, width=10, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Export", command=self.export_results).pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(f5, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        self.progress_label = ttk.Label(f5, text="Ready")
        self.progress_label.pack(anchor=tk.W)

        # Section 6: Results
        f6 = ttk.LabelFrame(main, text="6. Results", padding=10)
        f6.pack(fill=tk.X, pady=5, padx=10)

        self.result_summary = ttk.Label(f6, text="Run optimization to see results", font=('Helvetica', 11, 'bold'), foreground='blue')
        self.result_summary.pack(anchor=tk.W, pady=5)

        self.best_text = tk.Text(f6, height=10, font=('Courier', 9))
        self.best_text.pack(fill=tk.X, pady=5)

        # Section 7: Log
        f7 = ttk.LabelFrame(main, text="7. Log", padding=10)
        f7.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f7, height=15, font=('Courier', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log("="*70)
        self.log("Thickness Iteration Tool v3.0")
        self.log("Bottom-Up (Min to Target) Optimization")
        self.log("="*70)

    def _on_algorithm_change(self, event=None):
        """Show/hide GA parameters based on selected algorithm."""
        algo = self.algorithm_var.get()
        if 'GA' in algo:
            self.ga_frame.pack(fill=tk.X, pady=5)
        else:
            self.ga_frame.pack_forget()

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def update_progress(self, val, txt=""):
        self.progress['value'] = val
        self.progress_label.config(text=txt)
        self.root.update_idletasks()

    def browse_file(self, var):
        f = filedialog.askopenfilename()
        if f:
            var.set(f)

    def browse_folder(self):
        f = filedialog.askdirectory()
        if f:
            self.output_folder.set(f)


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
        """Add BDF file to the list."""
        files = filedialog.askopenfilenames(
            title="Select BDF files",
            filetypes=[("BDF files", "*.bdf *.dat *.nas"), ("All files", "*.*")]
        )
        for f in files:
            if f and f not in self.bdf_paths:
                self.bdf_paths.append(f)
                self.bdf_listbox.insert(tk.END, os.path.basename(f))
        self.bdf_status.config(text=f"{len(self.bdf_paths)} BDF files selected", foreground="blue")

    def remove_bdf(self):
        """Remove selected BDF from list."""
        selection = self.bdf_listbox.curselection()
        if selection:
            idx = selection[0]
            self.bdf_listbox.delete(idx)
            del self.bdf_paths[idx]
        self.bdf_status.config(text=f"{len(self.bdf_paths)} BDF files selected", foreground="blue")

    # ==================== BDF LOADING ====================
    def load_bdf(self):
        """Load all BDF files in the list."""
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one BDF file")
            return

        self.log("\n" + "="*70)
        self.log("LOADING BDF FILES")
        self.log("="*70)
        self.log(f"  BDF files to load: {len(self.bdf_paths)}")
        for i, p in enumerate(self.bdf_paths):
            self.log(f"    {i+1}. {os.path.basename(p)}")

        try:
            # Load the FIRST BDF as the main model (for geometry extraction)
            path = self.bdf_paths[0]
            self.log(f"\n  Loading main BDF: {os.path.basename(path)}")
            self.bdf_model = BDF(debug=False)
            self.bdf_model.read_bdf(path, validate=False, xref=True, read_includes=True, encoding='latin-1')

            # Store all BDF models for multi-BDF processing
            self.bdf_models = []
            for bdf_path in self.bdf_paths:
                self.log(f"  Loading: {os.path.basename(bdf_path)}")
                model = BDF(debug=False)
                model.read_bdf(bdf_path, validate=False, xref=True, read_includes=True, encoding='latin-1')
                self.bdf_models.append({'path': bdf_path, 'model': model, 'name': os.path.basename(bdf_path)})

            self.log(f"  Nodes: {len(self.bdf_model.nodes)}")
            self.log(f"  Elements: {len(self.bdf_model.elements)}")
            self.log(f"  Properties: {len(self.bdf_model.properties)}")
            self.log(f"  Materials: {len(self.bdf_model.materials)}")

            # Extract material densities (MAT1, MAT2, MAT8, MAT9, etc.)
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
            self.bar_elements = []  # List of bar element IDs
            self.shell_elements = []  # List of shell element IDs

            shell_count = bar_count = 0

            for eid, elem in self.bdf_model.elements.items():
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid:
                    self.elem_to_prop[eid] = pid
                    if pid not in self.prop_elements:
                        self.prop_elements[pid] = []
                    self.prop_elements[pid].append(eid)

                # Try to get centroid
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

            # Extract PBARL dimensions from BDF (for accurate Dim2 values)
            self.pbarl_dims = {}
            for pid, prop in self.bdf_model.properties.items():
                if prop.type == 'PBARL':
                    dims = prop.dim if hasattr(prop, 'dim') else []
                    if len(dims) >= 2:
                        self.pbarl_dims[pid] = {
                            'dim1': dims[0],
                            'dim2': dims[1],
                        }
                    elif len(dims) == 1:
                        self.pbarl_dims[pid] = {
                            'dim1': dims[0],
                            'dim2': dims[0],
                        }
            if self.pbarl_dims:
                self.log(f"  PBARL dimensions extracted: {len(self.pbarl_dims)} properties")
                for pid, d in self.pbarl_dims.items():
                    self.log(f"    PID {pid}: dim1={d['dim1']}, dim2={d['dim2']}")

            self.log(f"  Shells: {shell_count}, Bars: {bar_count}")
            self.log(f"  Centroids calculated: {len(self.element_centroids)}")
            self.log(f"\n  Total BDF models loaded: {len(self.bdf_models)}")

            self.bdf_status.config(text=f"✓ {len(self.bdf_models)} BDFs, {len(self.bdf_model.elements)} elements", foreground="green")

            if not self.output_folder.get():
                self.output_folder.set(os.path.dirname(path))

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.bdf_status.config(text="Error", foreground="red")

    def calculate_bar_skin_proximity(self):
        """Calculate which skin properties are near each bar property based on search distance."""
        import numpy as np

        search_dist = float(self.bar_skin_search_distance.get())
        self.log(f"\n  Calculating bar-skin proximity (distance={search_dist} mm)...")

        self.bar_to_nearby_skins = {}

        if not self.element_centroids:
            self.log("    WARNING: No element centroids available!")
            return

        # Get bar property -> bar element centroids
        bar_prop_centroids = {}  # bar_pid -> list of centroids
        for bar_eid in self.bar_elements:
            if bar_eid in self.element_centroids:
                bar_pid = self.elem_to_prop.get(bar_eid)
                if bar_pid and bar_pid in self.bar_properties:
                    if bar_pid not in bar_prop_centroids:
                        bar_prop_centroids[bar_pid] = []
                    bar_prop_centroids[bar_pid].append(self.element_centroids[bar_eid])

        # Get shell property -> shell element centroids
        shell_prop_centroids = {}  # skin_pid -> list of centroids
        for shell_eid in self.shell_elements:
            if shell_eid in self.element_centroids:
                shell_pid = self.elem_to_prop.get(shell_eid)
                if shell_pid and shell_pid in self.skin_properties:
                    if shell_pid not in shell_prop_centroids:
                        shell_prop_centroids[shell_pid] = []
                    shell_prop_centroids[shell_pid].append(self.element_centroids[shell_eid])

        # For each bar property, find nearby skin properties
        for bar_pid, bar_cents in bar_prop_centroids.items():
            nearby_skins = set()

            for skin_pid, skin_cents in shell_prop_centroids.items():
                # Check if any bar element is close to any shell element of this property
                is_nearby = False
                for bc in bar_cents:
                    for sc in skin_cents:
                        dist = np.sqrt((bc[0]-sc[0])**2 + (bc[1]-sc[1])**2 + (bc[2]-sc[2])**2)
                        if dist <= search_dist:
                            is_nearby = True
                            break
                    if is_nearby:
                        break

                if is_nearby:
                    nearby_skins.add(skin_pid)

            self.bar_to_nearby_skins[bar_pid] = nearby_skins

        # Create reverse mapping: skin -> nearby bars
        self.skin_to_nearby_bars = {}
        for bar_pid, skins in self.bar_to_nearby_skins.items():
            for skin_pid in skins:
                if skin_pid not in self.skin_to_nearby_bars:
                    self.skin_to_nearby_bars[skin_pid] = set()
                self.skin_to_nearby_bars[skin_pid].add(bar_pid)

        # Log summary
        total_connections = sum(len(skins) for skins in self.bar_to_nearby_skins.values())
        avg_skins = total_connections / len(self.bar_to_nearby_skins) if self.bar_to_nearby_skins else 0
        avg_bars = sum(len(bars) for bars in self.skin_to_nearby_bars.values()) / len(self.skin_to_nearby_bars) if self.skin_to_nearby_bars else 0
        self.log(f"    Bar properties: {len(self.bar_to_nearby_skins)}")
        self.log(f"    Skin properties with nearby bars: {len(self.skin_to_nearby_bars)}")
        self.log(f"    Total connections: {total_connections}")
        self.log(f"    Avg skins per bar: {avg_skins:.1f}")
        self.log(f"    Avg bars per skin: {avg_bars:.1f}")

    def load_properties(self):
        path = self.property_excel_path.get()
        if not path:
            messagebox.showerror("Error", "Select Property Excel")
            return

        self.log("\n" + "="*70)
        self.log("LOADING PROPERTIES")
        self.log("="*70)

        try:
            xl = pd.ExcelFile(path)
            self.log(f"Sheets: {xl.sheet_names}")

            bar_min = float(self.bar_min_thickness.get())
            skin_min = float(self.skin_min_thickness.get())

            self.bar_properties = {}
            self.skin_properties = {}
            self.current_bar_thicknesses = {}
            self.current_skin_thicknesses = {}
            self.gfem_bar_thicknesses = {}
            self.gfem_skin_thicknesses = {}

            for sheet in xl.sheet_names:
                sl = sheet.lower().replace('_', '').replace(' ', '')
                df = pd.read_excel(xl, sheet_name=sheet)

                if 'bar' in sl and 'prop' in sl:
                    self.log(f"\nReading bar properties from '{sheet}'...")
                    for _, row in df.iterrows():
                        pid = int(row.iloc[0]) if pd.notna(row.iloc[0]) else None
                        if pid:
                            gfem_dim1 = float(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else bar_min
                            self.bar_properties[pid] = {
                                'dim1': gfem_dim1,
                                'dim2': float(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else bar_min,
                            }
                            self.gfem_bar_thicknesses[pid] = gfem_dim1  # Store GFEM thickness
                            self.current_bar_thicknesses[pid] = bar_min
                    self.log(f"  Loaded {len(self.bar_properties)} bar properties")

                elif 'skin' in sl and 'prop' in sl:
                    self.log(f"\nReading skin properties from '{sheet}'...")
                    for _, row in df.iterrows():
                        pid = int(row.iloc[0]) if pd.notna(row.iloc[0]) else None
                        if pid:
                            gfem_thickness = float(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else skin_min
                            self.skin_properties[pid] = {
                                'thickness': gfem_thickness,
                            }
                            self.gfem_skin_thicknesses[pid] = gfem_thickness  # Store GFEM thickness
                            self.current_skin_thicknesses[pid] = skin_min
                    self.log(f"  Loaded {len(self.skin_properties)} skin properties")

            self.prop_status.config(text=f"✓ Bar: {len(self.bar_properties)}, Skin: {len(self.skin_properties)}", foreground="green")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.prop_status.config(text="Error", foreground="red")

    # ==================== RF CHECK v2.1 ALLOWABLE LOADING (EXACT COPY) ====================
    def rf_load_allowable(self):
        """Load Allowable stress data and fit power law - EXACT RF Check v2.1 logic."""
        path = self.allowable_excel_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Select a valid Allowable file")
            return

        self.log("\n" + "="*70)
        self.log("LOADING ALLOWABLE DATA & FITTING POWER LAW")
        self.log("(RF Check v2.1 exact logic)")
        self.log("="*70)

        try:
            r2_threshold = float(self.r2_threshold_var.get())
            min_data_pts = int(self.min_data_points_var.get())
        except:
            r2_threshold = 0.95
            min_data_pts = 3

        self.log(f"R² Threshold: {r2_threshold}, Min Data Points: {min_data_pts}")

        try:
            # Load file
            if path.endswith('.csv'):
                raw_df = pd.read_csv(path)
            else:
                xl = pd.ExcelFile(path)
                bar_sheet = None
                for sheet in xl.sheet_names:
                    if 'bar' in sheet.lower() or 'allowable' in sheet.lower() or 'summary' in sheet.lower():
                        bar_sheet = sheet
                        break
                raw_df = pd.read_excel(path, sheet_name=bar_sheet if bar_sheet else 0)

            self.log(f"\nLoaded: {len(raw_df)} rows")
            self.log(f"Columns: {list(raw_df.columns)}")

            # Clean column names
            clean_cols = {}
            for col in raw_df.columns:
                clean_name = str(col).replace('\n', ' ').replace('\r', ' ').strip()
                clean_name = ' '.join(clean_name.split())
                clean_cols[col] = clean_name
            raw_df = raw_df.rename(columns=clean_cols)

            # Map column names (RF Check v2.1 EXACT mapping)
            col_map = {}
            for col in raw_df.columns:
                col_up = col.upper().replace(' ', '_').replace('BAR_', '').replace('(MM)', '').strip('_')

                if col_up in ['PROPERTY_ID', 'PROPERTY', 'PROP_ID']:
                    col_map[col] = 'Property'
                elif col_up in ['ELEMENT_ID', 'ELEMENT']:
                    col_map[col] = 'Element_ID'
                elif col_up in ['ELEMENT_TYPE', 'ELEMENT_TYP']:
                    col_map[col] = 'Element_Type'
                elif col_up in ['T', 'THICKNESS', 'T_MM']:
                    col_map[col] = 'Thickness'
                elif col_up in ['ALLOWABLE', 'ALLOW', 'ALLOWABLE_STRESS']:
                    col_map[col] = 'Allowable'

            self.log(f"\nColumn mapping:")
            for orig, mapped in col_map.items():
                self.log(f"  '{orig}' -> '{mapped}'")

            df = raw_df.rename(columns=col_map)

            # Save full df for element-based fitting
            df_full_elements = df.copy() if 'Element_ID' in df.columns else None

            # Check for NEW format with Element_Type
            if 'Element_Type' in df.columns or 'Element_ID' in df.columns:
                self.log("\nDetected format with Element_Type/Element_ID")
                df = self._process_new_allowable_format(df)

            self.allowable_df = df

            # Convert to numeric
            df['Thickness'] = pd.to_numeric(df['Thickness'], errors='coerce')
            df['Allowable'] = pd.to_numeric(df['Allowable'], errors='coerce')
            df['Property'] = pd.to_numeric(df['Property'], errors='coerce')
            df = df.dropna(subset=['Property', 'Thickness', 'Allowable'])

            properties = df['Property'].unique()
            self.log(f"\nFitting {len(properties)} properties...")

            self.allowable_interp = {}
            excluded_r2 = []
            excluded_data = []
            valid_props = []

            for pid in properties:
                pid_int = int(pid)
                prop_data = df[df['Property'] == pid]
                n_pts = len(prop_data)

                if n_pts < min_data_pts:
                    avg = prop_data['Allowable'].mean()
                    self.allowable_interp[pid_int] = {'a': avg, 'b': 0, 'n_pts': n_pts, 'r2': 0, 'excluded': True}
                    excluded_data.append((pid_int, n_pts))
                    continue

                try:
                    x = prop_data['Thickness'].values.astype(float)
                    y = prop_data['Allowable'].values.astype(float)

                    valid = (x > 0) & (y > 0)
                    x, y = x[valid], y[valid]

                    if len(x) < 2:
                        self.allowable_interp[pid_int] = {'a': np.mean(y), 'b': 0, 'n_pts': len(x), 'r2': 0, 'excluded': True}
                        excluded_data.append((pid_int, len(x)))
                        continue

                    # Power law fit
                    log_x, log_y = np.log(x), np.log(y)
                    coeffs = np.polyfit(log_x, log_y, 1)
                    b, log_a = coeffs[0], coeffs[1]
                    a = np.exp(log_a)

                    # R²
                    y_pred = a * (x ** b)
                    ss_res = np.sum((y - y_pred) ** 2)
                    ss_tot = np.sum((y - np.mean(y)) ** 2)
                    r2 = 1 - (ss_res / ss_tot) if ss_tot > 0 else 0

                    if r2 < r2_threshold:
                        self.allowable_interp[pid_int] = {'a': np.mean(y), 'b': 0, 'n_pts': n_pts, 'r2': r2, 'excluded': True}
                        excluded_r2.append((pid_int, r2, n_pts))
                    else:
                        self.allowable_interp[pid_int] = {'a': a, 'b': b, 'n_pts': n_pts, 'r2': r2, 'excluded': False}
                        valid_props.append(pid_int)

                except Exception as e:
                    self.allowable_interp[pid_int] = {'a': 100, 'b': 0, 'n_pts': n_pts, 'r2': 0, 'excluded': True}
                    excluded_data.append((pid_int, n_pts))

            self.log(f"\n{'='*50}")
            self.log(f"Valid fits (R² >= {r2_threshold}): {len(valid_props)}")
            self.log(f"Excluded (R² < {r2_threshold}): {len(excluded_r2)}")
            self.log(f"Excluded (data < {min_data_pts}): {len(excluded_data)}")

            if valid_props:
                self.log(f"\nSample valid fits (Property):")
                for pid in valid_props[:5]:
                    p = self.allowable_interp[pid]
                    self.log(f"  Property {pid}: Allow = {p['a']:.4f} × T^({p['b']:.4f}), R²={p['r2']:.4f}")

            # ELEMENT-BASED CURVE FITTING
            self.allowable_elem_interp = {}

            if df_full_elements is not None and 'Element_ID' in df_full_elements.columns:
                self.log(f"\n{'='*50}")
                self.log("ELEMENT-BASED CURVE FITTING")
                self.log(f"{'='*50}")

                df_elem = df_full_elements.copy()
                df_elem['Thickness'] = pd.to_numeric(df_elem['Thickness'], errors='coerce')
                df_elem['Allowable'] = pd.to_numeric(df_elem['Allowable'], errors='coerce')
                df_elem['Element_ID'] = pd.to_numeric(df_elem['Element_ID'], errors='coerce')
                df_elem['Property'] = pd.to_numeric(df_elem['Property'], errors='coerce')
                df_elem = df_elem.dropna(subset=['Element_ID', 'Thickness', 'Allowable'])

                elements = df_elem['Element_ID'].unique()
                self.log(f"Fitting {len(elements)} elements...")

                valid_elems = []

                for elem_id in elements:
                    elem_int = int(elem_id)
                    elem_data = df_elem[df_elem['Element_ID'] == elem_id].copy()

                    elem_pid = elem_data['Property'].iloc[0] if len(elem_data) > 0 else None
                    elem_pid_int = int(elem_pid) if pd.notna(elem_pid) else None

                    # Filter by Element_Type if exists
                    if 'Element_Type' in elem_data.columns and elem_data['Element_Type'].notna().any():
                        elem_data_sorted = elem_data.sort_values('Allowable', ascending=True)
                        critical_elem_type = elem_data_sorted.iloc[0]['Element_Type']
                        filtered_data = elem_data[elem_data['Element_Type'] == critical_elem_type].copy()

                        fit_data = []
                        thickness_vals = pd.Series(filtered_data['Thickness'].values).dropna().unique()
                        for t in sorted(thickness_vals):
                            t_data = filtered_data[filtered_data['Thickness'] == t]
                            if len(t_data) > 0:
                                fit_data.append({'Thickness': float(t), 'Allowable': float(t_data['Allowable'].min())})
                        fit_df = pd.DataFrame(fit_data)
                    else:
                        fit_data = []
                        thickness_vals = pd.Series(elem_data['Thickness'].values).dropna().unique()
                        for t in sorted(thickness_vals):
                            t_data = elem_data[elem_data['Thickness'] == t]
                            if len(t_data) > 0:
                                fit_data.append({'Thickness': float(t), 'Allowable': float(t_data['Allowable'].min())})
                        fit_df = pd.DataFrame(fit_data)

                    n_pts = len(fit_df)

                    if n_pts < min_data_pts:
                        avg = fit_df['Allowable'].mean() if len(fit_df) > 0 else 0
                        self.allowable_elem_interp[elem_int] = {'a': avg, 'b': 0, 'n_pts': n_pts, 'r2': 0, 'excluded': True, 'property': elem_pid_int}
                        continue

                    try:
                        x = fit_df['Thickness'].values.astype(float)
                        y = fit_df['Allowable'].values.astype(float)
                        valid_mask = (x > 0) & (y > 0)
                        x, y = x[valid_mask], y[valid_mask]

                        if len(x) < 2:
                            self.allowable_elem_interp[elem_int] = {'a': np.mean(y), 'b': 0, 'n_pts': len(x), 'r2': 0, 'excluded': True, 'property': elem_pid_int}
                            continue

                        log_x, log_y = np.log(x), np.log(y)
                        coeffs = np.polyfit(log_x, log_y, 1)
                        b, log_a = coeffs[0], coeffs[1]
                        a = np.exp(log_a)

                        y_pred = a * (x ** b)
                        ss_res = np.sum((y - y_pred) ** 2)
                        ss_tot = np.sum((y - np.mean(y)) ** 2)
                        r2 = 1 - (ss_res / ss_tot) if ss_tot > 0 else 0

                        if r2 < r2_threshold:
                            self.allowable_elem_interp[elem_int] = {'a': np.mean(y), 'b': 0, 'n_pts': n_pts, 'r2': r2, 'excluded': True, 'property': elem_pid_int}
                        else:
                            self.allowable_elem_interp[elem_int] = {'a': a, 'b': b, 'n_pts': n_pts, 'r2': r2, 'excluded': False, 'property': elem_pid_int}
                            valid_elems.append(elem_int)

                    except:
                        self.allowable_elem_interp[elem_int] = {'a': 100, 'b': 0, 'n_pts': n_pts, 'r2': 0, 'excluded': True, 'property': elem_pid_int}

                # Count excluded elements
                excluded_elem_r2 = sum(1 for e in self.allowable_elem_interp.values() if e.get('excluded') and e.get('r2', 0) > 0)
                excluded_elem_data = sum(1 for e in self.allowable_elem_interp.values() if e.get('excluded') and e.get('n_pts', 0) < min_data_pts)

                self.log(f"Valid element fits (R² >= {r2_threshold}): {len(valid_elems)}")
                self.log(f"Excluded elements (R² < {r2_threshold}): {excluded_elem_r2}")
                self.log(f"Excluded elements (data < {min_data_pts}): {excluded_elem_data}")

                if valid_elems:
                    self.log(f"\nSample valid fits (Element):")
                    for eid in valid_elems[:5]:
                        e = self.allowable_elem_interp[eid]
                        self.log(f"  Element {eid}: Allow = {e['a']:.4f} × T^({e['b']:.4f}), R²={e['r2']:.4f}")

            # RF Check v2.1 format status
            n_elem_valid = len(valid_elems) if 'valid_elems' in dir() else 0
            n_elem_excl = len(self.allowable_elem_interp) - n_elem_valid if self.allowable_elem_interp else 0
            self.allow_status.config(text=f"✓ Prop: {len(valid_props)} valid | Elem: {n_elem_valid} valid, {n_elem_excl} excl", foreground="green")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.allow_status.config(text="Error", foreground="red")

    def _process_new_allowable_format(self, df):
        """Process format with Element_Type - RF Check v2.1 logic."""
        self.log("  Processing Element_Type format...")

        result_data = []

        # Get unique properties safely
        if 'Property' not in df.columns:
            self.log("  WARNING: No 'Property' column found")
            return pd.DataFrame(result_data)

        properties = pd.Series(df['Property'].values).dropna().unique()

        for pid in properties:
            prop_df = df[df['Property'] == pid].copy()
            if len(prop_df) == 0:
                continue

            prop_df_sorted = prop_df.sort_values('Allowable', ascending=True)
            critical_row = prop_df_sorted.iloc[0]

            crit_elem_id = critical_row.get('Element_ID', None)
            crit_elem_type = critical_row.get('Element_Type', None)

            if crit_elem_id is not None and crit_elem_type is not None:
                try:
                    crit_id = int(crit_elem_id)
                    mask = (prop_df['Element_ID'].astype(float).astype(int) == crit_id) & (prop_df['Element_Type'] == crit_elem_type)
                    filtered_df = prop_df[mask].copy()
                except:
                    filtered_df = prop_df.copy()
            elif crit_elem_type is not None:
                filtered_df = prop_df[prop_df['Element_Type'] == crit_elem_type].copy()
            else:
                filtered_df = prop_df.copy()

            # Get unique thicknesses safely (avoid DataFrame.unique() issue)
            if 'Thickness' not in filtered_df.columns or len(filtered_df) == 0:
                continue

            thickness_values = pd.Series(filtered_df['Thickness'].values).dropna().unique()

            for t in sorted(thickness_values):
                t_data = filtered_df[filtered_df['Thickness'] == t]
                if len(t_data) > 0:
                    min_allow = t_data['Allowable'].min()
                    result_data.append({'Property': int(pid), 'Thickness': float(t), 'Allowable': float(min_allow)})

        self.log(f"  Processed: {len(properties)} properties -> {len(result_data)} data points")
        return pd.DataFrame(result_data)

    def load_element_ids(self):
        path = self.element_excel_path.get()
        if not path:
            return

        self.log("\n" + "="*70)
        self.log("LOADING ELEMENT IDs FOR OFFSET")
        self.log("="*70)

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

            self.elem_status.config(text=f"✓ Landing: {len(self.landing_elem_ids)}, Bar: {len(self.bar_offset_elem_ids)}", foreground="green")

        except Exception as e:
            self.log(f"ERROR: {e}")
            self.elem_status.config(text="Error", foreground="red")

    def load_residual_strength(self):
        """Load Residual Strength Excel for stress combination - RF Check Tool logic."""
        path = self.residual_strength_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Select Residual Strength Excel file")
            return

        self.log("\n" + "="*70)
        self.log("LOADING RESIDUAL STRENGTH DATA")
        self.log("="*70)

        try:
            xl = pd.ExcelFile(path)
            self.log(f"Sheets: {xl.sheet_names}")

            # Find Residual Strength sheet (RF Check Tool logic)
            res_sh = None
            for sheet in xl.sheet_names:
                sl = sheet.lower().replace(' ', '').replace('_', '')
                if sl == 'residualstrength' or sl == 'residual strength':
                    res_sh = sheet
                    break

            # Fallback - search for sheets containing 'residual' or 'strength'
            if not res_sh:
                for sheet in xl.sheet_names:
                    sl = sheet.lower()
                    if 'residual' in sl or 'strength' in sl:
                        res_sh = sheet
                        break

            if not res_sh:
                # Use first sheet if no specific sheet found
                res_sh = xl.sheet_names[0]
                self.log(f"  No 'Residual Strength' sheet found, using: {res_sh}")
            else:
                self.log(f"  Found Residual Strength sheet: {res_sh}")

            self.residual_strength_df = pd.read_excel(xl, sheet_name=res_sh)
            self.log(f"  Loaded {len(self.residual_strength_df)} rows")
            self.log(f"  Columns: {list(self.residual_strength_df.columns)}")

            # Parse combination table structure (first col = Combined LC, then pairs of Case/Mult)
            cols = self.residual_strength_df.columns.tolist()
            comb_col = cols[0]
            self.log(f"  Combined LC column: {comb_col}")

            # Find Case ID + Multiplier column pairs
            self.combination_table = []
            i = 1
            while i < len(cols) - 1:
                col_name = str(cols[i]).upper()
                next_col_name = str(cols[i+1]).upper()
                if ('CASE' in col_name or 'ID' in col_name) and 'MULT' in next_col_name:
                    self.combination_table.append((cols[i], cols[i+1]))
                    self.log(f"    Found pair: {cols[i]} + {cols[i+1]}")
                    i += 2
                else:
                    i += 1

            self.log(f"  Total combination pairs: {len(self.combination_table)}")
            self.resid_status.config(text=f"✓ {len(self.residual_strength_df)} rows, {len(self.combination_table)} pairs", foreground="green")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.resid_status.config(text="Error", foreground="red")

    # ==================== HELPER FUNCTIONS ====================
    def get_allowable_stress(self, pid, thickness):
        pid_int = int(pid)
        if pid_int not in self.allowable_interp:
            return None
        params = self.allowable_interp[pid_int]
        if params.get('excluded', False) and params['b'] == 0:
            return params['a']
        return params['a'] * (thickness ** params['b'])

    def get_allowable_stress_elem(self, elem_id, thickness):
        elem_int = int(elem_id)
        if elem_int not in self.allowable_elem_interp:
            return None
        params = self.allowable_elem_interp[elem_int]
        if params.get('excluded', False) and params['b'] == 0:
            return params['a']
        return params['a'] * (thickness ** params['b'])

    def get_required_thickness(self, pid, stress, min_rf=1.0):
        pid_int = int(pid)
        if pid_int not in self.allowable_interp:
            return None
        params = self.allowable_interp[pid_int]
        a, b = params['a'], params['b']
        if params.get('excluded', False) or b == 0:
            return None
        required_allow = abs(stress) * min_rf
        if a <= 0:
            return None
        try:
            ratio = required_allow / a
            if ratio <= 0:
                return None
            t_req = ratio ** (1.0 / b)
            return t_req if 0 < t_req < 1000 else None
        except:
            return None

    def get_density(self, pid):
        """Get density from property's material (MAT1/MAT8/MAT9/etc.)"""
        if pid in self.prop_to_material:
            mid = self.prop_to_material[pid]
            if mid in self.material_densities:
                return self.material_densities[mid]
        return 2.7e-9  # Default aluminum

    # ==================== OPTIMIZATION ====================
    def start_optimization(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one BDF file")
            return
        if not self.bdf_model:
            self.load_bdf()
        if not self.bdf_model:
            return
        if not self.allowable_interp:
            messagebox.showerror("Error", "Load allowable data first")
            return

        # Validate constant thickness options
        if self.keep_bars_constant.get() and self.keep_skins_constant.get():
            messagebox.showerror("Error", "Cannot keep both Bar and Skin thicknesses constant.\n"
                                "Please select only one or none.")
            return
        if self.keep_bars_constant.get() and not self.gfem_bar_thicknesses:
            messagebox.showerror("Error", "Keep Bar constant requires Excel thickness data.\n"
                                "Please load Excel data with bar thicknesses first.")
            return
        if self.keep_skins_constant.get() and not self.gfem_skin_thicknesses:
            messagebox.showerror("Error", "Keep Skin constant requires Excel thickness data.\n"
                                "Please load Excel data with skin thicknesses first.")
            return

        self.is_running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.iteration_results = []
        self.best_solution = None

        # Run selected algorithm
        algo = self.algorithm_var.get()
        if algo == 'Decoupled Min Weight':
            threading.Thread(target=self._run_decoupled_min_weight, daemon=True).start()
        elif algo == 'Coupled Efficiency Analysis':
            threading.Thread(target=self._run_coupled_efficiency_analysis, daemon=True).start()
        else:
            # Default: Bottom-Up (Min to Target)
            threading.Thread(target=self._run_bottomup_algorithm, daemon=True).start()

    def stop_optimization(self):
        self.is_running = False
        self.log("\n*** STOPPING ***")

    def _run_simple_iterative(self):
        """Algorithm 1: Simple Iterative - increase failing, decrease over-designed."""
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: SIMPLE ITERATIVE")
            self.log("="*70)

            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            max_iter = int(self.max_iterations.get())

            self.log(f"\nBar range: {bar_min}-{bar_max} mm")
            self.log(f"Skin range: {skin_min}-{skin_max} mm")
            self.log(f"Target RF: {target_rf} ± {rf_tol}")
            self.log(f"Max iterations: {max_iter}")

            # Initialize all properties to minimum
            for pid in self.bar_properties:
                self.current_bar_thicknesses[pid] = bar_min
            for pid in self.skin_properties:
                self.current_skin_thicknesses[pid] = skin_min

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"opt_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            iteration = 0
            best_weight = float('inf')

            while iteration < max_iter and self.is_running:
                iteration += 1
                self.log(f"\n{'='*60}")
                self.log(f"ITERATION {iteration}")
                self.log(f"{'='*60}")

                self.update_progress((iteration/max_iter)*100, f"Iteration {iteration}/{max_iter}")

                iter_folder = os.path.join(base_folder, f"iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                # Run iteration
                result = self._run_iteration(iter_folder, iteration, target_rf)

                if not result:
                    self.log("  Iteration failed, increasing all thicknesses...")
                    for pid in self.current_bar_thicknesses:
                        self.current_bar_thicknesses[pid] = min(self.current_bar_thicknesses[pid] + step, bar_max)
                    for pid in self.current_skin_thicknesses:
                        self.current_skin_thicknesses[pid] = min(self.current_skin_thicknesses[pid] + step, skin_max)
                    continue

                self.iteration_results.append(result)

                # Check for best
                if result['min_rf'] >= target_rf - rf_tol and result['weight'] < best_weight:
                    best_weight = result['weight']
                    self.best_solution = result.copy()
                    self.log(f"\n  *** NEW BEST: Weight={best_weight:.6f}t, RF={result['min_rf']:.4f} ***")

                # Per-property update based on RF
                self._smart_thickness_update(result, step, bar_min, bar_max, skin_min, skin_max, target_rf, rf_tol)

            # Summary
            self.log("\n" + "="*70)
            self.log("OPTIMIZATION COMPLETE")
            self.log("="*70)

            if self.best_solution:
                self.log(f"\nBest: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}")
                self._update_ui()

            self._save_results(base_folder)
            self._generate_final_report(base_folder, "Simple Iterative", nastran_count=iteration)

            self.root.after(0, lambda: messagebox.showinfo("Done", f"Complete!\nIterations: {iteration}\nBest weight: {best_weight:.6f}t"))

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    # ==================== FAST GA (SURROGATE MODEL) ====================
    def _run_fast_ga(self):
        """Algorithm 2: Fast GA with surrogate model (no Nastran during optimization)."""
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: FAST GA (SURROGATE MODEL)")
            self.log("="*70)

            # Parameters
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())

            pop_size = int(self.ga_population.get())
            n_generations = int(self.ga_generations.get())
            mutation_rate = float(self.ga_mutation_rate.get())
            crossover_rate = float(self.ga_crossover_rate.get())

            self.log(f"\nGA Parameters:")
            self.log(f"  Population: {pop_size}")
            self.log(f"  Generations: {n_generations}")
            self.log(f"  Mutation Rate: {mutation_rate}")
            self.log(f"  Crossover Rate: {crossover_rate}")
            self.log(f"\nTarget RF: {target_rf} ± {rf_tol}")

            # Create chromosome structure
            bar_pids = list(self.bar_properties.keys())
            skin_pids = list(self.skin_properties.keys())
            n_bars = len(bar_pids)
            n_skins = len(skin_pids)
            n_genes = n_bars + n_skins

            self.log(f"\nChromosome: {n_bars} bar + {n_skins} skin = {n_genes} genes")

            if n_genes == 0:
                self.log("ERROR: No properties to optimize!")
                return

            # Initialize reference stresses (using average from allowable data)
            self._init_reference_stresses(bar_pids, skin_pids, bar_min, skin_min)

            # Initialize population
            population = []
            for _ in range(pop_size):
                chromosome = []
                # Bar thicknesses
                for pid in bar_pids:
                    chromosome.append(random.uniform(bar_min, bar_max))
                # Skin thicknesses
                for pid in skin_pids:
                    chromosome.append(random.uniform(skin_min, skin_max))
                population.append(chromosome)

            # Evolution
            best_fitness = float('inf')
            best_chromosome = None
            fitness_history = []

            for gen in range(n_generations):
                if not self.is_running:
                    break

                # Evaluate fitness
                fitness_values = []
                for chromosome in population:
                    fit = self._evaluate_surrogate_fitness(
                        chromosome, bar_pids, skin_pids,
                        bar_min, bar_max, skin_min, skin_max,
                        target_rf, rf_tol
                    )
                    fitness_values.append(fit)

                # Find best
                min_fit_idx = fitness_values.index(min(fitness_values))
                if fitness_values[min_fit_idx] < best_fitness:
                    best_fitness = fitness_values[min_fit_idx]
                    best_chromosome = population[min_fit_idx].copy()

                fitness_history.append(best_fitness)

                # Progress update
                if gen % 10 == 0 or gen == n_generations - 1:
                    self.update_progress((gen / n_generations) * 100, f"Generation {gen}/{n_generations}")
                    self.log(f"Gen {gen}: Best Fitness = {best_fitness:.6f}")

                # Selection (Tournament)
                new_population = []
                while len(new_population) < pop_size:
                    # Tournament selection
                    idx1, idx2 = random.sample(range(pop_size), 2)
                    parent1 = population[idx1] if fitness_values[idx1] < fitness_values[idx2] else population[idx2]

                    idx3, idx4 = random.sample(range(pop_size), 2)
                    parent2 = population[idx3] if fitness_values[idx3] < fitness_values[idx4] else population[idx4]

                    # Crossover (BLX-alpha)
                    if random.random() < crossover_rate:
                        child = self._blx_crossover(parent1, parent2, alpha=0.5)
                    else:
                        child = parent1.copy()

                    # Mutation (Gaussian)
                    child = self._gaussian_mutation(
                        child, mutation_rate, n_bars,
                        bar_min, bar_max, skin_min, skin_max
                    )

                    new_population.append(child)

                population = new_population

            # Apply best solution
            if best_chromosome:
                for i, pid in enumerate(bar_pids):
                    self.current_bar_thicknesses[pid] = best_chromosome[i]
                for i, pid in enumerate(skin_pids):
                    self.current_skin_thicknesses[pid] = best_chromosome[n_bars + i]

                # Calculate final weight and RF
                weight = self._calculate_weight()
                min_rf = self._estimate_min_rf(best_chromosome, bar_pids, skin_pids)

                self.best_solution = {
                    'iteration': n_generations,
                    'min_rf': min_rf,
                    'weight': weight,
                    'bar_thicknesses': self.current_bar_thicknesses.copy(),
                    'skin_thicknesses': self.current_skin_thicknesses.copy(),
                    'n_fail': 0 if min_rf >= target_rf else 1
                }

                self.log("\n" + "="*70)
                self.log("FAST GA COMPLETE")
                self.log("="*70)
                self.log(f"\nBest Solution:")
                self.log(f"  Weight: {weight:.6f} tonnes")
                self.log(f"  Estimated Min RF: {min_rf:.4f}")
                self.log(f"  Fitness: {best_fitness:.6f}")

                self._update_ui()

                # Save results
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_folder = os.path.join(self.output_folder.get(), f"fast_ga_{timestamp}")
                os.makedirs(base_folder, exist_ok=True)
                self._save_results(base_folder)
                self._generate_final_report(base_folder, "Fast GA (Surrogate)",
                    extra_info={"Note": "Results based on surrogate model - run Nastran to verify"})

                self.root.after(0, lambda: messagebox.showinfo(
                    "Fast GA Complete",
                    f"Optimization finished!\n\nWeight: {weight:.6f}t\nEstimated RF: {min_rf:.4f}\n\nNote: Run Nastran to verify results."
                ))

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _init_reference_stresses(self, bar_pids, skin_pids, bar_ref_t, skin_ref_t):
        """Initialize reference stresses for surrogate model."""
        self.reference_stresses = {}
        self.reference_thickness = {}

        # For bars - estimate stress from allowable curve (assuming RF=1 at some thickness)
        for pid in bar_pids:
            if pid in self.allowable_interp:
                params = self.allowable_interp[pid]
                # Reference: at min thickness, stress ≈ allowable (RF=1)
                ref_allow = params['a'] * (bar_ref_t ** params['b'])
                self.reference_stresses[pid] = ref_allow  # Stress = Allowable when RF=1
                self.reference_thickness[pid] = bar_ref_t
            else:
                self.reference_stresses[pid] = 100.0  # Default
                self.reference_thickness[pid] = bar_ref_t

        # For skins
        for pid in skin_pids:
            if pid in self.allowable_interp:
                params = self.allowable_interp[pid]
                ref_allow = params['a'] * (skin_ref_t ** params['b'])
                self.reference_stresses[pid] = ref_allow
                self.reference_thickness[pid] = skin_ref_t
            else:
                self.reference_stresses[pid] = 100.0
                self.reference_thickness[pid] = skin_ref_t

    def _evaluate_surrogate_fitness(self, chromosome, bar_pids, skin_pids,
                                     bar_min, bar_max, skin_min, skin_max,
                                     target_rf, rf_tol):
        """Evaluate fitness using surrogate model (no Nastran)."""
        n_bars = len(bar_pids)

        # Calculate weight
        weight = 0.0

        # Bar weight
        for i, pid in enumerate(bar_pids):
            dim1 = chromosome[i]  # Optimized dimension
            # Use BDF PBARL dim2 (original geometry), fallback to Excel
            if pid in self.pbarl_dims:
                dim2 = self.pbarl_dims[pid]['dim2']
            else:
                dim2 = self.bar_properties[pid].get('dim2', dim1)
            rho = self.get_density(pid)
            if pid in self.prop_elements:
                length = sum(self.bar_lengths.get(eid, 0) for eid in self.prop_elements[pid])
                weight += length * dim1 * dim2 * rho

        # Skin weight
        for i, pid in enumerate(skin_pids):
            t = chromosome[n_bars + i]
            rho = self.get_density(pid)
            if pid in self.prop_elements:
                area = sum(self.element_areas.get(eid, 0) for eid in self.prop_elements[pid])
                weight += area * t * rho

        # Calculate RF penalty using surrogate model
        penalty = 0.0
        penalty_factor = 1000.0  # Large penalty for constraint violation

        # Bar RF
        for i, pid in enumerate(bar_pids):
            t = chromosome[i]
            rf = self._estimate_rf_surrogate(pid, t, is_bar=True)
            if rf < target_rf:
                penalty += penalty_factor * (target_rf - rf) ** 2

        # Skin RF
        for i, pid in enumerate(skin_pids):
            t = chromosome[n_bars + i]
            rf = self._estimate_rf_surrogate(pid, t, is_bar=False)
            if rf < target_rf:
                penalty += penalty_factor * (target_rf - rf) ** 2

        return weight + penalty

    def _estimate_rf_surrogate(self, pid, thickness, is_bar=True):
        """Estimate RF using surrogate model: Stress scales with thickness."""
        if pid not in self.allowable_interp:
            return 1.0  # Default pass

        params = self.allowable_interp[pid]
        allowable = params['a'] * (thickness ** params['b'])

        # Stress scaling: Stress ∝ 1/t^α (α=1.0 for bars, 1.5 for skins)
        alpha = 1.0 if is_bar else 1.5
        ref_stress = self.reference_stresses.get(pid, 100.0)
        ref_t = self.reference_thickness.get(pid, thickness)

        if ref_t > 0 and thickness > 0:
            stress = ref_stress * (ref_t / thickness) ** alpha
        else:
            stress = ref_stress

        if stress > 0:
            return allowable / stress
        return 999.0

    def _estimate_min_rf(self, chromosome, bar_pids, skin_pids):
        """Estimate minimum RF from chromosome."""
        min_rf = 999.0
        n_bars = len(bar_pids)

        for i, pid in enumerate(bar_pids):
            rf = self._estimate_rf_surrogate(pid, chromosome[i], is_bar=True)
            min_rf = min(min_rf, rf)

        for i, pid in enumerate(skin_pids):
            rf = self._estimate_rf_surrogate(pid, chromosome[n_bars + i], is_bar=False)
            min_rf = min(min_rf, rf)

        return min_rf

    def _blx_crossover(self, parent1, parent2, alpha=0.5):
        """BLX-alpha crossover for real-valued chromosomes with bounds check."""
        child = []
        for g1, g2 in zip(parent1, parent2):
            min_g, max_g = min(g1, g2), max(g1, g2)
            range_g = max_g - min_g
            # Generate child gene
            new_val = random.uniform(min_g - alpha * range_g, max_g + alpha * range_g)
            # Ensure positive (will be bounded later in mutation)
            child.append(max(0.1, new_val))  # Never negative!
        return child

    def _gaussian_mutation(self, chromosome, mutation_rate, n_bars,
                           bar_min, bar_max, skin_min, skin_max):
        """Gaussian mutation with STRICT bounds enforcement."""
        mutated = chromosome.copy()
        for i in range(len(mutated)):
            if random.random() < mutation_rate:
                if i < n_bars:
                    # Bar gene
                    sigma = (bar_max - bar_min) * 0.1
                    mutated[i] += random.gauss(0, sigma)
                    # STRICT bounds - never negative!
                    mutated[i] = max(bar_min, min(bar_max, mutated[i]))
                else:
                    # Skin gene
                    sigma = (skin_max - skin_min) * 0.1
                    mutated[i] += random.gauss(0, sigma)
                    # STRICT bounds - never negative!
                    mutated[i] = max(skin_min, min(skin_max, mutated[i]))
            else:
                # Even without mutation, enforce bounds (for crossover results)
                if i < n_bars:
                    mutated[i] = max(bar_min, min(bar_max, mutated[i]))
                else:
                    mutated[i] = max(skin_min, min(skin_max, mutated[i]))
        return mutated

    # ==================== HYBRID GA + NASTRAN ====================
    def _run_hybrid_ga(self):
        """Algorithm 3: Hybrid GA - Fast GA first, then Nastran validation."""
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: HYBRID GA + NASTRAN")
            self.log("="*70)

            # Phase 1: Run Fast GA
            self.log("\n>>> PHASE 1: Fast GA Optimization <<<")
            self._run_fast_ga_internal()

            if not self.is_running:
                return

            # Phase 2: Nastran validation of best solutions
            self.log("\n>>> PHASE 2: Nastran Validation <<<")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"hybrid_ga_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            target_rf = float(self.target_rf.get())

            # Run Nastran with best solution
            self.log("\nValidating best solution with Nastran...")
            iter_folder = os.path.join(base_folder, "validation")
            os.makedirs(iter_folder, exist_ok=True)

            result = self._run_iteration(iter_folder, 1, target_rf)

            if result:
                self.best_solution = result
                self.log(f"\nNastran Validation Results:")
                self.log(f"  Actual Min RF: {result['min_rf']:.4f}")
                self.log(f"  Actual Weight: {result['weight']:.6f}t")
                self.log(f"  Failures: {result['n_fail']}")

                self._update_ui()
                self._save_results(base_folder)
                self._generate_final_report(base_folder, "Hybrid GA + Nastran", nastran_count=1,
                    extra_info={"Phase 1": "Fast GA (Surrogate)", "Phase 2": "Nastran Validation"})

                self.root.after(0, lambda: messagebox.showinfo(
                    "Hybrid GA Complete",
                    f"Optimization finished!\n\nWeight: {result['weight']:.6f}t\nActual RF: {result['min_rf']:.4f}\nFailures: {result['n_fail']}"
                ))
            else:
                self.log("ERROR: Nastran validation failed!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_full_ga_nastran(self):
        """Algorithm 4: Full GA with Nastran - runs Nastran for every fitness evaluation."""
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: FULL GA + NASTRAN")
            self.log("(Runs Nastran for every fitness evaluation - SLOW but ACCURATE)")
            self.log("="*70)

            # Parameters
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())

            pop_size = int(self.ga_population.get())
            n_generations = int(self.ga_generations.get())
            mutation_rate = float(self.ga_mutation_rate.get())
            crossover_rate = float(self.ga_crossover_rate.get())

            self.log(f"\nGA Parameters:")
            self.log(f"  Population: {pop_size}")
            self.log(f"  Generations: {n_generations}")
            self.log(f"  Mutation Rate: {mutation_rate}")
            self.log(f"  Crossover Rate: {crossover_rate}")
            self.log(f"\nTarget RF: {target_rf} ± {rf_tol}")
            self.log(f"\nWARNING: This will run {pop_size * n_generations} Nastran analyses!")

            bar_pids = list(self.bar_properties.keys())
            skin_pids = list(self.skin_properties.keys())
            n_bars = len(bar_pids)
            n_skins = len(skin_pids)
            n_genes = n_bars + n_skins

            if n_genes == 0:
                self.log("ERROR: No properties to optimize!")
                return

            self.log(f"\nChromosome: {n_bars} bars + {n_skins} skins = {n_genes} genes")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"full_ga_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # Initialize population
            population = []
            for _ in range(pop_size):
                chromosome = []
                for pid in bar_pids:
                    chromosome.append(random.uniform(bar_min, bar_max))
                for pid in skin_pids:
                    chromosome.append(random.uniform(skin_min, skin_max))
                population.append(chromosome)

            best_fitness = float('-inf')
            best_chromosome = None
            best_result = None
            eval_count = 0

            for gen in range(n_generations):
                if not self.is_running:
                    break

                self.log(f"\n{'='*50}")
                self.log(f"GENERATION {gen+1}/{n_generations}")
                self.log(f"{'='*50}")

                self.update_progress((gen/n_generations)*100, f"Gen {gen+1}/{n_generations}")

                # Evaluate fitness for each individual using Nastran
                fitness_scores = []
                for idx, chromosome in enumerate(population):
                    if not self.is_running:
                        break

                    eval_count += 1
                    self.log(f"\n  Individual {idx+1}/{pop_size} (Eval #{eval_count})")

                    # Set thicknesses from chromosome
                    for i, pid in enumerate(bar_pids):
                        self.current_bar_thicknesses[pid] = max(bar_min, min(bar_max, chromosome[i]))
                    for i, pid in enumerate(skin_pids):
                        self.current_skin_thicknesses[pid] = max(skin_min, min(skin_max, chromosome[n_bars + i]))

                    # Run Nastran
                    iter_folder = os.path.join(base_folder, f"gen_{gen+1:03d}_ind_{idx+1:03d}")
                    os.makedirs(iter_folder, exist_ok=True)

                    result = self._run_iteration(iter_folder, eval_count, target_rf)

                    if result:
                        min_rf = result['min_rf']
                        weight = result['weight']

                        # Fitness: minimize weight, penalize RF < target
                        if min_rf >= target_rf - rf_tol:
                            fitness = 1000.0 / (weight + 0.001)  # Higher is better
                        else:
                            # Penalty for failing RF
                            penalty = (target_rf - min_rf) * 100
                            fitness = 1000.0 / (weight + 0.001) - penalty

                        fitness_scores.append(fitness)
                        self.log(f"    RF={min_rf:.4f}, Weight={weight:.6f}t, Fitness={fitness:.2f}")

                        # Track best
                        if fitness > best_fitness:
                            best_fitness = fitness
                            best_chromosome = chromosome.copy()
                            best_result = result
                            self.best_solution = result
                            self.log(f"    *** NEW BEST! ***")
                    else:
                        fitness_scores.append(-1000)  # Failed evaluation
                        self.log(f"    FAILED!")

                if not self.is_running or len(fitness_scores) < pop_size:
                    break

                # Selection (Tournament)
                def tournament_select(pop, fit, k=3):
                    selected = random.sample(list(zip(pop, fit)), k)
                    return max(selected, key=lambda x: x[1])[0]

                # Create new population
                new_population = []

                # Elitism - keep best individual
                if best_chromosome:
                    new_population.append(best_chromosome.copy())

                while len(new_population) < pop_size:
                    # Select parents
                    parent1 = tournament_select(population, fitness_scores)
                    parent2 = tournament_select(population, fitness_scores)

                    # Crossover (BLX-α)
                    if random.random() < crossover_rate:
                        child = []
                        alpha = 0.5
                        for g1, g2 in zip(parent1, parent2):
                            d = abs(g2 - g1)
                            low = max(bar_min if len(child) < n_bars else skin_min, min(g1, g2) - alpha * d)
                            high = min(bar_max if len(child) < n_bars else skin_max, max(g1, g2) + alpha * d)
                            child.append(random.uniform(low, high))
                    else:
                        child = parent1.copy()

                    # Mutation (Gaussian)
                    for i in range(len(child)):
                        if random.random() < mutation_rate:
                            if i < n_bars:
                                sigma = (bar_max - bar_min) * 0.1
                                child[i] = max(bar_min, min(bar_max, child[i] + random.gauss(0, sigma)))
                            else:
                                sigma = (skin_max - skin_min) * 0.1
                                child[i] = max(skin_min, min(skin_max, child[i] + random.gauss(0, sigma)))

                    new_population.append(child)

                population = new_population[:pop_size]

                # Log generation summary
                valid_fitness = [f for f in fitness_scores if f > -500]
                if valid_fitness:
                    self.log(f"\n  Gen {gen+1} Summary: Best Fitness={max(valid_fitness):.2f}, Avg={sum(valid_fitness)/len(valid_fitness):.2f}")

            # Final results
            self.log("\n" + "="*70)
            self.log("FULL GA + NASTRAN COMPLETE")
            self.log("="*70)
            self.log(f"Total Nastran evaluations: {eval_count}")

            if best_result:
                # Set final thicknesses
                for i, pid in enumerate(bar_pids):
                    self.current_bar_thicknesses[pid] = best_chromosome[i]
                for i, pid in enumerate(skin_pids):
                    self.current_skin_thicknesses[pid] = best_chromosome[n_bars + i]

                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {best_result['min_rf']:.4f}")
                self.log(f"  Weight: {best_result['weight']:.6f}t")
                self.log(f"  Failures: {best_result['n_fail']}")

                self._update_ui()
                self._save_results(base_folder)
                self._generate_final_report(base_folder, "Full GA + Nastran", nastran_count=eval_count,
                    extra_info={"Population": pop_size, "Generations": n_generations})

                self.root.after(0, lambda: messagebox.showinfo(
                    "Full GA Complete",
                    f"Optimization finished!\n\nTotal Evaluations: {eval_count}\nWeight: {best_result['weight']:.6f}t\nMin RF: {best_result['min_rf']:.4f}\nFailures: {best_result['n_fail']}"
                ))
            else:
                self.log("ERROR: No valid solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_surrogate_assisted_ga(self):
        """Algorithm 5: Surrogate-Assisted GA - learns from previous Nastran runs."""
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: SURROGATE-ASSISTED GA")
            self.log("(Learns from Nastran results - SMART & EFFICIENT)")
            self.log("="*70)

            # Parameters
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())

            pop_size = int(self.ga_population.get())
            n_generations = int(self.ga_generations.get())
            mutation_rate = float(self.ga_mutation_rate.get())
            crossover_rate = float(self.ga_crossover_rate.get())

            # Surrogate parameters
            initial_samples = min(pop_size, 20)  # Initial Nastran runs to build surrogate
            nastran_per_gen = max(3, pop_size // 10)  # Nastran runs per generation for update

            self.log(f"\nGA Parameters:")
            self.log(f"  Population: {pop_size}")
            self.log(f"  Generations: {n_generations}")
            self.log(f"  Mutation Rate: {mutation_rate}")
            self.log(f"  Crossover Rate: {crossover_rate}")
            self.log(f"\nSurrogate Parameters:")
            self.log(f"  Initial Nastran samples: {initial_samples}")
            self.log(f"  Nastran updates per generation: {nastran_per_gen}")
            self.log(f"\nEstimated total Nastran runs: {initial_samples + n_generations * nastran_per_gen}")
            self.log(f"(vs Full GA: {pop_size * n_generations})")

            bar_pids = list(self.bar_properties.keys())
            skin_pids = list(self.skin_properties.keys())
            n_bars = len(bar_pids)
            n_skins = len(skin_pids)
            n_genes = n_bars + n_skins

            if n_genes == 0:
                self.log("ERROR: No properties to optimize!")
                return

            self.log(f"\nChromosome: {n_bars} bars + {n_skins} skins = {n_genes} genes")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"surrogate_ga_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== PHASE 1: Initial Sampling (Latin Hypercube-like) ==========
            self.log("\n" + "="*50)
            self.log("PHASE 1: Building Initial Surrogate Model")
            self.log("="*50)

            # Storage for surrogate training data
            X_train = []  # Chromosomes (inputs)
            y_rf_train = []  # Min RF values (outputs)
            y_weight_train = []  # Weight values (outputs)

            # Generate initial samples using stratified random sampling
            initial_population = []
            for i in range(initial_samples):
                chromosome = []
                for j, pid in enumerate(bar_pids):
                    # Stratified: divide range into initial_samples parts
                    low = bar_min + (bar_max - bar_min) * (i / initial_samples)
                    high = bar_min + (bar_max - bar_min) * ((i + 1) / initial_samples)
                    chromosome.append(random.uniform(low, high))
                for j, pid in enumerate(skin_pids):
                    low = skin_min + (skin_max - skin_min) * (i / initial_samples)
                    high = skin_min + (skin_max - skin_min) * ((i + 1) / initial_samples)
                    chromosome.append(random.uniform(low, high))
                # Shuffle to avoid correlation
                random.shuffle(chromosome[:n_bars])
                if n_skins > 0:
                    random.shuffle(chromosome[n_bars:])
                initial_population.append(chromosome)

            # Run Nastran for initial samples
            nastran_count = 0
            best_fitness = float('-inf')
            best_chromosome = None
            best_result = None

            for idx, chromosome in enumerate(initial_population):
                if not self.is_running:
                    break

                nastran_count += 1
                self.log(f"\n  Initial Sample {idx+1}/{initial_samples} (Nastran #{nastran_count})")
                self.update_progress((idx/initial_samples)*20, f"Initial sampling {idx+1}/{initial_samples}")

                # Set thicknesses
                for i, pid in enumerate(bar_pids):
                    self.current_bar_thicknesses[pid] = chromosome[i]
                for i, pid in enumerate(skin_pids):
                    self.current_skin_thicknesses[pid] = chromosome[n_bars + i]

                # Run Nastran
                iter_folder = os.path.join(base_folder, f"init_{idx+1:03d}")
                os.makedirs(iter_folder, exist_ok=True)
                result = self._run_iteration(iter_folder, nastran_count, target_rf)

                if result:
                    min_rf = result['min_rf']
                    weight = result['weight']

                    X_train.append(chromosome)
                    y_rf_train.append(min_rf)
                    y_weight_train.append(weight)

                    # Calculate fitness
                    if min_rf >= target_rf - rf_tol:
                        fitness = 1000.0 / (weight + 0.001)
                    else:
                        penalty = (target_rf - min_rf) * 100
                        fitness = 1000.0 / (weight + 0.001) - penalty

                    self.log(f"    RF={min_rf:.4f}, Weight={weight:.6f}t, Fitness={fitness:.2f}")

                    if fitness > best_fitness:
                        best_fitness = fitness
                        best_chromosome = chromosome.copy()
                        best_result = result
                        self.best_solution = result
                        self.log(f"    *** NEW BEST! ***")

            if len(X_train) < 3:
                self.log("ERROR: Not enough initial samples!")
                return

            self.log(f"\n  Surrogate built with {len(X_train)} samples")

            # ========== PHASE 2: GA with Surrogate ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: Surrogate-Assisted Optimization")
            self.log("="*50)

            # Initialize population (include best from initial + random)
            population = []
            if best_chromosome:
                population.append(best_chromosome.copy())

            while len(population) < pop_size:
                chromosome = []
                for pid in bar_pids:
                    chromosome.append(random.uniform(bar_min, bar_max))
                for pid in skin_pids:
                    chromosome.append(random.uniform(skin_min, skin_max))
                population.append(chromosome)

            for gen in range(n_generations):
                if not self.is_running:
                    break

                self.log(f"\n{'='*50}")
                self.log(f"GENERATION {gen+1}/{n_generations}")
                self.log(f"{'='*50}")

                progress = 20 + (gen / n_generations) * 80
                self.update_progress(progress, f"Gen {gen+1}/{n_generations}")

                # Evaluate all individuals using SURROGATE
                fitness_scores = []
                surrogate_predictions = []

                for idx, chromosome in enumerate(population):
                    # Surrogate prediction using weighted k-nearest neighbors
                    pred_rf, pred_weight, confidence = self._surrogate_predict(
                        chromosome, X_train, y_rf_train, y_weight_train,
                        bar_min, bar_max, skin_min, skin_max, n_bars
                    )

                    # Calculate predicted fitness
                    if pred_rf >= target_rf - rf_tol:
                        pred_fitness = 1000.0 / (pred_weight + 0.001)
                    else:
                        penalty = (target_rf - pred_rf) * 100
                        pred_fitness = 1000.0 / (pred_weight + 0.001) - penalty

                    fitness_scores.append(pred_fitness)
                    surrogate_predictions.append({
                        'chromosome': chromosome,
                        'pred_rf': pred_rf,
                        'pred_weight': pred_weight,
                        'pred_fitness': pred_fitness,
                        'confidence': confidence
                    })

                # Select top candidates for actual Nastran evaluation
                # Criteria: high predicted fitness OR low confidence (exploration)
                sorted_by_fitness = sorted(surrogate_predictions, key=lambda x: x['pred_fitness'], reverse=True)
                sorted_by_uncertainty = sorted(surrogate_predictions, key=lambda x: x['confidence'])

                candidates_for_nastran = []
                # Top performers (exploitation)
                for item in sorted_by_fitness[:nastran_per_gen // 2]:
                    if item['chromosome'] not in [c['chromosome'] for c in candidates_for_nastran]:
                        candidates_for_nastran.append(item)
                # Most uncertain (exploration)
                for item in sorted_by_uncertainty[:nastran_per_gen // 2 + 1]:
                    if item['chromosome'] not in [c['chromosome'] for c in candidates_for_nastran]:
                        candidates_for_nastran.append(item)

                candidates_for_nastran = candidates_for_nastran[:nastran_per_gen]

                # Run Nastran for selected candidates
                self.log(f"\n  Running Nastran for {len(candidates_for_nastran)} candidates...")

                for cand in candidates_for_nastran:
                    if not self.is_running:
                        break

                    chromosome = cand['chromosome']
                    nastran_count += 1

                    # Set thicknesses
                    for i, pid in enumerate(bar_pids):
                        self.current_bar_thicknesses[pid] = chromosome[i]
                    for i, pid in enumerate(skin_pids):
                        self.current_skin_thicknesses[pid] = chromosome[n_bars + i]

                    # Run Nastran
                    iter_folder = os.path.join(base_folder, f"gen_{gen+1:03d}_eval_{nastran_count:03d}")
                    os.makedirs(iter_folder, exist_ok=True)
                    result = self._run_iteration(iter_folder, nastran_count, target_rf)

                    if result:
                        min_rf = result['min_rf']
                        weight = result['weight']

                        # Update surrogate training data
                        X_train.append(chromosome)
                        y_rf_train.append(min_rf)
                        y_weight_train.append(weight)

                        # Calculate actual fitness
                        if min_rf >= target_rf - rf_tol:
                            actual_fitness = 1000.0 / (weight + 0.001)
                        else:
                            penalty = (target_rf - min_rf) * 100
                            actual_fitness = 1000.0 / (weight + 0.001) - penalty

                        pred_rf = cand['pred_rf']
                        self.log(f"    Nastran #{nastran_count}: Pred RF={pred_rf:.3f} → Actual RF={min_rf:.4f}, Weight={weight:.6f}t")

                        # Update best
                        if actual_fitness > best_fitness:
                            best_fitness = actual_fitness
                            best_chromosome = chromosome.copy()
                            best_result = result
                            self.best_solution = result
                            self.log(f"    *** NEW BEST! ***")

                        # Update fitness score in population
                        for i, p in enumerate(population):
                            if p == chromosome:
                                fitness_scores[i] = actual_fitness
                                break

                # Selection and reproduction
                def tournament_select(pop, fit, k=3):
                    indices = random.sample(range(len(pop)), min(k, len(pop)))
                    best_idx = max(indices, key=lambda i: fit[i])
                    return pop[best_idx]

                new_population = []
                if best_chromosome:
                    new_population.append(best_chromosome.copy())

                while len(new_population) < pop_size:
                    parent1 = tournament_select(population, fitness_scores)
                    parent2 = tournament_select(population, fitness_scores)

                    # BLX-α crossover
                    if random.random() < crossover_rate:
                        child = []
                        alpha = 0.5
                        for i, (g1, g2) in enumerate(zip(parent1, parent2)):
                            d = abs(g2 - g1)
                            if i < n_bars:
                                low = max(bar_min, min(g1, g2) - alpha * d)
                                high = min(bar_max, max(g1, g2) + alpha * d)
                            else:
                                low = max(skin_min, min(g1, g2) - alpha * d)
                                high = min(skin_max, max(g1, g2) + alpha * d)
                            child.append(random.uniform(low, high))
                    else:
                        child = parent1.copy()

                    # Gaussian mutation
                    for i in range(len(child)):
                        if random.random() < mutation_rate:
                            if i < n_bars:
                                sigma = (bar_max - bar_min) * 0.1
                                child[i] = max(bar_min, min(bar_max, child[i] + random.gauss(0, sigma)))
                            else:
                                sigma = (skin_max - skin_min) * 0.1
                                child[i] = max(skin_min, min(skin_max, child[i] + random.gauss(0, sigma)))

                    new_population.append(child)

                population = new_population[:pop_size]

                # Log generation summary
                self.log(f"\n  Gen {gen+1}: Best Fitness={best_fitness:.2f}, Surrogate samples={len(X_train)}")

            # ========== Final Results ==========
            self.log("\n" + "="*70)
            self.log("SURROGATE-ASSISTED GA COMPLETE")
            self.log("="*70)
            self.log(f"Total Nastran evaluations: {nastran_count}")
            self.log(f"Surrogate model size: {len(X_train)} samples")

            if best_result:
                for i, pid in enumerate(bar_pids):
                    self.current_bar_thicknesses[pid] = best_chromosome[i]
                for i, pid in enumerate(skin_pids):
                    self.current_skin_thicknesses[pid] = best_chromosome[n_bars + i]

                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {best_result['min_rf']:.4f}")
                self.log(f"  Weight: {best_result['weight']:.6f}t")
                self.log(f"  Failures: {best_result['n_fail']}")

                self._update_ui()
                self._save_results(base_folder)

                # Save surrogate training data
                surrogate_data = []
                for i in range(len(X_train)):
                    row = {'RF': y_rf_train[i], 'Weight': y_weight_train[i]}
                    for j, pid in enumerate(bar_pids):
                        row[f'Bar_{pid}'] = X_train[i][j]
                    for j, pid in enumerate(skin_pids):
                        row[f'Skin_{pid}'] = X_train[i][n_bars + j]
                    surrogate_data.append(row)
                pd.DataFrame(surrogate_data).to_csv(os.path.join(base_folder, "surrogate_data.csv"), index=False)
                self.log(f"\nSurrogate training data saved to surrogate_data.csv")

                self._generate_final_report(base_folder, "Surrogate-Assisted GA", nastran_count=nastran_count,
                    extra_info={
                        "Initial Samples": initial_samples,
                        "Nastran per Generation": nastran_per_gen,
                        "Surrogate Model Size": len(X_train),
                        "Savings vs Full GA": f"{(1 - nastran_count/(pop_size*n_generations))*100:.1f}%"
                    })

                self.root.after(0, lambda: messagebox.showinfo(
                    "Surrogate-Assisted GA Complete",
                    f"Optimization finished!\n\nNastran evaluations: {nastran_count}\nWeight: {best_result['weight']:.6f}t\nMin RF: {best_result['min_rf']:.4f}"
                ))
            else:
                self.log("ERROR: No valid solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _surrogate_predict(self, chromosome, X_train, y_rf_train, y_weight_train,
                           bar_min, bar_max, skin_min, skin_max, n_bars):
        """Predict RF and Weight using weighted k-nearest neighbors (RBF-like)."""
        if len(X_train) == 0:
            return 0.5, 1.0, 0.0  # Default prediction, low confidence

        # Normalize chromosome
        def normalize(chrom):
            norm = []
            for i, val in enumerate(chrom):
                if i < n_bars:
                    norm.append((val - bar_min) / (bar_max - bar_min + 1e-10))
                else:
                    norm.append((val - skin_min) / (skin_max - skin_min + 1e-10))
            return norm

        norm_query = normalize(chromosome)

        # Calculate distances to all training points
        distances = []
        for i, x in enumerate(X_train):
            norm_x = normalize(x)
            dist = sum((a - b) ** 2 for a, b in zip(norm_query, norm_x)) ** 0.5
            distances.append((dist, i))

        # Sort by distance and take k nearest
        k = min(5, len(X_train))
        distances.sort(key=lambda x: x[0])
        nearest = distances[:k]

        # Weighted average (inverse distance weighting)
        epsilon = 1e-6
        total_weight = 0
        pred_rf = 0
        pred_wt = 0

        for dist, idx in nearest:
            w = 1.0 / (dist + epsilon)
            total_weight += w
            pred_rf += w * y_rf_train[idx]
            pred_wt += w * y_weight_train[idx]

        pred_rf /= total_weight
        pred_wt /= total_weight

        # Confidence based on average distance to nearest neighbors
        avg_dist = sum(d for d, _ in nearest) / k
        confidence = 1.0 / (1.0 + avg_dist * 10)  # Higher distance = lower confidence

        return pred_rf, pred_wt, confidence

    def _run_rsm_sqp(self):
        """Algorithm 6: RSM + SQP - Response Surface Methodology with Sequential Quadratic Programming."""
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: RSM + SQP")
            self.log("(Response Surface Methodology + Sequential Quadratic Programming)")
            self.log("="*70)

            # Parameters
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            max_iter = int(self.max_iterations.get())

            bar_pids = list(self.bar_properties.keys())
            skin_pids = list(self.skin_properties.keys())
            n_bars = len(bar_pids)
            n_skins = len(skin_pids)
            n_vars = n_bars + n_skins

            if n_vars == 0:
                self.log("ERROR: No properties to optimize!")
                return

            self.log(f"\nVariables: {n_bars} bars + {n_skins} skins = {n_vars} total")
            self.log(f"Target RF: {target_rf}")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"rsm_sqp_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # Bounds for each variable
            bounds_low = [bar_min] * n_bars + [skin_min] * n_skins
            bounds_high = [bar_max] * n_bars + [skin_max] * n_skins

            # ========== PHASE 1: DOE - Latin Hypercube Sampling ==========
            self.log("\n" + "="*50)
            self.log("PHASE 1: Design of Experiments (Latin Hypercube)")
            self.log("="*50)

            # Number of initial samples: need at least n_vars + 1 for RSM fitting
            # Use 2*n_vars + 1 as rule of thumb for good fit
            n_samples = max(n_vars + 5, 2 * n_vars + 1, 20)

            # Cap to reasonable number to avoid too many Nastran runs
            max_doe_samples = 150
            if n_samples > max_doe_samples:
                self.log(f"WARNING: Large number of variables ({n_vars}). Limiting DOE to {max_doe_samples} samples.")
                self.log(f"         Consider using Surrogate-Assisted GA for problems with many variables.")
                n_samples = max_doe_samples

            self.log(f"Generating {n_samples} DOE samples for {n_vars} variables...")

            # Latin Hypercube Sampling
            doe_samples = self._latin_hypercube_sampling(n_samples, n_vars, bounds_low, bounds_high)

            # Evaluate DOE samples with Nastran
            X_data = []  # Design points
            y_rf_data = []  # RF responses
            y_weight_data = []  # Weight responses
            nastran_count = 0

            best_result = None
            best_fitness = float('-inf')

            for idx, sample in enumerate(doe_samples):
                if not self.is_running:
                    break

                nastran_count += 1
                self.log(f"\n  DOE Sample {idx+1}/{n_samples} (Nastran #{nastran_count})")
                self.update_progress((idx/n_samples)*30, f"DOE {idx+1}/{n_samples}")

                # Set thicknesses
                for i, pid in enumerate(bar_pids):
                    self.current_bar_thicknesses[pid] = sample[i]
                for i, pid in enumerate(skin_pids):
                    self.current_skin_thicknesses[pid] = sample[n_bars + i]

                # Run Nastran
                iter_folder = os.path.join(base_folder, f"doe_{idx+1:03d}")
                os.makedirs(iter_folder, exist_ok=True)
                result = self._run_iteration(iter_folder, nastran_count, target_rf)

                if result:
                    X_data.append(sample)
                    y_rf_data.append(result['min_rf'])
                    y_weight_data.append(result['weight'])

                    self.log(f"    RF={result['min_rf']:.4f}, Weight={result['weight']:.6f}t")

                    # Track best feasible solution
                    if result['min_rf'] >= target_rf - rf_tol:
                        fitness = 1000.0 / (result['weight'] + 0.001)
                        if fitness > best_fitness:
                            best_fitness = fitness
                            best_result = result
                            self.best_solution = result
                            self.log(f"    *** FEASIBLE SOLUTION ***")

            if len(X_data) < n_vars + 1:
                self.log(f"ERROR: Not enough valid DOE samples!")
                self.log(f"       Valid samples: {len(X_data)}, Required: {n_vars + 1} (n_vars + 1)")
                self.log(f"       Most Nastran runs may have failed. Check BDF file and solver settings.")
                self.log(f"       Alternatively, try Surrogate-Assisted GA algorithm instead.")
                return

            X_data = np.array(X_data)
            y_rf_data = np.array(y_rf_data)
            y_weight_data = np.array(y_weight_data)

            self.log(f"\n  DOE Complete: {len(X_data)} valid samples")

            # ========== PHASE 2: Fit RSM Models ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: Fitting Response Surface Models")
            self.log("="*50)

            # Fit quadratic RSM for RF and Weight
            rsm_rf = self._fit_rsm(X_data, y_rf_data, bounds_low, bounds_high)
            rsm_weight = self._fit_rsm(X_data, y_weight_data, bounds_low, bounds_high)

            if rsm_rf is None or rsm_weight is None:
                self.log("ERROR: RSM fitting failed!")
                return

            self.log(f"  RF Model R²: {rsm_rf['r2']:.4f}")
            self.log(f"  Weight Model R²: {rsm_weight['r2']:.4f}")

            # ========== PHASE 3: SQP Optimization Loop ==========
            self.log("\n" + "="*50)
            self.log("PHASE 3: SQP Optimization with Trust Region")
            self.log("="*50)

            # Start from best DOE point or center
            if best_result:
                x0 = list(best_result['bar_thicknesses'].values()) + list(best_result.get('skin_thicknesses', {}).values())
                if len(x0) != n_vars:
                    x0 = [(l + h) / 2 for l, h in zip(bounds_low, bounds_high)]
            else:
                x0 = [(l + h) / 2 for l, h in zip(bounds_low, bounds_high)]

            trust_radius = 1.0  # Normalized trust region radius
            min_trust_radius = 0.1
            sqp_iterations = 0
            max_sqp_iterations = max_iter

            while sqp_iterations < max_sqp_iterations and self.is_running:
                sqp_iterations += 1
                self.log(f"\n  SQP Iteration {sqp_iterations}")
                self.update_progress(30 + (sqp_iterations/max_sqp_iterations)*60, f"SQP Iter {sqp_iterations}")

                # Define objective: minimize weight subject to RF >= target
                def objective(x):
                    return self._rsm_predict(x, rsm_weight, bounds_low, bounds_high)

                def rf_constraint(x):
                    return self._rsm_predict(x, rsm_rf, bounds_low, bounds_high) - target_rf

                # Bounds with trust region
                tr_bounds = []
                for i in range(n_vars):
                    x_norm = (x0[i] - bounds_low[i]) / (bounds_high[i] - bounds_low[i])
                    range_i = bounds_high[i] - bounds_low[i]
                    tr_low = max(bounds_low[i], x0[i] - trust_radius * range_i)
                    tr_high = min(bounds_high[i], x0[i] + trust_radius * range_i)
                    tr_bounds.append((tr_low, tr_high))

                # SQP optimization
                constraints = {'type': 'ineq', 'fun': rf_constraint}

                try:
                    result_sqp = minimize(
                        objective, x0,
                        method='SLSQP',
                        bounds=tr_bounds,
                        constraints=constraints,
                        options={'maxiter': 100, 'ftol': 1e-6}
                    )
                    x_opt = result_sqp.x
                    pred_weight = result_sqp.fun
                    pred_rf = self._rsm_predict(x_opt, rsm_rf, bounds_low, bounds_high)

                    self.log(f"    SQP Result: Pred RF={pred_rf:.4f}, Pred Weight={pred_weight:.6f}")

                except Exception as e:
                    self.log(f"    SQP failed: {e}")
                    break

                # Validate with Nastran
                nastran_count += 1
                self.log(f"    Validating with Nastran #{nastran_count}...")

                for i, pid in enumerate(bar_pids):
                    self.current_bar_thicknesses[pid] = float(x_opt[i])
                for i, pid in enumerate(skin_pids):
                    self.current_skin_thicknesses[pid] = float(x_opt[n_bars + i])

                iter_folder = os.path.join(base_folder, f"sqp_{sqp_iterations:03d}")
                os.makedirs(iter_folder, exist_ok=True)
                result = self._run_iteration(iter_folder, nastran_count, target_rf)

                if result:
                    actual_rf = result['min_rf']
                    actual_weight = result['weight']

                    self.log(f"    Actual: RF={actual_rf:.4f}, Weight={actual_weight:.6f}t")

                    # Calculate prediction error
                    rf_error = abs(pred_rf - actual_rf) / max(actual_rf, 0.001)
                    weight_error = abs(pred_weight - actual_weight) / max(actual_weight, 0.001)

                    self.log(f"    Errors: RF={rf_error*100:.1f}%, Weight={weight_error*100:.1f}%")

                    # Update training data
                    X_data = np.vstack([X_data, x_opt])
                    y_rf_data = np.append(y_rf_data, actual_rf)
                    y_weight_data = np.append(y_weight_data, actual_weight)

                    # Check if solution is feasible
                    if actual_rf >= target_rf - rf_tol:
                        fitness = 1000.0 / (actual_weight + 0.001)
                        if fitness > best_fitness:
                            best_fitness = fitness
                            best_result = result
                            self.best_solution = result
                            self.log(f"    *** NEW BEST! ***")

                    # Trust region update
                    if rf_error < 0.1 and weight_error < 0.1:
                        # Good prediction, expand trust region
                        trust_radius = min(trust_radius * 1.5, 1.0)
                        x0 = list(x_opt)
                        self.log(f"    Trust region expanded: {trust_radius:.2f}")
                    elif rf_error > 0.3 or weight_error > 0.3:
                        # Poor prediction, shrink trust region and refit RSM
                        trust_radius = max(trust_radius * 0.5, min_trust_radius)
                        self.log(f"    Trust region shrunk: {trust_radius:.2f}")

                        # Refit RSM with new data
                        rsm_rf = self._fit_rsm(X_data, y_rf_data, bounds_low, bounds_high)
                        rsm_weight = self._fit_rsm(X_data, y_weight_data, bounds_low, bounds_high)
                        self.log(f"    RSM refitted: RF R²={rsm_rf['r2']:.4f}, Weight R²={rsm_weight['r2']:.4f}")
                    else:
                        # Acceptable prediction
                        x0 = list(x_opt)
                        # Refit RSM periodically
                        if sqp_iterations % 3 == 0:
                            rsm_rf = self._fit_rsm(X_data, y_rf_data, bounds_low, bounds_high)
                            rsm_weight = self._fit_rsm(X_data, y_weight_data, bounds_low, bounds_high)

                    # Check convergence
                    if trust_radius <= min_trust_radius and rf_error < 0.05:
                        self.log(f"\n  CONVERGED at iteration {sqp_iterations}")
                        break

                else:
                    self.log(f"    Nastran failed!")
                    trust_radius = max(trust_radius * 0.5, min_trust_radius)

            # ========== Final Results ==========
            self.log("\n" + "="*70)
            self.log("RSM + SQP OPTIMIZATION COMPLETE")
            self.log("="*70)
            self.log(f"Total Nastran evaluations: {nastran_count}")
            self.log(f"DOE samples: {n_samples}")
            self.log(f"SQP iterations: {sqp_iterations}")

            if best_result:
                for i, pid in enumerate(bar_pids):
                    self.current_bar_thicknesses[pid] = best_result['bar_thicknesses'].get(pid, bar_min)
                for i, pid in enumerate(skin_pids):
                    self.current_skin_thicknesses[pid] = best_result.get('skin_thicknesses', {}).get(pid, skin_min)

                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {best_result['min_rf']:.4f}")
                self.log(f"  Weight: {best_result['weight']:.6f}t")
                self.log(f"  Failures: {best_result['n_fail']}")

                self._update_ui()
                self._save_results(base_folder)
                self._generate_final_report(base_folder, "RSM + SQP", nastran_count=nastran_count,
                    extra_info={
                        "DOE Samples": n_samples,
                        "SQP Iterations": sqp_iterations,
                        "Final RF Model R²": f"{rsm_rf['r2']:.4f}",
                        "Final Weight Model R²": f"{rsm_weight['r2']:.4f}",
                        "RSM Model Size": len(X_data)
                    })

                # Save RSM data
                rsm_data = []
                for i in range(len(X_data)):
                    row = {'RF': y_rf_data[i], 'Weight': y_weight_data[i]}
                    for j, pid in enumerate(bar_pids):
                        row[f'Bar_{pid}'] = X_data[i][j]
                    for j, pid in enumerate(skin_pids):
                        row[f'Skin_{pid}'] = X_data[i][n_bars + j]
                    rsm_data.append(row)
                pd.DataFrame(rsm_data).to_csv(os.path.join(base_folder, "rsm_training_data.csv"), index=False)

                self.root.after(0, lambda: messagebox.showinfo(
                    "RSM + SQP Complete",
                    f"Optimization finished!\n\nNastran evaluations: {nastran_count}\nWeight: {best_result['weight']:.6f}t\nMin RF: {best_result['min_rf']:.4f}"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_topdown_algorithm(self):
        """Algorithm 7: Top-Down - Start from max thickness, reduce until RF reaches target.

        This is a Fully Stressed Design (FSD) approach:
        1. Start with maximum thicknesses (structure is over-designed, RF >> target)
        2. Reduce thicknesses proportionally based on RF margin
        3. Stop when RF approaches target (optimal weight)

        Advantages:
        - Always starts from feasible region (no failures initially)
        - Converges to minimum weight with RF ≈ target
        - Simple and robust
        """
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: TOP-DOWN (MAX TO TARGET)")
            self.log("="*70)
            self.log("Strategy: Start overdesigned, reduce until RF = target")
            self.log("="*70)

            max_iter = int(self.max_iterations.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())

            # Calculate bar-skin proximity mapping
            self.calculate_bar_skin_proximity()

            # Create output folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"topdown_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== PHASE 1: Initialize thicknesses ==========
            self.log("\n" + "="*50)
            self.log("PHASE 1: Initialize thicknesses")
            self.log("="*50)

            # Bar properties: ALL start at MAXIMUM
            for pid in self.bar_properties:
                self.current_bar_thicknesses[pid] = bar_max

            # Skin properties: Related (near bars) → MAX, Unrelated → MIN
            # Find all skin PIDs that are related to any bar
            related_skins = set()
            for bar_pid, nearby_skins in self.bar_to_nearby_skins.items():
                related_skins.update(nearby_skins)

            related_count = 0
            unrelated_count = 0
            for pid in self.skin_properties:
                if pid in related_skins:
                    # Related to bars - start at MAX (will be optimized)
                    self.current_skin_thicknesses[pid] = skin_max
                    related_count += 1
                else:
                    # Not related to any bar - start at MIN (save weight)
                    self.current_skin_thicknesses[pid] = skin_min
                    unrelated_count += 1

            self.log(f"  Bar thicknesses: {len(self.bar_properties)} properties → MAX ({bar_max})")
            self.log(f"  Skin thicknesses:")
            self.log(f"    Related (near bars): {related_count} properties → MAX ({skin_max})")
            self.log(f"    Unrelated (far): {unrelated_count} properties → MIN ({skin_min})")

            # ========== PHASE 2: Iterative Lightening ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: Iterative Lightening")
            self.log("="*50)

            best_weight = float('inf')
            converged = False
            no_improvement_count = 0
            max_no_improvement = 5

            for iteration in range(1, max_iter + 1):
                if not self.is_running:
                    break

                self.log(f"\n{'='*50}")
                self.log(f"ITERATION {iteration}/{max_iter}")
                self.log(f"{'='*50}")

                iter_folder = os.path.join(base_folder, f"iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                self.update_progress((iteration/max_iter)*100, f"Iter {iteration}/{max_iter}")

                # Run Nastran and get RF
                result = self._run_iteration(iter_folder, iteration, target_rf)

                if not result:
                    self.log("  ERROR: Iteration failed!")
                    continue

                self.iteration_results.append(result)
                min_rf = result['min_rf']
                weight = result['weight']
                n_fail = result['n_fail']

                self.log(f"\n  Summary: Min RF={min_rf:.4f}, Fails={n_fail}, Weight={weight:.6f}t")

                # Calculate per-property RF statistics
                rf_details = result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                # Count how many properties are converged (RF within tolerance)
                n_converged = 0
                n_total = 0
                for pid, rf in pid_min_rf.items():
                    n_total += 1
                    if target_rf - rf_tol <= rf <= target_rf + rf_tol:
                        n_converged += 1

                convergence_pct = (n_converged / n_total * 100) if n_total > 0 else 0
                self.log(f"  Per-Property Convergence: {n_converged}/{n_total} ({convergence_pct:.1f}%) properties at RF≈{target_rf}")

                # Check convergence: ALL properties should be within tolerance
                if n_total > 0 and n_converged == n_total:
                    self.log(f"\n  *** FULLY CONVERGED! All {n_total} properties have RF≈{target_rf} ***")
                    converged = True
                    if weight < best_weight:
                        best_weight = weight
                        self.best_solution = result
                    break

                # Track best feasible solution (min RF >= target)
                if min_rf >= target_rf - rf_tol and weight < best_weight:
                    best_weight = weight
                    self.best_solution = result
                    self.log(f"  *** NEW BEST: Weight={weight:.6f}t ***")
                    no_improvement_count = 0
                else:
                    no_improvement_count += 1

                # Check if stuck
                if no_improvement_count >= max_no_improvement:
                    self.log(f"\n  No improvement for {max_no_improvement} iterations.")

                # ========== FULLY STRESSED DESIGN: Per-Property Update ==========
                # Each property is updated based on its OWN RF, not global min RF
                # Goal: Make ALL properties have RF ≈ target (1.0)

                self.log(f"\n  Applying Fully Stressed Design (per-property)...")

                # pid_min_rf already calculated above

                # Stress ratio exponent (alpha < 1 for stability)
                alpha = 0.5  # Higher value = faster convergence but less stable

                # Track changes
                increased_count = 0
                decreased_count = 0
                unchanged_count = 0

                # Update BAR thicknesses - each property independently
                for pid in self.bar_properties:
                    old_t = self.current_bar_thicknesses[pid]
                    prop_rf = pid_min_rf.get(pid, None)

                    if prop_rf is None:
                        # No RF data for this property, skip
                        unchanged_count += 1
                        continue

                    # Stress Ratio Method: new_t = old_t * (target_rf / actual_rf)^alpha
                    if prop_rf > target_rf + rf_tol:
                        # Over-designed: REDUCE thickness
                        ratio = (target_rf / prop_rf) ** alpha
                        ratio = max(ratio, 0.7)  # Max 30% reduction per iteration
                        new_t = old_t * ratio
                        decreased_count += 1
                    elif prop_rf < target_rf - rf_tol:
                        # Under-designed: INCREASE thickness
                        ratio = (target_rf / prop_rf) ** alpha
                        ratio = min(ratio, 1.3)  # Max 30% increase per iteration
                        new_t = old_t * ratio
                        increased_count += 1
                    else:
                        # Within tolerance, keep it
                        new_t = old_t
                        unchanged_count += 1

                    new_t = max(bar_min, min(bar_max, new_t))
                    self.current_bar_thicknesses[pid] = new_t

                # Update SKIN thicknesses - PROXIMITY-BASED (coupled optimization)
                # Each skin is updated based on the RF of NEARBY bars (within search distance)
                # Skin doesn't have its own allowable, but affects bar stress

                skin_decreased = 0
                skin_increased = 0
                skin_unchanged = 0

                # Build: skin_pid -> controlling bar RF (min RF of nearby bars)
                skin_controlling_rf = {}
                for bar_pid, nearby_skins in self.bar_to_nearby_skins.items():
                    bar_rf = pid_min_rf.get(bar_pid, None)
                    if bar_rf is None:
                        continue
                    for skin_pid in nearby_skins:
                        if skin_pid not in skin_controlling_rf or bar_rf < skin_controlling_rf[skin_pid]:
                            skin_controlling_rf[skin_pid] = bar_rf

                for skin_pid in self.skin_properties:
                    old_t = self.current_skin_thicknesses[skin_pid]

                    # Use the controlling bar RF for this skin (min RF of nearby bars)
                    controlling_rf = skin_controlling_rf.get(skin_pid, None)

                    if controlling_rf is None:
                        # No nearby bars - skin is not coupled, keep unchanged
                        skin_unchanged += 1
                        new_t = old_t
                    elif controlling_rf > target_rf + rf_tol:
                        # Nearby bars are over-designed: REDUCE skin thickness
                        skin_ratio = (target_rf / controlling_rf) ** (alpha * 0.7)
                        skin_ratio = max(skin_ratio, 0.85)  # Max 15% reduction
                        new_t = old_t * skin_ratio
                        skin_decreased += 1
                    elif controlling_rf < target_rf - rf_tol:
                        # Nearby bars are under-designed: INCREASE skin thickness
                        skin_ratio = (target_rf / controlling_rf) ** (alpha * 0.7)
                        skin_ratio = min(skin_ratio, 1.15)  # Max 15% increase
                        new_t = old_t * skin_ratio
                        skin_increased += 1
                    else:
                        # Within tolerance
                        new_t = old_t
                        skin_unchanged += 1

                    new_t = max(skin_min, min(skin_max, new_t))
                    self.current_skin_thicknesses[skin_pid] = new_t

                self.log(f"    Bar properties: {increased_count} increased, {decreased_count} decreased, {unchanged_count} unchanged")
                self.log(f"    Skin properties: {skin_increased} increased, {skin_decreased} decreased, {skin_unchanged} unchanged")
                self.log(f"    (Skin updated based on nearby bar RF - {len(skin_controlling_rf)} skins coupled)")

            # ========== FINAL RESULTS ==========
            self.log("\n" + "="*70)
            self.log("TOP-DOWN OPTIMIZATION COMPLETE")
            self.log("="*70)

            if self.best_solution:
                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {self.best_solution['min_rf']:.4f}")
                self.log(f"  Weight: {self.best_solution['weight']:.6f} tonnes")
                self.log(f"  Found at iteration: {self.best_solution['iteration']}")
                self.log(f"  Converged: {'Yes' if converged else 'No'}")

                self._generate_final_report(base_folder, "Top-Down (Max to Target)",
                    nastran_count=iteration,
                    extra_info={
                        "Strategy": "Fully Stressed Design",
                        "Start": "Maximum thicknesses",
                        "Converged": "Yes" if converged else "No",
                        "Final Iterations": iteration
                    })

                self.root.after(0, lambda: self.result_summary.config(
                    text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}",
                    foreground="green"
                ))

                self.root.after(0, lambda: messagebox.showinfo(
                    "Top-Down Complete",
                    f"Optimization Complete!\n\n"
                    f"Best RF: {self.best_solution['min_rf']:.4f}\n"
                    f"Best Weight: {self.best_solution['weight']:.6f} tonnes\n"
                    f"Iterations: {iteration}\n"
                    f"Converged: {'Yes' if converged else 'No'}"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_bottomup_algorithm(self):
        """Algorithm 8: Bottom-Up - Start from min thickness, increase until RF reaches target.

        This is an inverse Fully Stressed Design (FSD) approach:
        1. Start with minimum thicknesses (structure is under-designed, RF << target)
        2. Increase thicknesses proportionally based on RF deficit
        3. Stop when RF approaches target (optimal weight)

        Advantages:
        - Always starts from lightest possible design
        - Builds up only where needed
        - Converges to minimum weight with RF ≈ target
        """
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: BOTTOM-UP (MIN TO TARGET)")
            self.log("="*70)
            self.log("Strategy: Start underdesigned, increase until RF = target")
            self.log("="*70)

            max_iter = int(self.max_iterations.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())
            use_gfem = self.use_gfem_thickness.get()
            keep_bars_const = self.keep_bars_constant.get()
            keep_skins_const = self.keep_skins_constant.get()

            # Calculate bar-skin proximity mapping
            self.calculate_bar_skin_proximity()

            # Create output folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"bottomup_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== PHASE 1: Initialize thicknesses ==========
            self.log("\n" + "="*50)
            if use_gfem:
                self.log("PHASE 1: Initialize thicknesses at GFEM values")
                self.log("(Using Excel thickness as initial and minimum)")
            else:
                self.log("PHASE 1: Initialize thicknesses at MINIMUM")
            if keep_bars_const:
                self.log("*** BAR thicknesses CONSTANT (from Excel) - only skins will be optimized ***")
            if keep_skins_const:
                self.log("*** SKIN thicknesses CONSTANT (from Excel) - only bars will be optimized ***")
            self.log("="*50)

            # Bar properties: Initialize based on mode
            for pid in self.bar_properties:
                if keep_bars_const and pid in self.gfem_bar_thicknesses:
                    self.current_bar_thicknesses[pid] = self.gfem_bar_thicknesses[pid]
                elif use_gfem and pid in self.gfem_bar_thicknesses:
                    self.current_bar_thicknesses[pid] = self.gfem_bar_thicknesses[pid]
                else:
                    self.current_bar_thicknesses[pid] = bar_min

            # Skin properties: Related (near bars) → initial, Unrelated → initial
            # Both start at initial value, but related skins will be optimized with bars
            related_skins = set()
            for bar_pid, nearby_skins in self.bar_to_nearby_skins.items():
                related_skins.update(nearby_skins)

            related_count = 0
            unrelated_count = 0
            for pid in self.skin_properties:
                if keep_skins_const and pid in self.gfem_skin_thicknesses:
                    init_val = self.gfem_skin_thicknesses[pid]
                elif use_gfem and pid in self.gfem_skin_thicknesses:
                    init_val = self.gfem_skin_thicknesses[pid]
                else:
                    init_val = skin_min
                if pid in related_skins:
                    self.current_skin_thicknesses[pid] = init_val
                    related_count += 1
                else:
                    self.current_skin_thicknesses[pid] = init_val
                    unrelated_count += 1

            if keep_bars_const:
                avg_bar = sum(self.current_bar_thicknesses.values()) / len(self.current_bar_thicknesses) if self.current_bar_thicknesses else 0
                self.log(f"  Bar thicknesses: {len(self.bar_properties)} properties → CONSTANT (avg: {avg_bar:.2f})")
            elif use_gfem:
                avg_bar = sum(self.gfem_bar_thicknesses.values()) / len(self.gfem_bar_thicknesses) if self.gfem_bar_thicknesses else 0
                self.log(f"  Bar thicknesses: {len(self.bar_properties)} properties → GFEM (avg: {avg_bar:.2f})")
            else:
                self.log(f"  Bar thicknesses: {len(self.bar_properties)} properties → MIN ({bar_min})")

            if keep_skins_const:
                avg_skin = sum(self.current_skin_thicknesses.values()) / len(self.current_skin_thicknesses) if self.current_skin_thicknesses else 0
                self.log(f"  Skin thicknesses: {len(self.skin_properties)} properties → CONSTANT (avg: {avg_skin:.2f})")
            elif use_gfem:
                avg_skin = sum(self.gfem_skin_thicknesses.values()) / len(self.gfem_skin_thicknesses) if self.gfem_skin_thicknesses else 0
                self.log(f"  Skin thicknesses:")
                self.log(f"    Related (near bars): {related_count} properties → GFEM (avg: {avg_skin:.2f})")
                self.log(f"    Unrelated (far): {unrelated_count} properties → GFEM")
            else:
                self.log(f"  Skin thicknesses:")
                self.log(f"    Related (near bars): {related_count} properties → MIN ({skin_min})")
                self.log(f"    Unrelated (far): {unrelated_count} properties → MIN ({skin_min})")

            # ========== PHASE 2: Iterative Strengthening ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: Iterative Strengthening")
            self.log("="*50)

            best_weight = float('inf')
            converged = False
            no_improvement_count = 0
            max_no_improvement = 5

            for iteration in range(1, max_iter + 1):
                if not self.is_running:
                    break

                self.log(f"\n{'='*50}")
                self.log(f"ITERATION {iteration}/{max_iter}")
                self.log(f"{'='*50}")

                iter_folder = os.path.join(base_folder, f"iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                self.update_progress((iteration/max_iter)*100, f"Iter {iteration}/{max_iter}")

                # Run Nastran and get RF
                result = self._run_iteration(iter_folder, iteration, target_rf)

                if not result:
                    self.log("  ERROR: Iteration failed!")
                    continue

                self.iteration_results.append(result)
                min_rf = result['min_rf']
                weight = result['weight']
                n_fail = result['n_fail']

                self.log(f"\n  Summary: Min RF={min_rf:.4f}, Fails={n_fail}, Weight={weight:.6f}t")

                # Calculate per-property RF statistics
                rf_details = result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                # Count how many properties are converged (RF within tolerance)
                n_converged = 0
                n_total = 0
                for pid, rf in pid_min_rf.items():
                    n_total += 1
                    if target_rf - rf_tol <= rf <= target_rf + rf_tol:
                        n_converged += 1

                convergence_pct = (n_converged / n_total * 100) if n_total > 0 else 0
                self.log(f"  Per-Property Convergence: {n_converged}/{n_total} ({convergence_pct:.1f}%) properties at RF≈{target_rf}")

                # Check convergence: ALL properties should be within tolerance
                if n_total > 0 and n_converged == n_total:
                    self.log(f"\n  *** FULLY CONVERGED! All {n_total} properties have RF≈{target_rf} ***")
                    converged = True
                    if weight < best_weight:
                        best_weight = weight
                        self.best_solution = result
                    break

                # Track best feasible solution (min RF >= target)
                if min_rf >= target_rf - rf_tol and weight < best_weight:
                    best_weight = weight
                    self.best_solution = result
                    self.log(f"  *** NEW BEST: Weight={weight:.6f}t ***")
                    no_improvement_count = 0
                else:
                    no_improvement_count += 1

                # Check if stuck
                if no_improvement_count >= max_no_improvement:
                    self.log(f"\n  No improvement for {max_no_improvement} iterations.")

                # ========== FULLY STRESSED DESIGN: Per-Property Update ==========
                # Each property is updated based on its OWN RF, not global min RF
                # Goal: Make ALL properties have RF ≈ target (1.0)

                self.log(f"\n  Applying Fully Stressed Design (per-property, bottom-up)...")

                # Stress ratio exponent (alpha < 1 for stability)
                alpha = 0.5

                # Track changes
                increased_count = 0
                decreased_count = 0
                unchanged_count = 0

                # Update BAR thicknesses - each property independently
                if keep_bars_const:
                    self.log(f"    Bars: CONSTANT (locked at Excel values)")
                else:
                    for pid in self.bar_properties:
                        old_t = self.current_bar_thicknesses[pid]
                        prop_rf = pid_min_rf.get(pid, None)

                        if prop_rf is None:
                            unchanged_count += 1
                            continue

                        # Stress Ratio Method: new_t = old_t * (target_rf / actual_rf)^alpha
                        if prop_rf < target_rf - rf_tol:
                            # Under-designed: INCREASE thickness (primary direction in bottom-up)
                            ratio = (target_rf / prop_rf) ** alpha
                            ratio = min(ratio, 1.3)  # Max 30% increase per iteration
                            new_t = old_t * ratio
                            increased_count += 1
                        elif prop_rf > target_rf + rf_tol:
                            # Over-designed: REDUCE thickness
                            ratio = (target_rf / prop_rf) ** alpha
                            ratio = max(ratio, 0.7)  # Max 30% reduction per iteration
                            new_t = old_t * ratio
                            decreased_count += 1
                        else:
                            # Within tolerance, keep it
                            new_t = old_t
                            unchanged_count += 1

                        # Use GFEM thickness as minimum if enabled
                        if use_gfem and pid in self.gfem_bar_thicknesses:
                            pid_min = self.gfem_bar_thicknesses[pid]
                        else:
                            pid_min = bar_min
                        new_t = max(pid_min, min(bar_max, new_t))
                        self.current_bar_thicknesses[pid] = new_t

                # Update SKIN thicknesses - PROXIMITY-BASED (coupled optimization)
                skin_decreased = 0
                skin_increased = 0
                skin_unchanged = 0

                if keep_skins_const:
                    self.log(f"    Skins: CONSTANT (locked at Excel values)")
                else:
                    # Build: skin_pid -> controlling bar RF (min RF of nearby bars)
                    skin_controlling_rf = {}
                    for bar_pid, nearby_skins in self.bar_to_nearby_skins.items():
                        bar_rf = pid_min_rf.get(bar_pid, None)
                        if bar_rf is None:
                            continue
                        for skin_pid in nearby_skins:
                            if skin_pid not in skin_controlling_rf or bar_rf < skin_controlling_rf[skin_pid]:
                                skin_controlling_rf[skin_pid] = bar_rf

                    for skin_pid in self.skin_properties:
                        old_t = self.current_skin_thicknesses[skin_pid]

                        controlling_rf = skin_controlling_rf.get(skin_pid, None)

                        if controlling_rf is None:
                            # No nearby bars - skin is not coupled, keep at MIN
                            skin_unchanged += 1
                            new_t = old_t
                        elif controlling_rf < target_rf - rf_tol:
                            # Nearby bars are under-designed: INCREASE skin thickness
                            skin_ratio = (target_rf / controlling_rf) ** (alpha * 0.7)
                            skin_ratio = min(skin_ratio, 1.15)  # Max 15% increase
                            new_t = old_t * skin_ratio
                            skin_increased += 1
                        elif controlling_rf > target_rf + rf_tol:
                            # Nearby bars are over-designed: REDUCE skin thickness
                            skin_ratio = (target_rf / controlling_rf) ** (alpha * 0.7)
                            skin_ratio = max(skin_ratio, 0.85)  # Max 15% reduction
                            new_t = old_t * skin_ratio
                            skin_decreased += 1
                        else:
                            # Within tolerance
                            new_t = old_t
                            skin_unchanged += 1

                        # Use GFEM thickness as minimum if enabled
                        if use_gfem and skin_pid in self.gfem_skin_thicknesses:
                            spid_min = self.gfem_skin_thicknesses[skin_pid]
                        else:
                            spid_min = skin_min
                        new_t = max(spid_min, min(skin_max, new_t))
                        self.current_skin_thicknesses[skin_pid] = new_t

                if not keep_bars_const:
                    self.log(f"    Bar properties: {increased_count} increased, {decreased_count} decreased, {unchanged_count} unchanged")
                if not keep_skins_const:
                    self.log(f"    Skin properties: {skin_increased} increased, {skin_decreased} decreased, {skin_unchanged} unchanged")
                    self.log(f"    (Skin updated based on nearby bar RF - {len(skin_controlling_rf)} skins coupled)")

            # ========== FINAL RESULTS ==========
            self.log("\n" + "="*70)
            self.log("BOTTOM-UP OPTIMIZATION COMPLETE")
            self.log("="*70)

            if self.best_solution:
                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {self.best_solution['min_rf']:.4f}")
                self.log(f"  Weight: {self.best_solution['weight']:.6f} tonnes")
                self.log(f"  Found at iteration: {self.best_solution['iteration']}")
                self.log(f"  Converged: {'Yes' if converged else 'No'}")

                self._generate_final_report(base_folder, "Bottom-Up (Min to Target)",
                    nastran_count=iteration,
                    extra_info={
                        "Strategy": "Inverse Fully Stressed Design",
                        "Start": "Minimum thicknesses",
                        "Converged": "Yes" if converged else "No",
                        "Final Iterations": iteration
                    })

                self.root.after(0, lambda: self.result_summary.config(
                    text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}",
                    foreground="green"
                ))

                self.root.after(0, lambda: messagebox.showinfo(
                    "Bottom-Up Complete",
                    f"Optimization Complete!\n\n"
                    f"Best RF: {self.best_solution['min_rf']:.4f}\n"
                    f"Best Weight: {self.best_solution['weight']:.6f} tonnes\n"
                    f"Iterations: {iteration}\n"
                    f"Converged: {'Yes' if converged else 'No'}"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_weight_efficient_bottomup(self):
        """Algorithm 9: Weight-Efficient Bottom-Up - Prioritize lowest weight-cost strengthening.

        Key insight: Different properties have different weight costs per thickness increase.
        - Skin with large area: high weight cost (dW/dt = Area × density)
        - Bar with small cross-section: low weight cost (dW/dt = Length × dim2 × density)

        Strategy: When RF < target, increase the property with BEST efficiency:
            Efficiency = RF_deficit / weight_sensitivity

        This ensures minimum weight solution by always choosing the cheapest fix.
        """
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: WEIGHT-EFFICIENT BOTTOM-UP")
            self.log("="*70)
            self.log("Strategy: Prioritize strengthening by weight efficiency (ΔRF/ΔWeight)")
            self.log("="*70)

            max_iter = int(self.max_iterations.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())

            # Calculate bar-skin proximity mapping
            self.calculate_bar_skin_proximity()

            # Create output folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"weight_efficient_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== PHASE 1: Initialize at MINIMUM thicknesses ==========
            self.log("\n" + "="*50)
            self.log("PHASE 1: Initialize thicknesses at MINIMUM")
            self.log("="*50)

            for pid in self.bar_properties:
                self.current_bar_thicknesses[pid] = bar_min

            for pid in self.skin_properties:
                self.current_skin_thicknesses[pid] = skin_min

            self.log(f"  Bar thicknesses: {len(self.bar_properties)} properties → MIN ({bar_min})")
            self.log(f"  Skin thicknesses: {len(self.skin_properties)} properties → MIN ({skin_min})")

            # ========== PHASE 2: Calculate Weight Sensitivities ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: Calculate Weight Sensitivities")
            self.log("="*50)

            weight_sens = self._calculate_weight_sensitivities()

            # Log sensitivity comparison
            bar_sens_list = [(pid, weight_sens.get(pid, 0)) for pid in self.bar_properties]
            skin_sens_list = [(pid, weight_sens.get(pid, 0)) for pid in self.skin_properties]

            if bar_sens_list:
                avg_bar_sens = sum(s for _, s in bar_sens_list) / len(bar_sens_list)
                self.log(f"  Bar avg weight sensitivity: {avg_bar_sens:.6f} tonnes/mm")
            if skin_sens_list:
                avg_skin_sens = sum(s for _, s in skin_sens_list) / len(skin_sens_list)
                self.log(f"  Skin avg weight sensitivity: {avg_skin_sens:.6f} tonnes/mm")

            if bar_sens_list and skin_sens_list and avg_bar_sens > 0:
                ratio = avg_skin_sens / avg_bar_sens
                self.log(f"  Skin/Bar sensitivity ratio: {ratio:.1f}x")
                self.log(f"  → Increasing bar thickness is {ratio:.1f}x more weight-efficient!")

            # ========== PHASE 3: Weight-Efficient Iterative Strengthening ==========
            self.log("\n" + "="*50)
            self.log("PHASE 3: Weight-Efficient Iterative Strengthening")
            self.log("="*50)

            best_weight = float('inf')
            converged = False

            for iteration in range(1, max_iter + 1):
                if not self.is_running:
                    break

                self.log(f"\n{'='*50}")
                self.log(f"ITERATION {iteration}/{max_iter}")
                self.log(f"{'='*50}")

                iter_folder = os.path.join(base_folder, f"iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                self.update_progress((iteration/max_iter)*100, f"Iter {iteration}/{max_iter}")

                # Run Nastran and get RF
                result = self._run_iteration(iter_folder, iteration, target_rf)

                if not result:
                    self.log("  ERROR: Iteration failed!")
                    continue

                self.iteration_results.append(result)
                min_rf = result['min_rf']
                weight = result['weight']
                n_fail = result['n_fail']

                self.log(f"\n  Summary: Min RF={min_rf:.4f}, Fails={n_fail}, Weight={weight:.6f}t")

                # Calculate per-property RF
                rf_details = result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                # Check convergence
                all_converged = True
                for pid, rf in pid_min_rf.items():
                    if rf < target_rf - rf_tol:
                        all_converged = False
                        break

                if all_converged and min_rf >= target_rf - rf_tol:
                    self.log(f"\n  *** CONVERGED! All properties have RF ≥ {target_rf - rf_tol:.3f} ***")
                    converged = True
                    if weight < best_weight:
                        best_weight = weight
                        self.best_solution = result
                    break

                # Track best feasible solution
                if min_rf >= target_rf - rf_tol and weight < best_weight:
                    best_weight = weight
                    self.best_solution = result
                    self.log(f"  *** NEW BEST: Weight={weight:.6f}t ***")

                # ========== WEIGHT-EFFICIENT UPDATE ==========
                # Calculate efficiency for each under-designed property
                # Efficiency = RF_deficit / weight_sensitivity (higher = better to increase)

                self.log(f"\n  Weight-Efficient Update Analysis:")

                candidates = []  # [(pid, type, efficiency, rf_deficit, current_t, max_t)]

                # Analyze BAR properties
                for pid in self.bar_properties:
                    prop_rf = pid_min_rf.get(pid, target_rf)
                    current_t = self.current_bar_thicknesses[pid]
                    sens = weight_sens.get(pid, 0)

                    if prop_rf < target_rf - rf_tol and current_t < bar_max and sens > 0:
                        rf_deficit = target_rf - prop_rf
                        # Efficiency: how much RF we need per unit weight cost
                        # Higher efficiency = this property gives more RF bang for weight buck
                        efficiency = rf_deficit / sens
                        candidates.append({
                            'pid': pid,
                            'type': 'bar',
                            'efficiency': efficiency,
                            'rf_deficit': rf_deficit,
                            'rf': prop_rf,
                            'current_t': current_t,
                            'max_t': bar_max,
                            'sens': sens
                        })

                # Analyze SKIN properties
                # KEY FIX: If skin has no RF data, use nearby bars' RF instead of target_rf
                for pid in self.skin_properties:
                    prop_rf = pid_min_rf.get(pid, None)  # None if no RF data
                    current_t = self.current_skin_thicknesses[pid]
                    sens = weight_sens.get(pid, 0)

                    # If skin has no direct RF data, estimate from nearby bars
                    if prop_rf is None:
                        nearby_bars = self.skin_to_nearby_bars.get(pid, [])
                        nearby_rfs = [pid_min_rf.get(b, None) for b in nearby_bars]
                        nearby_rfs = [rf for rf in nearby_rfs if rf is not None and rf < 999]
                        if nearby_rfs:
                            # Use minimum RF of nearby bars (most critical)
                            prop_rf = min(nearby_rfs)
                        else:
                            # No nearby bar data either - skip this skin
                            continue

                    if prop_rf < target_rf - rf_tol and current_t < skin_max and sens > 0:
                        rf_deficit = target_rf - prop_rf
                        efficiency = rf_deficit / sens
                        candidates.append({
                            'pid': pid,
                            'type': 'skin',
                            'efficiency': efficiency,
                            'rf_deficit': rf_deficit,
                            'rf': prop_rf,
                            'current_t': current_t,
                            'max_t': skin_max,
                            'sens': sens
                        })

                if not candidates:
                    self.log("    No properties need strengthening or all at max thickness.")
                    # Try reducing over-designed properties
                    self._reduce_overdesigned(pid_min_rf, target_rf, rf_tol,
                                              bar_min, bar_max, skin_min, skin_max)
                    continue

                # Sort by efficiency (descending - highest efficiency first)
                candidates.sort(key=lambda x: x['efficiency'], reverse=True)

                # Log top candidates
                self.log(f"    Top efficiency candidates (higher = better to increase):")
                for i, c in enumerate(candidates[:5]):
                    self.log(f"      {i+1}. PID {c['pid']} ({c['type']}): "
                             f"eff={c['efficiency']:.2e}, RF={c['rf']:.3f}, "
                             f"sens={c['sens']:.6f} t/mm")

                # Strategy: Increase properties with highest efficiency
                # But limit how many we increase per iteration for stability
                n_to_update = min(len(candidates), max(3, len(candidates) // 3))

                bar_increased = 0
                skin_increased = 0

                for c in candidates[:n_to_update]:
                    pid = c['pid']
                    current_t = c['current_t']
                    max_t = c['max_t']
                    rf_deficit = c['rf_deficit']

                    # Calculate increase amount based on stress ratio
                    # Use larger steps for higher efficiency properties
                    alpha = 0.5
                    if c['rf'] > 0:
                        ratio = (target_rf / c['rf']) ** alpha
                        ratio = min(ratio, 1.3)  # Max 30% increase
                    else:
                        ratio = 1.3

                    new_t = current_t * ratio

                    # Also ensure minimum step
                    if new_t - current_t < step * 0.5:
                        new_t = current_t + step * 0.5

                    new_t = min(new_t, max_t)

                    if c['type'] == 'bar':
                        self.current_bar_thicknesses[pid] = new_t
                        bar_increased += 1
                    else:
                        self.current_skin_thicknesses[pid] = new_t
                        skin_increased += 1

                self.log(f"    Updated: {bar_increased} bars, {skin_increased} skins (by efficiency)")

                # Also handle coupled skins for bars that were increased
                # KEY FIX: If bar fails, also increase nearby skins (they share load)
                coupled_updates = 0
                for c in candidates[:n_to_update]:
                    if c['type'] == 'bar':
                        bar_pid = c['pid']
                        bar_rf = c['rf']
                        nearby_skins = self.bar_to_nearby_skins.get(bar_pid, [])
                        for skin_pid in nearby_skins:
                            if skin_pid in self.current_skin_thicknesses:
                                current = self.current_skin_thicknesses[skin_pid]
                                if current >= skin_max:
                                    continue  # Already at max

                                # Get skin RF (use bar RF as proxy if no data)
                                skin_rf = pid_min_rf.get(skin_pid, None)
                                if skin_rf is None:
                                    skin_rf = bar_rf  # Use nearby bar's RF as estimate

                                # If skin or nearby bar is under-designed, increase skin
                                if skin_rf < target_rf - rf_tol or bar_rf < target_rf - rf_tol:
                                    # Scale increase by how bad the bar is failing
                                    # Worse bar RF = larger skin increase
                                    if bar_rf > 0:
                                        increase_factor = min(1.15, (target_rf / bar_rf) ** 0.3)
                                    else:
                                        increase_factor = 1.15

                                    new_t = min(current * increase_factor, skin_max)
                                    if new_t > current:
                                        self.current_skin_thicknesses[skin_pid] = new_t
                                        coupled_updates += 1

                if coupled_updates > 0:
                    self.log(f"    Coupled skin updates: {coupled_updates} (scaled by bar RF)")

                # Reduce over-designed properties
                reduced = self._reduce_overdesigned(pid_min_rf, target_rf, rf_tol,
                                                    bar_min, bar_max, skin_min, skin_max)
                if reduced > 0:
                    self.log(f"    Reduced over-designed: {reduced} properties")

            # ========== FINAL RESULTS ==========
            self.log("\n" + "="*70)
            self.log("WEIGHT-EFFICIENT BOTTOM-UP COMPLETE")
            self.log("="*70)

            if self.best_solution:
                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {self.best_solution['min_rf']:.4f}")
                self.log(f"  Weight: {self.best_solution['weight']:.6f} tonnes")
                self.log(f"  Found at iteration: {self.best_solution['iteration']}")
                self.log(f"  Converged: {'Yes' if converged else 'No'}")

                self._generate_final_report(base_folder, "Weight-Efficient Bottom-Up",
                    nastran_count=iteration,
                    extra_info={
                        "Strategy": "Weight-Sensitivity Based Optimization",
                        "Start": "Minimum thicknesses",
                        "Converged": "Yes" if converged else "No",
                        "Final Iterations": iteration
                    })

                self.root.after(0, lambda: self.result_summary.config(
                    text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}",
                    foreground="green"
                ))

                self.root.after(0, lambda: messagebox.showinfo(
                    "Weight-Efficient Complete",
                    f"Optimization Complete!\n\n"
                    f"Best RF: {self.best_solution['min_rf']:.4f}\n"
                    f"Best Weight: {self.best_solution['weight']:.6f} tonnes\n"
                    f"Iterations: {iteration}\n"
                    f"Converged: {'Yes' if converged else 'No'}"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_balanced_bottomup(self):
        """Algorithm 10: Balanced Bottom-Up - Equal priority for bars AND skins.

        Unlike Weight-Efficient, this algorithm:
        1. Updates ALL failing properties, not just most efficient ones
        2. Uses nearby bar RF for skins without direct RF data
        3. Guarantees minimum increase for any failing property

        Best for: Cases where skin thicknesses also need optimization.
        """
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: BALANCED BOTTOM-UP")
            self.log("="*70)
            self.log("Strategy: Equal priority for bars AND skins")
            self.log("="*70)

            max_iter = int(self.max_iterations.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())

            # Calculate bar-skin proximity mapping
            self.calculate_bar_skin_proximity()

            # Create output folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"balanced_bottomup_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== PHASE 1: Initialize at MINIMUM thicknesses ==========
            self.log("\n" + "="*50)
            self.log("PHASE 1: Initialize at MINIMUM thicknesses")
            self.log("="*50)

            for pid in self.bar_properties:
                self.current_bar_thicknesses[pid] = bar_min

            for pid in self.skin_properties:
                self.current_skin_thicknesses[pid] = skin_min

            self.log(f"  Bar: {len(self.bar_properties)} properties → {bar_min} mm")
            self.log(f"  Skin: {len(self.skin_properties)} properties → {skin_min} mm")

            # ========== PHASE 2: Iterative Optimization ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: Balanced Iterative Optimization")
            self.log("="*50)

            best_weight = float('inf')
            converged = False

            for iteration in range(1, max_iter + 1):
                if not self.is_running:
                    break

                self.log(f"\n{'='*50}")
                self.log(f"ITERATION {iteration}/{max_iter}")
                self.log(f"{'='*50}")

                iter_folder = os.path.join(base_folder, f"iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                self.update_progress((iteration/max_iter)*100, f"Iter {iteration}/{max_iter}")

                # Run Nastran
                result = self._run_iteration(iter_folder, iteration, target_rf)

                if not result:
                    self.log("  ERROR: Iteration failed!")
                    continue

                self.iteration_results.append(result)
                min_rf = result['min_rf']
                weight = result['weight']
                n_fail = result['n_fail']

                # Calculate per-property RF
                rf_details = result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid and rf is not None:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                self.log(f"\n  Results: Min RF={min_rf:.4f}, Fails={n_fail}, Weight={weight:.6f}t")

                # Count failing bars and skins separately
                failing_bars = sum(1 for p in self.bar_properties if pid_min_rf.get(p, target_rf) < target_rf - rf_tol)
                failing_skins = sum(1 for p in self.skin_properties if pid_min_rf.get(p, target_rf) < target_rf - rf_tol)
                self.log(f"  Failing: {failing_bars} bars, {failing_skins} skins")

                # Check convergence
                if min_rf >= target_rf - rf_tol and n_fail == 0:
                    self.log(f"\n  *** CONVERGED! Min RF ≥ {target_rf - rf_tol:.3f} ***")
                    converged = True
                    if weight < best_weight:
                        best_weight = weight
                        self.best_solution = result
                    break

                # Track best feasible
                if min_rf >= target_rf - rf_tol and weight < best_weight:
                    best_weight = weight
                    self.best_solution = result
                    self.log(f"  *** NEW BEST: Weight={weight:.6f}t ***")

                # ========== BALANCED UPDATE ==========
                self.log(f"\n  Balanced Update:")

                bar_updated = 0
                skin_updated = 0

                # Update ALL failing BAR properties
                for pid in self.bar_properties:
                    prop_rf = pid_min_rf.get(pid, target_rf)
                    current_t = self.current_bar_thicknesses[pid]

                    if prop_rf < target_rf - rf_tol and current_t < bar_max:
                        # Calculate increase based on RF deficit
                        if prop_rf > 0:
                            ratio = (target_rf / prop_rf) ** 0.5
                            ratio = min(ratio, 1.5)  # Max 50% increase
                        else:
                            ratio = 1.5

                        new_t = current_t * ratio

                        # Ensure minimum step
                        if new_t - current_t < step:
                            new_t = current_t + step

                        new_t = min(new_t, bar_max)
                        self.current_bar_thicknesses[pid] = new_t
                        bar_updated += 1

                # Update SKIN properties - with special handling for no-RF cases
                for pid in self.skin_properties:
                    current_t = self.current_skin_thicknesses[pid]

                    if current_t >= skin_max:
                        continue  # Already at max

                    # Get skin RF (direct or estimated)
                    prop_rf = pid_min_rf.get(pid, None)

                    # If no direct RF, estimate from nearby bars
                    if prop_rf is None:
                        nearby_bars = self.skin_to_nearby_bars.get(pid, set())
                        nearby_rfs = [pid_min_rf.get(b, None) for b in nearby_bars]
                        nearby_rfs = [rf for rf in nearby_rfs if rf is not None and 0 < rf < 999]
                        if nearby_rfs:
                            prop_rf = min(nearby_rfs)  # Use most critical nearby bar
                        else:
                            prop_rf = target_rf  # No data - assume OK

                    # Update if failing
                    if prop_rf < target_rf - rf_tol:
                        if prop_rf > 0:
                            ratio = (target_rf / prop_rf) ** 0.5
                            ratio = min(ratio, 1.5)
                        else:
                            ratio = 1.5

                        new_t = current_t * ratio

                        # Ensure minimum step
                        if new_t - current_t < step:
                            new_t = current_t + step

                        new_t = min(new_t, skin_max)
                        self.current_skin_thicknesses[pid] = new_t
                        skin_updated += 1

                self.log(f"    Updated: {bar_updated} bars, {skin_updated} skins")

                # Log thickness ranges
                bar_ts = list(self.current_bar_thicknesses.values())
                skin_ts = list(self.current_skin_thicknesses.values())
                if bar_ts:
                    self.log(f"    Bar thickness range: {min(bar_ts):.2f} - {max(bar_ts):.2f} mm")
                if skin_ts:
                    self.log(f"    Skin thickness range: {min(skin_ts):.2f} - {max(skin_ts):.2f} mm")

                # Reduce over-designed
                reduced = self._reduce_overdesigned(pid_min_rf, target_rf, rf_tol,
                                                    bar_min, bar_max, skin_min, skin_max)
                if reduced > 0:
                    self.log(f"    Reduced over-designed: {reduced}")

            # ========== FINAL RESULTS ==========
            self.log("\n" + "="*70)
            self.log("BALANCED BOTTOM-UP COMPLETE")
            self.log("="*70)

            if self.best_solution:
                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {self.best_solution['min_rf']:.4f}")
                self.log(f"  Weight: {self.best_solution['weight']:.6f} tonnes")
                self.log(f"  Converged: {'Yes' if converged else 'No'}")

                # Log final thickness summary
                bar_ts = list(self.best_solution['bar_thicknesses'].values())
                skin_ts = list(self.best_solution.get('skin_thicknesses', {}).values())
                if bar_ts:
                    self.log(f"  Bar thicknesses: {min(bar_ts):.2f} - {max(bar_ts):.2f} mm (avg: {sum(bar_ts)/len(bar_ts):.2f})")
                if skin_ts:
                    self.log(f"  Skin thicknesses: {min(skin_ts):.2f} - {max(skin_ts):.2f} mm (avg: {sum(skin_ts)/len(skin_ts):.2f})")

                self._generate_final_report(base_folder, "Balanced Bottom-Up",
                    nastran_count=iteration,
                    extra_info={
                        "Strategy": "Equal priority for bars AND skins",
                        "Start": "Minimum thicknesses",
                        "Converged": "Yes" if converged else "No"
                    })

                self.root.after(0, lambda: self.result_summary.config(
                    text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}",
                    foreground="green"
                ))

                self.root.after(0, lambda: messagebox.showinfo(
                    "Balanced Bottom-Up Complete",
                    f"Optimization Complete!\n\n"
                    f"Best RF: {self.best_solution['min_rf']:.4f}\n"
                    f"Best Weight: {self.best_solution['weight']:.6f} tonnes\n"
                    f"Converged: {'Yes' if converged else 'No'}"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_decoupled_min_weight(self):
        """Algorithm: Decoupled Minimum Weight - Bars first, then only failing skins.

        This algorithm achieves MINIMUM WEIGHT by:
        1. Phase 1: Optimize ONLY bars (skins stay at minimum) until bars converge
        2. Phase 2: Run analysis and update ONLY skins that have their OWN RF < target
        3. NO coupled updates - each property is judged by its OWN RF, not neighbors

        Why this works:
        - Bars have much lower weight sensitivity (kg/mm) than skins
        - By optimizing bars first, we maximize structural efficiency
        - Skins are only increased when THEY actually fail, not when nearby bars fail
        - Avoids the "early skin thickening" problem of coupled algorithms
        """
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: DECOUPLED MINIMUM WEIGHT")
            self.log("="*70)
            self.log("Strategy: Bars first (cheap), then only failing skins (expensive)")
            self.log("Key: NO coupled updates - each property judged by its OWN RF")
            self.log("="*70)

            max_iter = int(self.max_iterations.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())
            use_gfem = self.use_gfem_thickness.get()
            keep_bars_const = self.keep_bars_constant.get()
            keep_skins_const = self.keep_skins_constant.get()

            # Create output folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"decoupled_minweight_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== PHASE 1: Initialize at MINIMUM/GFEM thicknesses ==========
            self.log("\n" + "="*50)
            if use_gfem:
                self.log("INITIALIZATION: All properties at GFEM values")
                self.log("(Using Excel thickness as initial and minimum)")
            else:
                self.log("INITIALIZATION: All properties at MINIMUM")
            if keep_bars_const:
                self.log("*** BAR thicknesses CONSTANT (from Excel) - only skins will be optimized ***")
            if keep_skins_const:
                self.log("*** SKIN thicknesses CONSTANT (from Excel) - only bars will be optimized ***")
            self.log("="*50)

            for pid in self.bar_properties:
                if keep_bars_const and pid in self.gfem_bar_thicknesses:
                    self.current_bar_thicknesses[pid] = self.gfem_bar_thicknesses[pid]
                elif use_gfem and pid in self.gfem_bar_thicknesses:
                    self.current_bar_thicknesses[pid] = self.gfem_bar_thicknesses[pid]
                else:
                    self.current_bar_thicknesses[pid] = bar_min

            for pid in self.skin_properties:
                if keep_skins_const and pid in self.gfem_skin_thicknesses:
                    self.current_skin_thicknesses[pid] = self.gfem_skin_thicknesses[pid]
                elif use_gfem and pid in self.gfem_skin_thicknesses:
                    self.current_skin_thicknesses[pid] = self.gfem_skin_thicknesses[pid]
                else:
                    self.current_skin_thicknesses[pid] = skin_min

            if keep_bars_const:
                avg_bar = sum(self.current_bar_thicknesses.values()) / len(self.current_bar_thicknesses) if self.current_bar_thicknesses else 0
                self.log(f"  Bars: {len(self.bar_properties)} properties -> CONSTANT (avg: {avg_bar:.2f} mm)")
            elif use_gfem:
                avg_bar = sum(self.gfem_bar_thicknesses.values()) / len(self.gfem_bar_thicknesses) if self.gfem_bar_thicknesses else 0
                self.log(f"  Bars: {len(self.bar_properties)} properties -> GFEM (avg: {avg_bar:.2f} mm)")
            else:
                self.log(f"  Bars: {len(self.bar_properties)} properties -> {bar_min} mm")

            if keep_skins_const:
                avg_skin = sum(self.current_skin_thicknesses.values()) / len(self.current_skin_thicknesses) if self.current_skin_thicknesses else 0
                self.log(f"  Skins: {len(self.skin_properties)} properties -> CONSTANT (avg: {avg_skin:.2f} mm)")
            elif use_gfem:
                avg_skin = sum(self.gfem_skin_thicknesses.values()) / len(self.gfem_skin_thicknesses) if self.gfem_skin_thicknesses else 0
                self.log(f"  Skins: {len(self.skin_properties)} properties -> GFEM (avg: {avg_skin:.2f} mm, LOCKED in Phase 1)")
            else:
                self.log(f"  Skins: {len(self.skin_properties)} properties -> {skin_min} mm (LOCKED in Phase 1)")

            # ========== PHASE 1: BAR-ONLY OPTIMIZATION ==========
            self.log("\n" + "="*50)
            self.log("PHASE 1: BAR-ONLY OPTIMIZATION")
            self.log("="*50)
            self.log("Skins are LOCKED at minimum - only bars will be adjusted")

            best_weight = float('inf')
            bars_converged = False
            phase1_iterations = 0
            max_phase1_iter = max(10, max_iter // 2)  # At least 10, or half of max_iter

            for iteration in range(1, max_phase1_iter + 1):
                if not self.is_running:
                    break

                phase1_iterations = iteration
                self.log(f"\n{'='*50}")
                self.log(f"PHASE 1 - ITERATION {iteration}/{max_phase1_iter}")
                self.log(f"{'='*50}")

                iter_folder = os.path.join(base_folder, f"phase1_iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                self.update_progress((iteration/max_iter)*50, f"Phase 1: Bars only - Iter {iteration}")

                # Run Nastran
                result = self._run_iteration(iter_folder, iteration, target_rf)

                if not result:
                    self.log("  ERROR: Iteration failed!")
                    continue

                self.iteration_results.append(result)
                min_rf = result['min_rf']
                weight = result['weight']

                # Get per-property RF
                rf_details = result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                # Count bar convergence (how many bars have RF >= target)
                bars_ok = 0
                bars_total = 0
                bars_failing = []
                for pid in self.bar_properties:
                    bar_rf = pid_min_rf.get(pid, None)
                    if bar_rf is not None:
                        bars_total += 1
                        if bar_rf >= target_rf - rf_tol:
                            bars_ok += 1
                        else:
                            bars_failing.append((pid, bar_rf))

                bar_convergence_pct = (bars_ok / bars_total * 100) if bars_total > 0 else 0
                self.log(f"\n  BAR STATUS: {bars_ok}/{bars_total} converged ({bar_convergence_pct:.1f}%)")
                self.log(f"  Min RF: {min_rf:.4f}, Weight: {weight:.6f}t")

                # Check if bars are converged (90% or more at target)
                if bar_convergence_pct >= 90:
                    self.log(f"\n  *** BARS CONVERGED! {bar_convergence_pct:.1f}% at RF >= {target_rf - rf_tol:.3f} ***")
                    bars_converged = True
                    if weight < best_weight:
                        best_weight = weight
                        self.best_solution = result
                    break

                # Update ONLY BAR thicknesses (skins stay locked)
                if keep_bars_const:
                    self.log(f"\n  Bars: CONSTANT (locked at Excel values) - skipping Phase 1 updates")
                    bars_converged = True
                    break
                else:
                    self.log(f"\n  Updating BAR thicknesses only...")
                    bars_increased = 0
                    bars_decreased = 0

                    alpha = 0.5  # Stress ratio exponent

                    for pid in self.bar_properties:
                        old_t = self.current_bar_thicknesses[pid]
                        bar_rf = pid_min_rf.get(pid, None)

                        if bar_rf is None:
                            continue  # No data for this bar

                        if bar_rf < target_rf - rf_tol:
                            # Under-designed: INCREASE
                            ratio = (target_rf / bar_rf) ** alpha
                            ratio = min(ratio, 1.3)  # Max 30% increase
                            new_t = old_t * ratio
                            bars_increased += 1
                        elif bar_rf > target_rf + rf_tol * 2:
                            # Over-designed: REDUCE (more conservative)
                            ratio = (target_rf / bar_rf) ** alpha
                            ratio = max(ratio, 0.85)  # Max 15% reduction
                            new_t = old_t * ratio
                            bars_decreased += 1
                        else:
                            new_t = old_t  # Within tolerance

                        # Use GFEM thickness as minimum if enabled
                        if use_gfem and pid in self.gfem_bar_thicknesses:
                            pid_min = self.gfem_bar_thicknesses[pid]
                        else:
                            pid_min = bar_min
                        new_t = max(pid_min, min(bar_max, new_t))
                        self.current_bar_thicknesses[pid] = new_t

                    self.log(f"    Bars: {bars_increased} increased, {bars_decreased} decreased")
                    if keep_skins_const:
                        self.log(f"    Skins: CONSTANT (locked at Excel values)")
                    elif use_gfem:
                        self.log(f"    Skins: LOCKED at GFEM values (no changes)")
                    else:
                        self.log(f"    Skins: LOCKED at {skin_min} mm (no changes)")

            # ========== PHASE 2: SKIN-ONLY OPTIMIZATION (DECOUPLED) ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: SKIN-ONLY OPTIMIZATION (DECOUPLED)")
            self.log("="*50)
            self.log("Bars are now LOCKED - only skins with their OWN RF < target will be increased")
            self.log("KEY: NO coupled updates - each skin judged by its OWN stress, not neighbor bars")

            phase2_iterations = 0
            max_phase2_iter = max_iter - phase1_iterations
            skins_converged = False

            for iteration in range(1, max_phase2_iter + 1):
                if not self.is_running:
                    break

                phase2_iterations = iteration
                total_iter = phase1_iterations + iteration

                self.log(f"\n{'='*50}")
                self.log(f"PHASE 2 - ITERATION {iteration}/{max_phase2_iter} (Total: {total_iter})")
                self.log(f"{'='*50}")

                iter_folder = os.path.join(base_folder, f"phase2_iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                self.update_progress(50 + (iteration/max_phase2_iter)*50, f"Phase 2: Skins - Iter {iteration}")

                # Run Nastran
                result = self._run_iteration(iter_folder, total_iter, target_rf)

                if not result:
                    self.log("  ERROR: Iteration failed!")
                    continue

                self.iteration_results.append(result)
                min_rf = result['min_rf']
                weight = result['weight']

                # Get per-property RF
                rf_details = result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                # Count skin status - ONLY skins with THEIR OWN RF data
                skins_ok = 0
                skins_total_with_data = 0
                skins_failing = []
                skins_no_data = 0

                for pid in self.skin_properties:
                    skin_rf = pid_min_rf.get(pid, None)
                    if skin_rf is not None:
                        skins_total_with_data += 1
                        if skin_rf >= target_rf - rf_tol:
                            skins_ok += 1
                        else:
                            skins_failing.append((pid, skin_rf))
                    else:
                        skins_no_data += 1

                if skins_total_with_data > 0:
                    skin_convergence_pct = (skins_ok / skins_total_with_data * 100)
                else:
                    skin_convergence_pct = 100  # No data = assume OK

                self.log(f"\n  SKIN STATUS (DECOUPLED - own RF only):")
                self.log(f"    With RF data: {skins_total_with_data} ({skins_ok} OK, {len(skins_failing)} failing)")
                self.log(f"    Without RF data: {skins_no_data} (assumed OK - not increased)")
                self.log(f"    Convergence: {skin_convergence_pct:.1f}%")
                self.log(f"  Overall: Min RF={min_rf:.4f}, Weight={weight:.6f}t")

                # Track best solution
                if min_rf >= target_rf - rf_tol and weight < best_weight:
                    best_weight = weight
                    self.best_solution = result
                    self.log(f"  *** NEW BEST: Weight={weight:.6f}t ***")

                # Check convergence
                if min_rf >= target_rf - rf_tol:
                    self.log(f"\n  *** FULLY CONVERGED! All RF >= {target_rf - rf_tol:.3f} ***")
                    skins_converged = True
                    break

                # Update ONLY SKIN thicknesses that have their OWN RF failing
                # KEY DIFFERENCE: We do NOT use nearby bar RF - only skin's own RF
                skins_increased = 0

                if keep_skins_const:
                    self.log(f"\n  Skins: CONSTANT (locked at Excel values) - skipping skin updates")
                    self.log(f"    Bars: LOCKED (no changes in Phase 2)")
                else:
                    self.log(f"\n  Updating SKIN thicknesses (DECOUPLED mode)...")

                    alpha = 0.5

                    for pid in self.skin_properties:
                        old_t = self.current_skin_thicknesses[pid]
                        skin_rf = pid_min_rf.get(pid, None)

                        # KEY: If skin has no RF data, DO NOT increase it
                        # This is the DECOUPLED behavior - no guessing from neighbors
                        if skin_rf is None:
                            continue  # Skip - no data means we can't judge this skin

                        if skin_rf < target_rf - rf_tol:
                            # Skin's OWN RF is failing: INCREASE
                            ratio = (target_rf / skin_rf) ** alpha
                            ratio = min(ratio, 1.25)  # Max 25% increase for skins
                            new_t = old_t * ratio
                            skins_increased += 1
                            # Use GFEM thickness as minimum if enabled
                            if use_gfem and pid in self.gfem_skin_thicknesses:
                                spid_min = self.gfem_skin_thicknesses[pid]
                            else:
                                spid_min = skin_min
                            new_t = max(spid_min, min(skin_max, new_t))
                            self.current_skin_thicknesses[pid] = new_t
                        # Note: We don't decrease skins here - keep minimum weight

                    self.log(f"    Skins increased: {skins_increased} (only those with their OWN RF < {target_rf - rf_tol:.3f})")
                    self.log(f"    Skins skipped: {skins_no_data} (no RF data - NOT increased)")
                    self.log(f"    Bars: LOCKED (no changes in Phase 2)")

                # If no skins need updating and we still haven't converged,
                # the problem might be with bars that need more adjustment
                if skins_increased == 0 and not skins_converged and not keep_bars_const:
                    self.log(f"\n  Note: No skins updated, but min RF still < target")
                    self.log(f"  This may indicate bars need further adjustment...")

                    # Allow some bar adjustment in phase 2 if needed
                    bars_adjusted = 0
                    for pid in self.bar_properties:
                        bar_rf = pid_min_rf.get(pid, None)
                        if bar_rf is not None and bar_rf < target_rf - rf_tol:
                            old_t = self.current_bar_thicknesses[pid]
                            ratio = (target_rf / bar_rf) ** 0.4
                            ratio = min(ratio, 1.15)  # Smaller increase in phase 2
                            # Use GFEM thickness as minimum if enabled
                            if use_gfem and pid in self.gfem_bar_thicknesses:
                                pid_min = self.gfem_bar_thicknesses[pid]
                            else:
                                pid_min = bar_min
                            new_t = max(pid_min, min(old_t * ratio, bar_max))
                            self.current_bar_thicknesses[pid] = new_t
                            bars_adjusted += 1

                    if bars_adjusted > 0:
                        self.log(f"    Also adjusted {bars_adjusted} bars that were still failing")

            # ========== FINAL RESULTS ==========
            self.log("\n" + "="*70)
            self.log("DECOUPLED MINIMUM WEIGHT OPTIMIZATION COMPLETE")
            self.log("="*70)

            total_iterations = phase1_iterations + phase2_iterations

            if self.best_solution:
                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {self.best_solution['min_rf']:.4f}")
                self.log(f"  Weight: {self.best_solution['weight']:.6f} tonnes")
                self.log(f"  Phase 1 (bars): {phase1_iterations} iterations")
                self.log(f"  Phase 2 (skins): {phase2_iterations} iterations")
                self.log(f"  Total: {total_iterations} iterations")
                self.log(f"  Bars converged: {'Yes' if bars_converged else 'No'}")
                self.log(f"  Skins converged: {'Yes' if skins_converged else 'No'}")

                self._generate_final_report(base_folder, "Decoupled Minimum Weight",
                    nastran_count=total_iterations,
                    extra_info={
                        "Strategy": "Bars first, then only failing skins (no coupled updates)",
                        "Phase 1 (Bars)": f"{phase1_iterations} iterations",
                        "Phase 2 (Skins)": f"{phase2_iterations} iterations",
                        "Key Difference": "Skins updated only by their OWN RF, not neighbors"
                    })

                self.root.after(0, lambda: self.result_summary.config(
                    text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}",
                    foreground="green"
                ))

                self.root.after(0, lambda: messagebox.showinfo(
                    "Decoupled Min Weight Complete",
                    f"Optimization Complete!\n\n"
                    f"Best RF: {self.best_solution['min_rf']:.4f}\n"
                    f"Best Weight: {self.best_solution['weight']:.6f} tonnes\n"
                    f"Phase 1 (bars): {phase1_iterations} iters\n"
                    f"Phase 2 (skins): {phase2_iterations} iters\n"
                    f"Total: {total_iterations} iterations"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_coupled_efficiency_analysis(self):
        """Algorithm: Coupled Efficiency Analysis - Finds optimal bar-skin thickness combination.

        This algorithm analyzes the RF vs Weight trade-off between bar and skin:
        - Skin panel carries load from bar (coupled system)
        - When bar thickens, it takes more load, skin stress may decrease
        - When skin thickens, it takes more load, bar stress may decrease

        Strategy:
        1. Start at minimum thicknesses
        2. Each iteration: Test efficiency of bar vs skin increase
        3. Calculate: Efficiency = ΔRF / ΔWeight (RF gain per unit weight)
        4. Choose the more efficient option
        5. Build Pareto front (RF vs Weight)
        6. Find optimal combination where efficiency drops significantly

        Output:
        - Efficiency history for bars and skins
        - RF vs Weight Pareto front
        - Optimal point recommendation
        """
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: COUPLED EFFICIENCY ANALYSIS")
            self.log("="*70)
            self.log("Strategy: Compare bar vs skin efficiency (ΔRF/ΔWeight) each iteration")
            self.log("Goal: Find optimal balance between bar and skin thickness")
            self.log("="*70)

            max_iter = int(self.max_iterations.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())
            use_gfem = self.use_gfem_thickness.get()
            keep_bars_const = self.keep_bars_constant.get()
            keep_skins_const = self.keep_skins_constant.get()

            # Calculate bar-skin proximity mapping
            self.calculate_bar_skin_proximity()

            # Create output folder
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"coupled_efficiency_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== Data structures for analysis ==========
            efficiency_history = []  # Track efficiency per iteration
            pareto_front = []  # (weight, rf, bar_avg_t, skin_avg_t, iteration)
            bar_sensitivity_history = []  # ΔRF_bar / ΔWeight_bar
            skin_sensitivity_history = []  # ΔRF_skin / ΔWeight_skin

            # ========== PHASE 1: Initialize at MINIMUM/GFEM thicknesses ==========
            self.log("\n" + "="*50)
            if use_gfem:
                self.log("INITIALIZATION: All properties at GFEM values")
                self.log("(Using Excel thickness as initial and minimum)")
            else:
                self.log("INITIALIZATION: All properties at MINIMUM")
            if keep_bars_const:
                self.log("*** BAR thicknesses CONSTANT (from Excel) - only skins will be optimized ***")
            if keep_skins_const:
                self.log("*** SKIN thicknesses CONSTANT (from Excel) - only bars will be optimized ***")
            self.log("="*50)

            for pid in self.bar_properties:
                if keep_bars_const and pid in self.gfem_bar_thicknesses:
                    self.current_bar_thicknesses[pid] = self.gfem_bar_thicknesses[pid]
                elif use_gfem and pid in self.gfem_bar_thicknesses:
                    self.current_bar_thicknesses[pid] = self.gfem_bar_thicknesses[pid]
                else:
                    self.current_bar_thicknesses[pid] = bar_min
            for pid in self.skin_properties:
                if keep_skins_const and pid in self.gfem_skin_thicknesses:
                    self.current_skin_thicknesses[pid] = self.gfem_skin_thicknesses[pid]
                elif use_gfem and pid in self.gfem_skin_thicknesses:
                    self.current_skin_thicknesses[pid] = self.gfem_skin_thicknesses[pid]
                else:
                    self.current_skin_thicknesses[pid] = skin_min

            if keep_bars_const:
                avg_bar = sum(self.current_bar_thicknesses.values()) / len(self.current_bar_thicknesses) if self.current_bar_thicknesses else 0
                self.log(f"  Bars: {len(self.bar_properties)} properties -> CONSTANT (avg: {avg_bar:.2f} mm)")
            elif use_gfem:
                avg_bar = sum(self.gfem_bar_thicknesses.values()) / len(self.gfem_bar_thicknesses) if self.gfem_bar_thicknesses else 0
                self.log(f"  Bars: {len(self.bar_properties)} properties -> GFEM (avg: {avg_bar:.2f} mm)")
            else:
                self.log(f"  Bars: {len(self.bar_properties)} properties -> {bar_min} mm")

            if keep_skins_const:
                avg_skin = sum(self.current_skin_thicknesses.values()) / len(self.current_skin_thicknesses) if self.current_skin_thicknesses else 0
                self.log(f"  Skins: {len(self.skin_properties)} properties -> CONSTANT (avg: {avg_skin:.2f} mm)")
            elif use_gfem:
                avg_skin = sum(self.gfem_skin_thicknesses.values()) / len(self.gfem_skin_thicknesses) if self.gfem_skin_thicknesses else 0
                self.log(f"  Skins: {len(self.skin_properties)} properties -> GFEM (avg: {avg_skin:.2f} mm)")
            else:
                self.log(f"  Skins: {len(self.skin_properties)} properties -> {skin_min} mm")

            # Get initial baseline
            self.log("\n  Running initial baseline analysis...")
            iter_folder = os.path.join(base_folder, "iter_000_baseline")
            os.makedirs(iter_folder, exist_ok=True)

            baseline_result = self._run_iteration(iter_folder, 0, target_rf)
            if not baseline_result:
                self.log("  ERROR: Baseline analysis failed!")
                return

            self.iteration_results.append(baseline_result)
            prev_rf = baseline_result['min_rf']
            prev_weight = baseline_result['weight']

            pareto_front.append({
                'iteration': 0,
                'weight': prev_weight,
                'rf': prev_rf,
                'bar_avg_t': bar_min,
                'skin_avg_t': skin_min,
                'action': 'baseline'
            })

            self.log(f"\n  Baseline: RF={prev_rf:.4f}, Weight={prev_weight:.6f}t")

            # ========== PHASE 2: Iterative Efficiency Analysis ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: ITERATIVE EFFICIENCY ANALYSIS")
            self.log("="*50)

            best_weight = float('inf')
            converged = False

            # Track cumulative changes
            cumulative_bar_increase = 0.0
            cumulative_skin_increase = 0.0
            cumulative_bar_rf_gain = 0.0
            cumulative_skin_rf_gain = 0.0

            for iteration in range(1, max_iter + 1):
                if not self.is_running:
                    break

                self.log(f"\n{'='*60}")
                self.log(f"ITERATION {iteration}/{max_iter}")
                self.log(f"{'='*60}")

                self.update_progress((iteration/max_iter)*100, f"Iter {iteration}/{max_iter}")

                # Get per-property RF from previous result
                rf_details = baseline_result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                # ========== Identify failing properties ==========
                failing_bars = []
                failing_skins = []

                for pid in self.bar_properties:
                    bar_rf = pid_min_rf.get(pid, target_rf)
                    if bar_rf < target_rf - rf_tol:
                        failing_bars.append((pid, bar_rf))

                for pid in self.skin_properties:
                    skin_rf = pid_min_rf.get(pid, target_rf)
                    if skin_rf < target_rf - rf_tol:
                        failing_skins.append((pid, skin_rf))

                # Also consider skins near failing bars (coupled effect)
                for bar_pid, bar_rf in failing_bars:
                    if bar_pid in self.bar_to_nearby_skins:
                        for skin_pid in self.bar_to_nearby_skins[bar_pid]:
                            if skin_pid not in [s[0] for s in failing_skins]:
                                skin_rf = pid_min_rf.get(skin_pid, target_rf)
                                failing_skins.append((skin_pid, skin_rf))

                self.log(f"\n  Status: {len(failing_bars)} failing bars, {len(failing_skins)} failing/coupled skins")

                # ========== Check convergence ==========
                if len(failing_bars) == 0 and len(failing_skins) == 0:
                    if prev_rf >= target_rf - rf_tol:
                        self.log(f"\n  *** CONVERGED! All RF >= {target_rf - rf_tol:.3f} ***")
                        converged = True
                        if prev_weight < best_weight:
                            best_weight = prev_weight
                            self.best_solution = baseline_result
                        break

                # ========== Calculate weight sensitivities ==========
                weight_sens = self._calculate_weight_sensitivities()

                # ========== Estimate efficiency for BAR increase ==========
                bar_weight_delta = 0.0
                bar_thickness_increases = {}

                alpha = 0.5  # Stress ratio exponent

                for pid, bar_rf in failing_bars:
                    old_t = self.current_bar_thicknesses[pid]
                    if bar_rf > 0:
                        ratio = (target_rf / bar_rf) ** alpha
                        ratio = min(ratio, 1.3)  # Max 30% increase
                        delta_t = old_t * (ratio - 1)
                        delta_t = max(step * 0.5, min(delta_t, step * 2))  # Bound the change
                    else:
                        delta_t = step

                    new_t = min(old_t + delta_t, bar_max)
                    actual_delta = new_t - old_t

                    bar_thickness_increases[pid] = actual_delta
                    bar_weight_delta += weight_sens.get(pid, 0) * actual_delta

                # ========== Estimate efficiency for SKIN increase ==========
                skin_weight_delta = 0.0
                skin_thickness_increases = {}

                for pid, skin_rf in failing_skins:
                    old_t = self.current_skin_thicknesses[pid]
                    if skin_rf > 0 and skin_rf < target_rf:
                        ratio = (target_rf / skin_rf) ** (alpha * 0.7)  # Smaller exponent for skins
                        ratio = min(ratio, 1.2)  # Max 20% increase for skins
                        delta_t = old_t * (ratio - 1)
                        delta_t = max(step * 0.3, min(delta_t, step * 1.5))
                    else:
                        delta_t = step * 0.5

                    new_t = min(old_t + delta_t, skin_max)
                    actual_delta = new_t - old_t

                    skin_thickness_increases[pid] = actual_delta
                    skin_weight_delta += weight_sens.get(pid, 0) * actual_delta

                # ========== Decision: Which to prioritize? ==========
                # Calculate expected efficiency based on previous iterations

                # Use history to estimate RF gain
                if len(bar_sensitivity_history) > 0:
                    avg_bar_efficiency = sum(bar_sensitivity_history) / len(bar_sensitivity_history)
                else:
                    # Bars typically have better efficiency (less weight per RF gain)
                    avg_bar_efficiency = 0.1  # Initial estimate

                if len(skin_sensitivity_history) > 0:
                    avg_skin_efficiency = sum(skin_sensitivity_history) / len(skin_sensitivity_history)
                else:
                    avg_skin_efficiency = 0.05  # Initial estimate (skins are heavier)

                expected_bar_rf_gain = avg_bar_efficiency * bar_weight_delta if bar_weight_delta > 0 else 0
                expected_skin_rf_gain = avg_skin_efficiency * skin_weight_delta if skin_weight_delta > 0 else 0

                # Calculate efficiency ratio
                if bar_weight_delta > 0 and skin_weight_delta > 0:
                    bar_efficiency_ratio = expected_bar_rf_gain / bar_weight_delta
                    skin_efficiency_ratio = expected_skin_rf_gain / skin_weight_delta
                else:
                    bar_efficiency_ratio = 0
                    skin_efficiency_ratio = 0

                self.log(f"\n  Efficiency Analysis:")
                self.log(f"    Bar increase: ΔWeight={bar_weight_delta*1e6:.2f}g, Expected ΔRF={expected_bar_rf_gain:.4f}")
                self.log(f"    Skin increase: ΔWeight={skin_weight_delta*1e6:.2f}g, Expected ΔRF={expected_skin_rf_gain:.4f}")

                # ========== Apply updates based on efficiency ==========
                # Strategy: Apply BOTH but with weighted proportions
                # More efficient component gets larger updates

                total_efficiency = bar_efficiency_ratio + skin_efficiency_ratio
                if total_efficiency > 0:
                    bar_factor = 0.5 + 0.3 * (bar_efficiency_ratio / total_efficiency)  # 0.5-0.8
                    skin_factor = 0.5 + 0.3 * (skin_efficiency_ratio / total_efficiency)  # 0.5-0.8
                else:
                    bar_factor = 0.7  # Default: favor bars
                    skin_factor = 0.5

                self.log(f"    Update factors: Bar={bar_factor:.2f}, Skin={skin_factor:.2f}")

                # Apply BAR updates
                bars_updated = 0
                if keep_bars_const:
                    self.log(f"\n  Bars: CONSTANT (locked at Excel values) - skipping bar updates")
                else:
                    for pid, delta_t in bar_thickness_increases.items():
                        old_t = self.current_bar_thicknesses[pid]
                        # Use GFEM thickness as minimum if enabled
                        if use_gfem and pid in self.gfem_bar_thicknesses:
                            pid_min = self.gfem_bar_thicknesses[pid]
                        else:
                            pid_min = bar_min
                        new_t = max(pid_min, min(old_t + delta_t * bar_factor, bar_max))
                        if new_t > old_t:
                            self.current_bar_thicknesses[pid] = new_t
                            bars_updated += 1

                # Apply SKIN updates
                skins_updated = 0
                if keep_skins_const:
                    self.log(f"  Skins: CONSTANT (locked at Excel values) - skipping skin updates")
                else:
                    for pid, delta_t in skin_thickness_increases.items():
                        old_t = self.current_skin_thicknesses[pid]
                        # Use GFEM thickness as minimum if enabled
                        if use_gfem and pid in self.gfem_skin_thicknesses:
                            spid_min = self.gfem_skin_thicknesses[pid]
                        else:
                            spid_min = skin_min
                        new_t = max(spid_min, min(old_t + delta_t * skin_factor, skin_max))
                        if new_t > old_t:
                            self.current_skin_thicknesses[pid] = new_t
                            skins_updated += 1

                self.log(f"\n  Applied: {bars_updated} bars updated, {skins_updated} skins updated")

                # ========== Run analysis with new thicknesses ==========
                iter_folder = os.path.join(base_folder, f"iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)

                result = self._run_iteration(iter_folder, iteration, target_rf)

                if not result:
                    self.log("  ERROR: Iteration failed!")
                    continue

                self.iteration_results.append(result)
                curr_rf = result['min_rf']
                curr_weight = result['weight']

                # ========== Calculate actual efficiency ==========
                delta_rf = curr_rf - prev_rf
                delta_weight = curr_weight - prev_weight

                # Update sensitivity history
                actual_bar_efficiency = 0
                actual_skin_efficiency = 0

                if bar_weight_delta > 0 and bars_updated > 0:
                    # Estimate bar's contribution to RF change
                    bar_rf_contribution = delta_rf * (bar_weight_delta / (bar_weight_delta + skin_weight_delta + 1e-9))
                    actual_bar_efficiency = bar_rf_contribution / bar_weight_delta if bar_weight_delta > 0 else 0
                    bar_sensitivity_history.append(actual_bar_efficiency)
                    cumulative_bar_increase += bar_weight_delta
                    cumulative_bar_rf_gain += bar_rf_contribution

                if skin_weight_delta > 0 and skins_updated > 0:
                    # Estimate skin's contribution to RF change
                    skin_rf_contribution = delta_rf * (skin_weight_delta / (bar_weight_delta + skin_weight_delta + 1e-9))
                    actual_skin_efficiency = skin_rf_contribution / skin_weight_delta if skin_weight_delta > 0 else 0
                    skin_sensitivity_history.append(actual_skin_efficiency)
                    cumulative_skin_increase += skin_weight_delta
                    cumulative_skin_rf_gain += skin_rf_contribution

                # Calculate average thicknesses
                avg_bar_t = sum(self.current_bar_thicknesses.values()) / len(self.current_bar_thicknesses) if self.current_bar_thicknesses else 0
                avg_skin_t = sum(self.current_skin_thicknesses.values()) / len(self.current_skin_thicknesses) if self.current_skin_thicknesses else 0

                # Add to Pareto front
                pareto_front.append({
                    'iteration': iteration,
                    'weight': curr_weight,
                    'rf': curr_rf,
                    'bar_avg_t': avg_bar_t,
                    'skin_avg_t': avg_skin_t,
                    'delta_rf': delta_rf,
                    'delta_weight': delta_weight,
                    'efficiency': delta_rf / delta_weight if delta_weight > 0 else 0,
                    'action': f"bar_factor={bar_factor:.2f}, skin_factor={skin_factor:.2f}"
                })

                efficiency_history.append({
                    'iteration': iteration,
                    'bar_efficiency': actual_bar_efficiency if bar_weight_delta > 0 else 0,
                    'skin_efficiency': actual_skin_efficiency if skin_weight_delta > 0 else 0,
                    'overall_efficiency': delta_rf / delta_weight if delta_weight > 0 else 0
                })

                # Log results
                self.log(f"\n  Results:")
                self.log(f"    RF: {prev_rf:.4f} -> {curr_rf:.4f} (Δ={delta_rf:+.4f})")
                self.log(f"    Weight: {prev_weight:.6f}t -> {curr_weight:.6f}t (Δ={delta_weight*1e6:+.2f}g)")
                self.log(f"    Avg thickness: Bar={avg_bar_t:.2f}mm, Skin={avg_skin_t:.2f}mm")

                if delta_weight > 0:
                    overall_efficiency = delta_rf / delta_weight
                    self.log(f"    Overall Efficiency: {overall_efficiency:.4f} RF/tonne")

                # Track best solution
                if curr_rf >= target_rf - rf_tol and curr_weight < best_weight:
                    best_weight = curr_weight
                    self.best_solution = result
                    self.log(f"  *** NEW BEST: Weight={curr_weight:.6f}t ***")

                # Update for next iteration
                prev_rf = curr_rf
                prev_weight = curr_weight
                baseline_result = result

                # Check if efficiency is dropping (diminishing returns)
                if len(efficiency_history) >= 3:
                    recent_eff = [e['overall_efficiency'] for e in efficiency_history[-3:]]
                    if all(e <= 0.001 for e in recent_eff) and curr_rf >= target_rf - rf_tol:
                        self.log(f"\n  *** Diminishing returns detected - stopping early ***")
                        converged = True
                        break

            # ========== PHASE 3: Analysis and Report ==========
            self.log("\n" + "="*70)
            self.log("COUPLED EFFICIENCY ANALYSIS COMPLETE")
            self.log("="*70)

            # Summary statistics
            self.log("\n  EFFICIENCY SUMMARY:")
            self.log("  " + "-" * 50)

            if cumulative_bar_increase > 0:
                avg_bar_eff = cumulative_bar_rf_gain / cumulative_bar_increase
                self.log(f"    Bar: Total ΔWeight={cumulative_bar_increase*1e6:.2f}g, Total ΔRF={cumulative_bar_rf_gain:.4f}")
                self.log(f"          Average Efficiency: {avg_bar_eff:.4f} RF/tonne")
            else:
                avg_bar_eff = 0
                self.log(f"    Bar: No updates applied")

            if cumulative_skin_increase > 0:
                avg_skin_eff = cumulative_skin_rf_gain / cumulative_skin_increase
                self.log(f"    Skin: Total ΔWeight={cumulative_skin_increase*1e6:.2f}g, Total ΔRF={cumulative_skin_rf_gain:.4f}")
                self.log(f"          Average Efficiency: {avg_skin_eff:.4f} RF/tonne")
            else:
                avg_skin_eff = 0
                self.log(f"    Skin: No updates applied")

            # Determine which is more efficient
            self.log("\n  CONCLUSION:")
            if avg_bar_eff > avg_skin_eff * 1.2:
                self.log(f"    >>> BAR increases are {avg_bar_eff/avg_skin_eff:.1f}x more efficient than SKIN <<<")
                self.log(f"    Recommendation: Prioritize bar thickness for weight-efficient design")
            elif avg_skin_eff > avg_bar_eff * 1.2:
                self.log(f"    >>> SKIN increases are {avg_skin_eff/avg_bar_eff:.1f}x more efficient than BAR <<<")
                self.log(f"    Recommendation: Prioritize skin thickness for weight-efficient design")
            else:
                self.log(f"    >>> BAR and SKIN have similar efficiency <<<")
                self.log(f"    Recommendation: Balance both for optimal design")

            # Save Pareto front to CSV
            pareto_df = pd.DataFrame(pareto_front)
            pareto_df.to_csv(os.path.join(base_folder, "pareto_front.csv"), index=False)
            self.log(f"\n  Pareto front saved to: pareto_front.csv")

            # Save efficiency history
            eff_df = pd.DataFrame(efficiency_history)
            eff_df.to_csv(os.path.join(base_folder, "efficiency_history.csv"), index=False)
            self.log(f"  Efficiency history saved to: efficiency_history.csv")

            # Final results
            if self.best_solution:
                self.log(f"\n  BEST SOLUTION:")
                self.log(f"    Min RF: {self.best_solution['min_rf']:.4f}")
                self.log(f"    Weight: {self.best_solution['weight']:.6f} tonnes")
                self.log(f"    Iteration: {self.best_solution['iteration']}")
                self.log(f"    Converged: {'Yes' if converged else 'No'}")

                # Calculate final average thicknesses
                final_bar_avg = sum(self.current_bar_thicknesses.values()) / len(self.current_bar_thicknesses) if self.current_bar_thicknesses else 0
                final_skin_avg = sum(self.current_skin_thicknesses.values()) / len(self.current_skin_thicknesses) if self.current_skin_thicknesses else 0
                self.log(f"    Final Avg Bar: {final_bar_avg:.2f} mm")
                self.log(f"    Final Avg Skin: {final_skin_avg:.2f} mm")

                self._generate_final_report(base_folder, "Coupled Efficiency Analysis",
                    nastran_count=iteration + 1,
                    extra_info={
                        "Strategy": "RF vs Weight trade-off analysis",
                        "Bar Avg Efficiency": f"{avg_bar_eff:.4f} RF/tonne",
                        "Skin Avg Efficiency": f"{avg_skin_eff:.4f} RF/tonne",
                        "Final Avg Bar": f"{final_bar_avg:.2f} mm",
                        "Final Avg Skin": f"{final_skin_avg:.2f} mm",
                        "Converged": "Yes" if converged else "No"
                    })

                self.root.after(0, lambda: self.result_summary.config(
                    text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}",
                    foreground="green"
                ))

                self.root.after(0, lambda: messagebox.showinfo(
                    "Coupled Efficiency Analysis Complete",
                    f"Optimization Complete!\n\n"
                    f"Best RF: {self.best_solution['min_rf']:.4f}\n"
                    f"Best Weight: {self.best_solution['weight']:.6f} tonnes\n"
                    f"Iterations: {iteration}\n\n"
                    f"Bar Efficiency: {avg_bar_eff:.4f} RF/tonne\n"
                    f"Skin Efficiency: {avg_skin_eff:.4f} RF/tonne\n\n"
                    f"Results saved to:\n{base_folder}"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _run_adaptive_coupled(self):
        """Algorithm 11: Adaptive Coupled Optimization - Learns bar-skin interaction.

        This algorithm understands that bars and skins share load:
        - When bar thickness increases → bar takes more load → skin stress may decrease
        - The algorithm LEARNS this coupling from iteration history

        Strategy:
        1. Track RF changes for both bars AND skins after each update
        2. Calculate effective sensitivity: ΔRF / ΔWeight
        3. Adaptively choose: increase bars, skins, or both based on learned coupling

        Best for: Coupled bar-skin structures where optimal weight requires understanding
        the load sharing between components.
        """
        try:
            self.log("\n" + "="*70)
            self.log("ALGORITHM: ADAPTIVE COUPLED OPTIMIZATION")
            self.log("="*70)
            self.log("Strategy: Learn bar-skin coupling, optimize for minimum weight")
            self.log("="*70)

            max_iter = int(self.max_iterations.get())
            target_rf = float(self.target_rf.get())
            rf_tol = float(self.rf_tolerance.get())
            bar_min = float(self.bar_min_thickness.get())
            bar_max = float(self.bar_max_thickness.get())
            skin_min = float(self.skin_min_thickness.get())
            skin_max = float(self.skin_max_thickness.get())
            step = float(self.thickness_step.get())

            self.calculate_bar_skin_proximity()

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_folder = os.path.join(self.output_folder.get(), f"adaptive_coupled_{timestamp}")
            os.makedirs(base_folder, exist_ok=True)

            # ========== LEARNING PARAMETERS ==========
            # Track history to learn sensitivities
            history = []  # [{bar_delta_weight, skin_delta_weight, delta_min_rf, bar_rf_change, skin_rf_change}]

            # Learned sensitivities (updated each iteration)
            bar_rf_sensitivity = 1.0   # RF gain per unit bar weight
            skin_rf_sensitivity = 1.0  # RF gain per unit skin weight
            coupling_factor = 0.5      # How much bar changes affect skin RF

            # Calculate weight sensitivities (static - geometry based)
            weight_sens = self._calculate_weight_sensitivities()
            avg_bar_sens = sum(weight_sens.get(p, 0) for p in self.bar_properties) / max(len(self.bar_properties), 1)
            avg_skin_sens = sum(weight_sens.get(p, 0) for p in self.skin_properties) / max(len(self.skin_properties), 1)

            self.log(f"\n  Weight Sensitivities:")
            self.log(f"    Avg bar: {avg_bar_sens:.6f} t/mm")
            self.log(f"    Avg skin: {avg_skin_sens:.6f} t/mm")
            if avg_bar_sens > 0:
                self.log(f"    Skin/Bar ratio: {avg_skin_sens/avg_bar_sens:.1f}x")

            # ========== PHASE 1: Initialize ==========
            self.log("\n" + "="*50)
            self.log("PHASE 1: Initialize at MINIMUM")
            self.log("="*50)

            for pid in self.bar_properties:
                self.current_bar_thicknesses[pid] = bar_min
            for pid in self.skin_properties:
                self.current_skin_thicknesses[pid] = skin_min

            # ========== PHASE 2: Adaptive Iteration ==========
            self.log("\n" + "="*50)
            self.log("PHASE 2: Adaptive Coupled Optimization")
            self.log("="*50)

            best_weight = float('inf')
            converged = False
            prev_result = None
            prev_bar_weight = sum(self.current_bar_thicknesses.values())
            prev_skin_weight = sum(self.current_skin_thicknesses.values())

            for iteration in range(1, max_iter + 1):
                if not self.is_running:
                    break

                self.log(f"\n{'='*50}")
                self.log(f"ITERATION {iteration}/{max_iter}")
                self.log(f"{'='*50}")

                iter_folder = os.path.join(base_folder, f"iter_{iteration:03d}")
                os.makedirs(iter_folder, exist_ok=True)
                self.update_progress((iteration/max_iter)*100, f"Iter {iteration}/{max_iter}")

                # Run Nastran
                result = self._run_iteration(iter_folder, iteration, target_rf)
                if not result:
                    continue

                self.iteration_results.append(result)
                min_rf = result['min_rf']
                weight = result['weight']
                n_fail = result['n_fail']

                # Calculate per-property RF
                rf_details = result.get('rf_details', [])
                pid_min_rf = {}
                for d in rf_details:
                    pid = d.get('pid')
                    rf = d.get('rf', 0)
                    if pid and rf is not None:
                        if pid not in pid_min_rf or rf < pid_min_rf[pid]:
                            pid_min_rf[pid] = rf

                # Calculate average RF for bars and skins
                bar_rfs = [pid_min_rf.get(p, target_rf) for p in self.bar_properties if pid_min_rf.get(p, 999) < 999]
                skin_rfs = [pid_min_rf.get(p, target_rf) for p in self.skin_properties if pid_min_rf.get(p, 999) < 999]
                avg_bar_rf = sum(bar_rfs) / len(bar_rfs) if bar_rfs else target_rf
                avg_skin_rf = sum(skin_rfs) / len(skin_rfs) if skin_rfs else target_rf
                min_bar_rf = min(bar_rfs) if bar_rfs else target_rf
                min_skin_rf = min(skin_rfs) if skin_rfs else target_rf

                self.log(f"\n  Results:")
                self.log(f"    Min RF: {min_rf:.4f}, Fails: {n_fail}, Weight: {weight:.6f}t")
                self.log(f"    Bar RF:  min={min_bar_rf:.3f}, avg={avg_bar_rf:.3f}")
                self.log(f"    Skin RF: min={min_skin_rf:.3f}, avg={avg_skin_rf:.3f}")

                # ========== LEARN FROM HISTORY ==========
                if prev_result is not None:
                    curr_bar_weight = sum(self.current_bar_thicknesses.values())
                    curr_skin_weight = sum(self.current_skin_thicknesses.values())

                    delta_bar_weight = curr_bar_weight - prev_bar_weight
                    delta_skin_weight = curr_skin_weight - prev_skin_weight
                    delta_rf = min_rf - prev_result['min_rf']

                    # Record history
                    history.append({
                        'iter': iteration,
                        'delta_bar_weight': delta_bar_weight,
                        'delta_skin_weight': delta_skin_weight,
                        'delta_rf': delta_rf,
                        'min_rf': min_rf
                    })

                    # Update learned sensitivities (exponential moving average)
                    alpha = 0.3  # Learning rate
                    if abs(delta_bar_weight) > 0.01:
                        new_bar_sens = delta_rf / delta_bar_weight if delta_bar_weight != 0 else 0
                        bar_rf_sensitivity = alpha * new_bar_sens + (1 - alpha) * bar_rf_sensitivity

                    if abs(delta_skin_weight) > 0.01:
                        new_skin_sens = delta_rf / delta_skin_weight if delta_skin_weight != 0 else 0
                        skin_rf_sensitivity = alpha * new_skin_sens + (1 - alpha) * skin_rf_sensitivity

                    self.log(f"\n  Learned Sensitivities (ΔRF/ΔWeight):")
                    self.log(f"    Bar: {bar_rf_sensitivity:.4f}")
                    self.log(f"    Skin: {skin_rf_sensitivity:.4f}")

                    prev_bar_weight = curr_bar_weight
                    prev_skin_weight = curr_skin_weight

                # Check convergence
                if min_rf >= target_rf - rf_tol and n_fail == 0:
                    self.log(f"\n  *** CONVERGED! ***")
                    converged = True
                    if weight < best_weight:
                        best_weight = weight
                        self.best_solution = result
                    break

                if min_rf >= target_rf - rf_tol and weight < best_weight:
                    best_weight = weight
                    self.best_solution = result

                # ========== ADAPTIVE UPDATE STRATEGY ==========
                self.log(f"\n  Adaptive Update Strategy:")

                # Decide strategy based on learned sensitivities and current state
                # Effective efficiency = RF_sensitivity / Weight_sensitivity
                bar_efficiency = bar_rf_sensitivity / avg_bar_sens if avg_bar_sens > 0 else 0
                skin_efficiency = skin_rf_sensitivity / avg_skin_sens if avg_skin_sens > 0 else 0

                self.log(f"    Bar efficiency (ΔRF/ΔCost): {bar_efficiency:.4f}")
                self.log(f"    Skin efficiency (ΔRF/ΔCost): {skin_efficiency:.4f}")

                # Count what needs updating
                failing_bars = [(p, pid_min_rf.get(p, target_rf)) for p in self.bar_properties
                               if pid_min_rf.get(p, target_rf) < target_rf - rf_tol
                               and self.current_bar_thicknesses[p] < bar_max]
                failing_skins = [(p, pid_min_rf.get(p, target_rf)) for p in self.skin_properties
                                if pid_min_rf.get(p, target_rf) < target_rf - rf_tol
                                and self.current_skin_thicknesses[p] < skin_max]

                # Also check skins near failing bars (coupling effect)
                skins_near_failing_bars = set()
                for bar_pid, bar_rf in failing_bars:
                    for skin_pid in self.bar_to_nearby_skins.get(bar_pid, set()):
                        if self.current_skin_thicknesses.get(skin_pid, skin_max) < skin_max:
                            skins_near_failing_bars.add(skin_pid)

                self.log(f"    Failing bars: {len(failing_bars)}")
                self.log(f"    Failing skins: {len(failing_skins)}")
                self.log(f"    Skins near failing bars: {len(skins_near_failing_bars)}")

                # Adaptive strategy decision
                bar_updates = 0
                skin_updates = 0

                # Strategy 1: If bar efficiency >> skin efficiency, focus on bars
                # Strategy 2: If skin efficiency >> bar efficiency, focus on skins
                # Strategy 3: If similar, update both proportionally

                efficiency_ratio = bar_efficiency / skin_efficiency if skin_efficiency > 0.001 else 10

                if efficiency_ratio > 2:
                    strategy = "BAR_FOCUS"
                    bar_increase_factor = 1.3
                    skin_increase_factor = 1.05
                elif efficiency_ratio < 0.5:
                    strategy = "SKIN_FOCUS"
                    bar_increase_factor = 1.05
                    skin_increase_factor = 1.3
                else:
                    strategy = "BALANCED"
                    bar_increase_factor = 1.15
                    skin_increase_factor = 1.15

                self.log(f"    Strategy: {strategy} (efficiency ratio: {efficiency_ratio:.2f})")

                # Update BARS
                for bar_pid, bar_rf in failing_bars:
                    current_t = self.current_bar_thicknesses[bar_pid]
                    if bar_rf > 0:
                        rf_ratio = (target_rf / bar_rf) ** 0.4
                        rf_ratio = min(rf_ratio, bar_increase_factor)
                    else:
                        rf_ratio = bar_increase_factor

                    new_t = current_t * rf_ratio
                    if new_t - current_t < step:
                        new_t = current_t + step
                    new_t = min(new_t, bar_max)
                    self.current_bar_thicknesses[bar_pid] = new_t
                    bar_updates += 1

                # Update SKINS - both failing and those near failing bars
                skins_to_update = set(p for p, _ in failing_skins) | skins_near_failing_bars
                for skin_pid in skins_to_update:
                    current_t = self.current_skin_thicknesses.get(skin_pid, skin_min)
                    if current_t >= skin_max:
                        continue

                    # Get skin RF (direct or from nearby bars)
                    skin_rf = pid_min_rf.get(skin_pid, None)
                    if skin_rf is None:
                        nearby_bars = self.skin_to_nearby_bars.get(skin_pid, set())
                        nearby_rfs = [pid_min_rf.get(b, target_rf) for b in nearby_bars]
                        nearby_rfs = [rf for rf in nearby_rfs if rf < 999]
                        skin_rf = min(nearby_rfs) if nearby_rfs else target_rf

                    if skin_rf < target_rf - rf_tol:
                        if skin_rf > 0:
                            rf_ratio = (target_rf / skin_rf) ** 0.4
                            rf_ratio = min(rf_ratio, skin_increase_factor)
                        else:
                            rf_ratio = skin_increase_factor

                        new_t = current_t * rf_ratio
                        if new_t - current_t < step * 0.5:
                            new_t = current_t + step * 0.5
                        new_t = min(new_t, skin_max)
                        self.current_skin_thicknesses[skin_pid] = new_t
                        skin_updates += 1

                self.log(f"    Updated: {bar_updates} bars, {skin_updates} skins")

                # Log current state
                bar_ts = list(self.current_bar_thicknesses.values())
                skin_ts = list(self.current_skin_thicknesses.values())
                self.log(f"    Bar range: {min(bar_ts):.2f} - {max(bar_ts):.2f} mm")
                self.log(f"    Skin range: {min(skin_ts):.2f} - {max(skin_ts):.2f} mm")

                # Reduce over-designed
                reduced = self._reduce_overdesigned(pid_min_rf, target_rf, rf_tol,
                                                    bar_min, bar_max, skin_min, skin_max)
                if reduced > 0:
                    self.log(f"    Reduced over-designed: {reduced}")

                prev_result = result

            # ========== FINAL RESULTS ==========
            self.log("\n" + "="*70)
            self.log("ADAPTIVE COUPLED OPTIMIZATION COMPLETE")
            self.log("="*70)

            # Log learned parameters
            self.log(f"\nLearned Parameters:")
            self.log(f"  Final Bar RF Sensitivity: {bar_rf_sensitivity:.4f}")
            self.log(f"  Final Skin RF Sensitivity: {skin_rf_sensitivity:.4f}")

            if self.best_solution:
                self.log(f"\nBest Solution:")
                self.log(f"  Min RF: {self.best_solution['min_rf']:.4f}")
                self.log(f"  Weight: {self.best_solution['weight']:.6f} tonnes")
                self.log(f"  Converged: {'Yes' if converged else 'No'}")

                bar_ts = list(self.best_solution['bar_thicknesses'].values())
                skin_ts = list(self.best_solution.get('skin_thicknesses', {}).values())
                if bar_ts:
                    self.log(f"  Bar: {min(bar_ts):.2f} - {max(bar_ts):.2f} mm (avg: {sum(bar_ts)/len(bar_ts):.2f})")
                if skin_ts:
                    self.log(f"  Skin: {min(skin_ts):.2f} - {max(skin_ts):.2f} mm (avg: {sum(skin_ts)/len(skin_ts):.2f})")

                self._generate_final_report(base_folder, "Adaptive Coupled",
                    nastran_count=iteration,
                    extra_info={
                        "Strategy": "Learned bar-skin coupling",
                        "Final Bar Sensitivity": f"{bar_rf_sensitivity:.4f}",
                        "Final Skin Sensitivity": f"{skin_rf_sensitivity:.4f}",
                        "Converged": "Yes" if converged else "No"
                    })

                self.root.after(0, lambda: self.result_summary.config(
                    text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}",
                    foreground="green"
                ))

                self.root.after(0, lambda: messagebox.showinfo(
                    "Adaptive Coupled Complete",
                    f"Optimization Complete!\n\n"
                    f"Best RF: {self.best_solution['min_rf']:.4f}\n"
                    f"Best Weight: {self.best_solution['weight']:.6f} tonnes\n"
                    f"Converged: {'Yes' if converged else 'No'}\n\n"
                    f"Learned Sensitivities:\n"
                    f"  Bar: {bar_rf_sensitivity:.4f}\n"
                    f"  Skin: {skin_rf_sensitivity:.4f}"
                ))
            else:
                self.log("ERROR: No feasible solution found!")

        except Exception as e:
            self.log(f"ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: [self.btn_start.config(state=tk.NORMAL), self.btn_stop.config(state=tk.DISABLED)])

    def _reduce_overdesigned(self, pid_min_rf, target_rf, rf_tol, bar_min, bar_max, skin_min, skin_max):
        """Reduce thickness of over-designed properties to save weight."""
        reduced_count = 0
        reduce_threshold = target_rf + rf_tol + 0.1  # Only reduce if RF significantly above target

        for pid in self.bar_properties:
            prop_rf = pid_min_rf.get(pid, 0)
            if prop_rf > reduce_threshold:
                current = self.current_bar_thicknesses[pid]
                # Reduce proportionally but conservatively
                ratio = (target_rf / prop_rf) ** 0.3  # Small exponent for stability
                ratio = max(ratio, 0.85)  # Max 15% reduction
                new_t = max(current * ratio, bar_min)
                if new_t < current:
                    self.current_bar_thicknesses[pid] = new_t
                    reduced_count += 1

        for pid in self.skin_properties:
            prop_rf = pid_min_rf.get(pid, 0)
            if prop_rf > reduce_threshold:
                current = self.current_skin_thicknesses[pid]
                ratio = (target_rf / prop_rf) ** 0.3
                ratio = max(ratio, 0.85)
                new_t = max(current * ratio, skin_min)
                if new_t < current:
                    self.current_skin_thicknesses[pid] = new_t
                    reduced_count += 1

        return reduced_count

    def _latin_hypercube_sampling(self, n_samples, n_vars, bounds_low, bounds_high):
        """Generate Latin Hypercube Samples."""
        samples = []

        # Create intervals for each variable
        for i in range(n_samples):
            sample = []
            for j in range(n_vars):
                # Stratified random within interval
                low = bounds_low[j] + (bounds_high[j] - bounds_low[j]) * (i / n_samples)
                high = bounds_low[j] + (bounds_high[j] - bounds_low[j]) * ((i + 1) / n_samples)
                sample.append(random.uniform(low, high))
            samples.append(sample)

        # Shuffle each column independently for better space-filling
        samples = np.array(samples)
        for j in range(n_vars):
            np.random.shuffle(samples[:, j])

        return samples.tolist()

    def _fit_rsm(self, X, y, bounds_low, bounds_high):
        """Fit a 2nd order polynomial Response Surface Model."""
        try:
            n_samples, n_vars = X.shape

            # Normalize X to [-1, 1]
            X_norm = np.zeros_like(X)
            for j in range(n_vars):
                mid = (bounds_low[j] + bounds_high[j]) / 2
                half_range = (bounds_high[j] - bounds_low[j]) / 2
                X_norm[:, j] = (X[:, j] - mid) / half_range if half_range > 0 else 0

            # Build design matrix for quadratic model
            # [1, x1, x2, ..., x1^2, x2^2, ..., x1*x2, x1*x3, ...]
            design_matrix = [np.ones(n_samples)]  # Constant term

            # Linear terms
            for j in range(n_vars):
                design_matrix.append(X_norm[:, j])

            # Quadratic terms
            for j in range(n_vars):
                design_matrix.append(X_norm[:, j] ** 2)

            # Interaction terms (limit to avoid overfitting)
            if n_vars <= 10:
                for i, j in combinations(range(n_vars), 2):
                    design_matrix.append(X_norm[:, i] * X_norm[:, j])

            A = np.column_stack(design_matrix)

            # Solve least squares
            coeffs, residuals, rank, s = np.linalg.lstsq(A, y, rcond=None)

            # Calculate R²
            y_pred = A @ coeffs
            ss_res = np.sum((y - y_pred) ** 2)
            ss_tot = np.sum((y - np.mean(y)) ** 2)
            r2 = 1 - (ss_res / ss_tot) if ss_tot > 0 else 0

            return {
                'coeffs': coeffs,
                'n_vars': n_vars,
                'bounds_low': bounds_low,
                'bounds_high': bounds_high,
                'r2': r2
            }

        except Exception as e:
            self.log(f"RSM fitting error: {e}")
            return None

    def _rsm_predict(self, x, rsm, bounds_low, bounds_high):
        """Predict using RSM model."""
        n_vars = rsm['n_vars']
        coeffs = rsm['coeffs']

        # Normalize x to [-1, 1]
        x_norm = []
        for j in range(n_vars):
            mid = (bounds_low[j] + bounds_high[j]) / 2
            half_range = (bounds_high[j] - bounds_low[j]) / 2
            x_norm.append((x[j] - mid) / half_range if half_range > 0 else 0)

        # Build feature vector
        features = [1.0]  # Constant

        # Linear
        for j in range(n_vars):
            features.append(x_norm[j])

        # Quadratic
        for j in range(n_vars):
            features.append(x_norm[j] ** 2)

        # Interactions
        if n_vars <= 10:
            for i, j in combinations(range(n_vars), 2):
                features.append(x_norm[i] * x_norm[j])

        return np.dot(features, coeffs)

    def _run_fast_ga_internal(self):
        """Internal Fast GA without UI cleanup (for Hybrid)."""
        # Parameters
        bar_min = float(self.bar_min_thickness.get())
        bar_max = float(self.bar_max_thickness.get())
        skin_min = float(self.skin_min_thickness.get())
        skin_max = float(self.skin_max_thickness.get())
        target_rf = float(self.target_rf.get())
        rf_tol = float(self.rf_tolerance.get())

        pop_size = int(self.ga_population.get())
        n_generations = int(self.ga_generations.get())
        mutation_rate = float(self.ga_mutation_rate.get())
        crossover_rate = float(self.ga_crossover_rate.get())

        bar_pids = list(self.bar_properties.keys())
        skin_pids = list(self.skin_properties.keys())
        n_bars = len(bar_pids)
        n_genes = n_bars + len(skin_pids)

        if n_genes == 0:
            return

        self._init_reference_stresses(bar_pids, skin_pids, bar_min, skin_min)

        # Initialize population
        population = []
        for _ in range(pop_size):
            chromosome = []
            for pid in bar_pids:
                chromosome.append(random.uniform(bar_min, bar_max))
            for pid in skin_pids:
                chromosome.append(random.uniform(skin_min, skin_max))
            population.append(chromosome)

        best_fitness = float('inf')
        best_chromosome = None

        for gen in range(n_generations):
            if not self.is_running:
                break

            fitness_values = []
            for chromosome in population:
                fit = self._evaluate_surrogate_fitness(
                    chromosome, bar_pids, skin_pids,
                    bar_min, bar_max, skin_min, skin_max,
                    target_rf, rf_tol
                )
                fitness_values.append(fit)

            min_fit_idx = fitness_values.index(min(fitness_values))
            if fitness_values[min_fit_idx] < best_fitness:
                best_fitness = fitness_values[min_fit_idx]
                best_chromosome = population[min_fit_idx].copy()

            if gen % 20 == 0:
                self.update_progress((gen / n_generations) * 50, f"GA Gen {gen}/{n_generations}")
                self.log(f"  Gen {gen}: Fitness = {best_fitness:.6f}")

            # Evolution
            new_population = []
            while len(new_population) < pop_size:
                idx1, idx2 = random.sample(range(pop_size), 2)
                parent1 = population[idx1] if fitness_values[idx1] < fitness_values[idx2] else population[idx2]

                idx3, idx4 = random.sample(range(pop_size), 2)
                parent2 = population[idx3] if fitness_values[idx3] < fitness_values[idx4] else population[idx4]

                if random.random() < crossover_rate:
                    child = self._blx_crossover(parent1, parent2)
                else:
                    child = parent1.copy()

                child = self._gaussian_mutation(child, mutation_rate, n_bars,
                                                bar_min, bar_max, skin_min, skin_max)
                new_population.append(child)

            population = new_population

        # Apply best solution
        if best_chromosome:
            for i, pid in enumerate(bar_pids):
                self.current_bar_thicknesses[pid] = best_chromosome[i]
            for i, pid in enumerate(skin_pids):
                self.current_skin_thicknesses[pid] = best_chromosome[n_bars + i]

            self.log(f"\nFast GA Result: Fitness = {best_fitness:.6f}")

    def _run_iteration(self, folder, iteration, target_rf):
        try:
            # Handle multiple BDFs
            n_bdfs = len(self.bdf_models) if hasattr(self, 'bdf_models') and self.bdf_models else 1
            all_stresses = []

            # Step 0: Update maneuver BDF with current thicknesses for offset geometry
            updated_maneuver = None
            if self.maneuver_bdfs:
                maneuver_folder = os.path.join(folder, "maneuver_updated")
                os.makedirs(maneuver_folder, exist_ok=True)
                updated_maneuver = self._write_bdf_for_model(
                    maneuver_folder, None, self.maneuver_bdfs[0])
                self.log(f"  Maneuver BDF updated: {os.path.basename(updated_maneuver)}")

            if n_bdfs > 1:
                self.log(f"  Processing {n_bdfs} BDF files...")

            for bdf_idx, bdf_info in enumerate(self.bdf_models if n_bdfs > 1 else [{'model': self.bdf_model, 'path': self.bdf_paths[0] if self.bdf_paths else '', 'name': 'main.bdf'}]):
                bdf_name = bdf_info['name']
                bdf_subfolder = os.path.join(folder, f"bdf_{bdf_idx+1}_{os.path.splitext(bdf_name)[0]}") if n_bdfs > 1 else folder
                os.makedirs(bdf_subfolder, exist_ok=True)

                if n_bdfs > 1:
                    self.log(f"\n  --- BDF {bdf_idx+1}/{n_bdfs}: {bdf_name} ---")

                # 1. Write thermal BDF with current thicknesses
                self.log("  [1] Writing BDF...")
                bdf_path = self._write_bdf_for_model(bdf_subfolder, bdf_info['model'], bdf_info['path'])

                # 2. Apply offsets (from updated maneuver BDF)
                self.log("  [2] Applying offsets...")
                offset_bdf = self._apply_offsets(bdf_path, bdf_subfolder, updated_maneuver)

                # 3. Run Nastran
                self.log("  [3] Running Nastran...")
                self._run_nastran(offset_bdf or bdf_path, bdf_subfolder)

                # 4. Extract stresses from this BDF's OP2
                self.log("  [4] Extracting stresses...")
                stresses = self._extract_stresses(bdf_subfolder)
                all_stresses.extend(stresses)

                if stresses:
                    bar_stresses = [s for s in stresses if s['type'] == 'bar']
                    shell_stresses = [s for s in stresses if s['type'] == 'shell']
                    self.log(f"    Extracted: {len(bar_stresses)} bar, {len(shell_stresses)} shell stresses")

            # Combine all stresses from all BDFs
            stresses = all_stresses
            if not stresses:
                self.log("  WARNING: No stress data extracted! Check if OP2 files exist.")
                op2_files = [f for f in os.listdir(folder) if f.lower().endswith('.op2')]
                self.log(f"    OP2 files found: {op2_files}")
            else:
                total_bar = sum(1 for s in stresses if s['type'] == 'bar')
                total_shell = sum(1 for s in stresses if s['type'] == 'shell')
                self.log(f"\n  Total stresses from all BDFs: {total_bar} bar, {total_shell} shell")

            # 4.5 Combine stresses using Residual Strength (if loaded)
            combined_stresses = None
            if self.residual_strength_df is not None:
                self.log("  [4.5] Combining stresses...")
                combined_stresses = self._combine_stresses_multi(folder, n_bdfs)
                if combined_stresses:
                    self.log(f"    Combined: {len(combined_stresses)} stress combinations")

            # 5. Calculate RF
            self.log("  [5] Calculating RF...")
            rf_results = self._calculate_rf(stresses, target_rf, combined_stresses)

            # 6. Calculate weight
            self.log("  [6] Calculating weight...")
            weight = self._calculate_weight()

            min_rf = rf_results['min_rf']
            n_fail = rf_results['n_fail']

            self.log(f"\n  Results: Min RF={min_rf:.4f}, Fails={n_fail}, Weight={weight:.6f}t")

            result = {
                'iteration': iteration,
                'min_rf': min_rf,
                'n_fail': n_fail,
                'weight': weight,
                'folder': folder,
                'bar_thicknesses': self.current_bar_thicknesses.copy(),
                'skin_thicknesses': self.current_skin_thicknesses.copy(),
                'rf_details': rf_results['details'],
                'failing_pids': rf_results['failing_pids']
            }

            self._save_iteration(folder, result)
            return result

        except Exception as e:
            self.log(f"  ERROR: {e}")
            return None

    def _write_bdf(self, folder):
        """Write BDF with updated thicknesses - dim1 updated, dim2 kept original."""
        input_bdf = self.input_bdf_path.get()
        output_bdf = os.path.join(folder, "model.bdf")

        with open(input_bdf, 'r', encoding='latin-1') as f:
            lines = f.readlines()

        new_lines = []
        i = 0
        while i < len(lines):
            line = lines[i]

            if line.startswith('PBARL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in self.current_bar_thicknesses:
                        t = self.current_bar_thicknesses[pid]
                        # Ensure thickness is positive
                        t = max(0.1, t)
                        new_lines.append(line)
                        i += 1
                        # Process continuation lines
                        while i < len(lines) and (lines[i].startswith('+') or lines[i].startswith('*') or (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))):
                            cont = lines[i]
                            if cont.strip() and not cont.strip().startswith('$'):
                                # Parse original dim2 value (keep it)
                                try:
                                    # Original format: +name   DIM1    DIM2    ...
                                    # Field positions: 0-8=cont name, 8-16=DIM1, 16-24=DIM2
                                    original_dim2 = cont[16:24].strip()
                                    if original_dim2:
                                        dim2_val = float(original_dim2)
                                    else:
                                        dim2_val = t  # fallback
                                except:
                                    dim2_val = t  # fallback if parsing fails

                                # Write: continuation + new DIM1 + original DIM2 + rest
                                cont_name = cont[:8]
                                rest = cont[24:] if len(cont) > 24 else '\n'
                                new_cont = f"{cont_name}{t:8.4f}{dim2_val:8.4f}{rest}"
                                new_lines.append(new_cont)
                            else:
                                new_lines.append(cont)
                            i += 1
                        continue
                except Exception as e:
                    pass

            elif line.startswith('PSHELL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in self.current_skin_thicknesses:
                        t = self.current_skin_thicknesses[pid]
                        # Ensure thickness is positive
                        t = max(0.1, t)
                        # PSHELL format: PSHELL  PID     MID1    T       ...
                        # Field: 0-8=PSHELL, 8-16=PID, 16-24=MID1, 24-32=T
                        new_line = line[:24] + f"{t:8.4f}" + line[32:]
                        new_lines.append(new_line)
                        i += 1
                        continue
                except:
                    pass

            new_lines.append(line)
            i += 1

        with open(output_bdf, 'w', encoding='latin-1') as f:
            f.writelines(new_lines)

        self.log(f"    BDF written: {output_bdf}")
        return output_bdf

    def _write_bdf_for_model(self, folder, model, original_path):
        """Write BDF with updated thicknesses for a specific model (multi-BDF support)."""
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
                    if pid in self.current_bar_thicknesses:
                        t = self.current_bar_thicknesses[pid]
                        t = max(0.1, t)
                        new_lines.append(line)
                        i += 1
                        while i < len(lines) and (lines[i].startswith('+') or lines[i].startswith('*') or (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))):
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

            elif line.startswith('PSHELL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in self.current_skin_thicknesses:
                        t = self.current_skin_thicknesses[pid]
                        t = max(0.1, t)
                        prefix = line[:16]
                        mid1 = line[16:24]
                        rest = line[32:] if len(line) > 32 else '\n'
                        new_line = f"{prefix}{mid1}{t:8.4f}{rest}"
                        new_lines.append(new_line)
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

    def _apply_offsets(self, bdf_path, folder, maneuver_bdf_path=None):
        """Apply landing zoffset and bar WA/WB offsets to BDF."""
        if not self.landing_elem_ids and not self.bar_offset_elem_ids:
            self.log("    No offset elements defined, skipping offset application")
            return bdf_path

        try:
            # Use UPDATED maneuver BDF for offset geometry if provided
            if maneuver_bdf_path:
                geom_path = maneuver_bdf_path
                self.log(f"    Using UPDATED maneuver BDF for offsets: {os.path.basename(geom_path)}")
            else:
                geom_path = bdf_path
                self.log(f"    Using working BDF for offset geometry")
            self.log(f"    Applying offsets...")
            bdf = BDF(debug=False)
            bdf.read_bdf(geom_path, validate=False, xref=True, read_includes=True, encoding='latin-1')

            # Calculate landing (shell) offsets: zoffset = -t/2
            landing_offsets = {}
            landing_normals = {}

            for eid in self.landing_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                t = self.current_skin_thicknesses.get(pid, float(self.skin_min_thickness.get()))
                t = max(0.1, t)  # Ensure positive
                landing_offsets[eid] = -t / 2.0

                # Calculate normal vector for bar offset calculation
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    try:
                        nids = elem.node_ids[:3]
                        nodes = [bdf.nodes[n] for n in nids if n in bdf.nodes]
                        if len(nodes) >= 3:
                            p1, p2, p3 = [np.array(n.get_position()) for n in nodes]
                            normal = np.cross(p2 - p1, p3 - p1)
                            nl = np.linalg.norm(normal)
                            if nl > 1e-10:
                                landing_normals[eid] = normal / nl
                    except:
                        pass

            self.log(f"    Landing offsets calculated: {len(landing_offsets)} elements")

            # Build node -> shell mapping for bar offset calculation
            node_to_shells = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    for nid in elem.node_ids:
                        if nid not in node_to_shells:
                            node_to_shells[nid] = []
                        node_to_shells[nid].append(eid)

            # Calculate bar offsets: WA = WB = -normal * (landing_t + bar_t/2)
            bar_offsets = {}
            for eid in self.bar_offset_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                if elem.type not in ['CBAR', 'CBEAM']:
                    continue

                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                bar_t = self.current_bar_thicknesses.get(pid, float(self.bar_min_thickness.get()))
                bar_t = max(0.1, bar_t)  # Ensure positive

                bar_nodes = elem.node_ids[:2]
                if bar_nodes[0] in node_to_shells and bar_nodes[1] in node_to_shells:
                    common = set(node_to_shells[bar_nodes[0]]) & set(node_to_shells[bar_nodes[1]])
                    max_t = 0
                    best_normal = None

                    for shell_eid in common:
                        if shell_eid in landing_offsets:
                            shell_elem = bdf.elements[shell_eid]
                            shell_pid = shell_elem.pid if hasattr(shell_elem, 'pid') else None
                            if shell_pid:
                                shell_t = self.current_skin_thicknesses.get(shell_pid, 0)
                                if shell_t > max_t:
                                    max_t = shell_t
                                    if shell_eid in landing_normals:
                                        best_normal = landing_normals[shell_eid]

                    if best_normal is not None and max_t > 0:
                        # Offset = landing_thickness + bar_thickness/2
                        offset_mag = max_t + bar_t / 2.0
                        bar_offsets[eid] = tuple(-best_normal * offset_mag)

            self.log(f"    Bar offsets calculated: {len(bar_offsets)} elements")

            # Apply offsets to BDF file
            with open(bdf_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()

            def fmt(v, w=8):
                """Format value for Nastran field."""
                s = f"{v:.4f}"
                return s[:w].ljust(w) if len(s) <= w else f"{v:.2E}"[:w].ljust(w)

            new_lines = []
            i = 0
            applied_landing = 0
            applied_bar = 0

            while i < len(lines):
                line = lines[i]

                # CQUAD4: Apply zoffset at field 9 (position 64-72)
                if line.startswith('CQUAD4'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in landing_offsets:
                            # Ensure line is long enough
                            padded = line.rstrip().ljust(72)
                            new_line = padded[:64] + fmt(landing_offsets[eid]) + '\n'
                            new_lines.append(new_line)
                            applied_landing += 1
                            i += 1
                            continue
                    except:
                        pass

                # CTRIA3: Apply zoffset at field 7 (position 48-56)
                elif line.startswith('CTRIA3'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in landing_offsets:
                            padded = line.rstrip().ljust(56)
                            new_line = padded[:48] + fmt(landing_offsets[eid]) + '\n'
                            new_lines.append(new_line)
                            applied_landing += 1
                            i += 1
                            continue
                    except:
                        pass

                # CBAR: Apply WA, WB offset vectors as continuation
                elif line.startswith('CBAR'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in bar_offsets:
                            offset_vec = bar_offsets[eid]

                            # Check if next line is continuation (multi-line CBAR)
                            if i + 1 < len(lines) and (lines[i+1].startswith('+') or lines[i+1].startswith('*') or (lines[i+1][0] == ' ' and lines[i+1].strip())):
                                # Multi-line CBAR - modify existing continuation line
                                # Keep first 24 chars (cont marker + PA + PB), replace WA/WB
                                cont_line = lines[i+1]
                                # Ensure cont_line is at least 24 chars
                                if len(cont_line) < 24:
                                    cont_line = cont_line.rstrip().ljust(24)
                                new_cont = cont_line[:24]  # Keep +marker, PA, PB
                                new_cont += fmt(offset_vec[0])  # W1A (pos 24-32)
                                new_cont += fmt(offset_vec[1])  # W2A (pos 32-40)
                                new_cont += fmt(offset_vec[2])  # W3A (pos 40-48)
                                new_cont += fmt(offset_vec[0])  # W1B (pos 48-56)
                                new_cont += fmt(offset_vec[1])  # W2B (pos 56-64)
                                new_cont += fmt(offset_vec[2])  # W3B (pos 64-72)
                                new_cont += '\n'

                                new_lines.append(line)
                                new_lines.append(new_cont)
                                applied_bar += 1
                                i += 2
                                continue
                            else:
                                # Single line CBAR - add continuation for offsets
                                # Add continuation marker to main line
                                cont_name = '+CB' + str(eid)[-4:]
                                new_lines.append(line.rstrip() + cont_name + '\n')

                                # Create continuation line: +name, PA(blank), PB(blank), W1A, W2A, W3A, W1B, W2B, W3B
                                new_cont = cont_name.ljust(8)      # pos 0-7: continuation name
                                new_cont += '        '              # pos 8-15: PA (blank)
                                new_cont += '        '              # pos 16-23: PB (blank)
                                new_cont += fmt(offset_vec[0])     # pos 24-31: W1A
                                new_cont += fmt(offset_vec[1])     # pos 32-39: W2A
                                new_cont += fmt(offset_vec[2])     # pos 40-47: W3A
                                new_cont += fmt(offset_vec[0])     # pos 48-55: W1B
                                new_cont += fmt(offset_vec[1])     # pos 56-63: W2B
                                new_cont += fmt(offset_vec[2])     # pos 64-71: W3B
                                new_cont += '\n'
                                new_lines.append(new_cont)

                                applied_bar += 1
                                i += 1
                                continue
                    except:
                        pass

                new_lines.append(line)
                i += 1

            # Write as _offseted.bdf
            base, ext = os.path.splitext(bdf_path)
            output_bdf = base + "_offseted" + ext
            with open(output_bdf, 'w', encoding='latin-1') as f:
                f.writelines(new_lines)

            self.log(f"    Offsets applied: {applied_landing} landing, {applied_bar} bar")
            self.log(f"    Offset BDF written: {output_bdf}")

            return output_bdf

        except Exception as e:
            self.log(f"    Offset error: {e}")
            return bdf_path

    def _run_nastran(self, bdf_path, folder):
        nastran = self.nastran_path.get()
        if not nastran or not os.path.exists(nastran):
            return False
        try:
            cmd = f'"{nastran}" "{bdf_path}" out="{folder}" scratch=yes batch=no'
            proc = subprocess.Popen(cmd, shell=True)
            proc.wait(timeout=600)
            return True
        except:
            return False

    def _extract_stresses(self, folder):
        """Extract stresses from OP2 using cbar_force (RF Check Tool logic)."""
        results = []
        bar_stress_rows = []  # For bar_stress_results.csv

        for f in os.listdir(folder):
            if f.lower().endswith('.op2'):
                op2_name = f
                try:
                    op2 = OP2(debug=False)
                    op2.read_op2(os.path.join(folder, f))

                    # BAR STRESS from cbar_force (RF Check Tool exact logic)
                    if hasattr(op2, 'cbar_force') and op2.cbar_force:
                        for sc_id, force in op2.cbar_force.items():
                            for i, eid in enumerate(force.element):
                                # Axial force at index 6 (bending moment MA1)
                                axial = force.data[0, i, 6] if len(force.data.shape) == 3 else force.data[i, 6]
                                pid = self.elem_to_prop.get(int(eid))
                                d1 = d2 = area = stress = None

                                # Get dimensions: dim1 from current optimization, dim2 from BDF PBARL
                                if pid and pid in self.bar_properties:
                                    d1 = self.current_bar_thicknesses.get(pid, self.bar_properties[pid].get('dim1', 0))
                                    # Use BDF PBARL dim2 (original geometry), fallback to Excel
                                    if pid in self.pbarl_dims:
                                        d2 = self.pbarl_dims[pid]['dim2']
                                    else:
                                        d2 = self.bar_properties[pid].get('dim2', d1)
                                    area = d1 * d2
                                    if area > 0:
                                        stress = axial / area

                                results.append({
                                    'eid': int(eid), 'type': 'bar',
                                    'stress': float(stress) if stress else 0,
                                    'subcase': int(sc_id)
                                })

                                bar_stress_rows.append({
                                    'OP2': op2_name, 'Subcase': int(sc_id), 'Element': int(eid),
                                    'Property': pid, 'Axial': float(axial) if axial else 0,
                                    'Dim1': d1, 'Dim2': d2, 'Area': area,
                                    'Stress': float(stress) if stress else None
                                })

                    # SHELL STRESS - try multiple stress result types
                    shell_stress_count = 0

                    # List of shell stress attributes to try
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
                                        # Try to get von Mises stress (usually last column)
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
                                            shell_stress_count += 1
                                        except Exception:
                                            pass

                    # Log available stress attributes for debugging
                    if shell_stress_count == 0:
                        available_attrs = [attr for attr in dir(op2) if 'stress' in attr.lower() and not attr.startswith('_')]
                        self.log(f"    DEBUG: No shell stress found. Available: {available_attrs[:10]}")

                except Exception as e:
                    self.log(f"    OP2 read error: {e}")

        # Save bar_stress_results.csv
        if bar_stress_rows:
            csv_path = os.path.join(folder, 'bar_stress_results.csv')
            with open(csv_path, 'w', newline='') as f:
                w = csv.DictWriter(f, fieldnames=['OP2', 'Subcase', 'Element', 'Property', 'Axial', 'Dim1', 'Dim2', 'Area', 'Stress'])
                w.writeheader()
                w.writerows(bar_stress_rows)
            self.log(f"    Saved: bar_stress_results.csv ({len(bar_stress_rows)} rows)")

        return results

    def _combine_stresses(self, folder):
        """Combine stresses using Residual Strength table (RF Check Tool logic)."""
        if self.residual_strength_df is None or len(self.combination_table) == 0:
            self.log("    No Residual Strength data - skipping combination")
            return None

        stress_csv = os.path.join(folder, 'bar_stress_results.csv')
        if not os.path.exists(stress_csv):
            self.log("    bar_stress_results.csv not found")
            return None

        try:
            stress_df = pd.read_csv(stress_csv)

            # Build lookup: (subcase, element) -> stress
            lookup = {}
            for _, row in stress_df.iterrows():
                key = (int(row['Subcase']), int(row['Element']))
                lookup[key] = row['Stress'] if pd.notna(row['Stress']) else 0

            elements = stress_df['Element'].unique()

            rs_df = self.residual_strength_df
            cols = rs_df.columns.tolist()
            comb_col = cols[0]  # First column is Combined LC

            results = []
            for _, rs_row in rs_df.iterrows():
                comb_lc = rs_row[comb_col]
                if pd.isna(comb_lc):
                    continue
                comb_lc = int(comb_lc)

                for eid in elements:
                    total_stress = 0.0
                    components = []

                    for case_col, mult_col in self.combination_table:
                        case_id = rs_row[case_col]
                        multiplier = rs_row[mult_col]
                        if pd.isna(case_id) or pd.isna(multiplier):
                            continue
                        case_id = int(case_id)
                        multiplier = float(multiplier)

                        key = (case_id, int(eid))
                        if key in lookup:
                            stress = lookup[key]
                            if stress is not None:
                                total_stress += stress * multiplier
                                components.append(f"{case_id}*{multiplier}")

                    if components:
                        results.append({
                            'Combined_LC': comb_lc, 'Element': int(eid),
                            'Combined_Stress': total_stress,
                            'Components': ' + '.join(components)
                        })

            # Save combined_stress_results.csv
            if results:
                comb_csv = os.path.join(folder, 'combined_stress_results.csv')
                with open(comb_csv, 'w', newline='') as f:
                    w = csv.DictWriter(f, fieldnames=['Combined_LC', 'Element', 'Combined_Stress', 'Components'])
                    w.writeheader()
                    w.writerows(results)
                self.log(f"    Saved: combined_stress_results.csv ({len(results)} rows)")
                return results

        except Exception as e:
            self.log(f"    Combine error: {e}")

        return None

    def _combine_stresses_multi(self, folder, n_bdfs):
        """Combine stresses from multiple BDFs using Residual Strength table."""
        if self.residual_strength_df is None or len(self.combination_table) == 0:
            self.log("    No Residual Strength data - skipping combination")
            return None

        try:
            # Collect stress data from all BDF subfolders
            all_stress_data = []

            if n_bdfs > 1:
                # Look for subfolders
                for item in os.listdir(folder):
                    subfolder = os.path.join(folder, item)
                    if os.path.isdir(subfolder) and item.startswith('bdf_'):
                        stress_csv = os.path.join(subfolder, 'bar_stress_results.csv')
                        if os.path.exists(stress_csv):
                            df = pd.read_csv(stress_csv)
                            all_stress_data.append(df)
                            self.log(f"    Loaded stresses from: {item}")
            else:
                # Single BDF case
                stress_csv = os.path.join(folder, 'bar_stress_results.csv')
                if os.path.exists(stress_csv):
                    all_stress_data.append(pd.read_csv(stress_csv))

            if not all_stress_data:
                self.log("    No bar_stress_results.csv found in any BDF folder")
                return None

            # Merge all stress data
            stress_df = pd.concat(all_stress_data, ignore_index=True)
            self.log(f"    Combined stress data: {len(stress_df)} rows from {len(all_stress_data)} BDFs")

            # Build lookup: (subcase, element) -> stress (use max stress if duplicates)
            lookup = {}
            for _, row in stress_df.iterrows():
                key = (int(row['Subcase']), int(row['Element']))
                stress_val = row['Stress'] if pd.notna(row['Stress']) else 0
                if key not in lookup or abs(stress_val) > abs(lookup[key]):
                    lookup[key] = stress_val

            elements = stress_df['Element'].unique()

            rs_df = self.residual_strength_df
            cols = rs_df.columns.tolist()
            comb_col = cols[0]

            results = []
            for _, rs_row in rs_df.iterrows():
                comb_lc = rs_row[comb_col]
                if pd.isna(comb_lc):
                    continue
                comb_lc = int(comb_lc)

                for eid in elements:
                    total_stress = 0.0
                    components = []

                    for case_col, mult_col in self.combination_table:
                        case_id = rs_row[case_col]
                        multiplier = rs_row[mult_col]
                        if pd.isna(case_id) or pd.isna(multiplier):
                            continue
                        case_id = int(case_id)
                        multiplier = float(multiplier)

                        key = (case_id, int(eid))
                        if key in lookup:
                            stress = lookup[key]
                            if stress is not None:
                                total_stress += stress * multiplier
                                components.append(f"{case_id}*{multiplier}")

                    if components:
                        results.append({
                            'Combined_LC': comb_lc, 'Element': int(eid),
                            'Combined_Stress': total_stress,
                            'Components': ' + '.join(components)
                        })

            # Save combined results
            if results:
                comb_csv = os.path.join(folder, 'combined_stress_results.csv')
                with open(comb_csv, 'w', newline='') as f:
                    w = csv.DictWriter(f, fieldnames=['Combined_LC', 'Element', 'Combined_Stress', 'Components'])
                    w.writeheader()
                    w.writerows(results)
                self.log(f"    Saved: combined_stress_results.csv ({len(results)} rows)")
                return results

        except Exception as e:
            self.log(f"    Multi-BDF combine error: {e}")

        return None

    def _calculate_rf(self, stresses, target_rf, combined_stresses=None):
        """Calculate RF using combined stresses if available, otherwise raw stresses."""
        details = []
        failing_pids = set()
        elem_fit = prop_fit = 0

        # If combined stresses are available, build lookup
        combined_lookup = {}
        if combined_stresses:
            # Find max combined stress per element (critical case)
            for cs in combined_stresses:
                eid = cs['Element']
                comb_stress = abs(cs['Combined_Stress'])
                if eid not in combined_lookup or comb_stress > combined_lookup[eid]:
                    combined_lookup[eid] = comb_stress

        for s in stresses:
            eid = s['eid']
            etype = s['type']
            pid = self.elem_to_prop.get(eid)

            # Use combined stress if available, otherwise use raw stress
            if eid in combined_lookup and etype == 'bar':
                stress = combined_lookup[eid]
                stress_src = 'combined'
            else:
                stress = abs(s['stress'])
                stress_src = 'raw'

            if etype == 'bar' and pid:
                t = self.current_bar_thicknesses.get(pid, float(self.bar_min_thickness.get()))
            else:
                t = self.current_skin_thicknesses.get(pid, float(self.skin_min_thickness.get())) if pid else float(self.skin_min_thickness.get())

            # Try element fit first, then property fit
            allowable = None
            fit_src = "none"

            if eid in self.allowable_elem_interp:
                allowable = self.get_allowable_stress_elem(eid, t)
                if allowable:
                    fit_src = "element"
                    elem_fit += 1

            if allowable is None and pid:
                allowable = self.get_allowable_stress(pid, t)
                if allowable:
                    fit_src = "property"
                    prop_fit += 1

            if stress == 0:
                rf = 999.0
                status = 'PASS'
            elif allowable and allowable > 0:
                rf = allowable / stress
                status = 'PASS' if rf >= target_rf else 'FAIL'
            else:
                rf = 0
                status = 'NO_ALLOW'

            if status == 'FAIL' and pid:
                failing_pids.add(pid)

            req_t = self.get_required_thickness(pid, stress, target_rf) if pid and status == 'FAIL' else None

            details.append({
                'eid': eid, 'pid': pid, 'type': etype, 'thickness': t,
                'stress': stress, 'allowable': allowable, 'rf': rf,
                'status': status, 'fit_src': fit_src, 'req_thickness': req_t,
                'stress_src': stress_src
            })

        valid_rf = [d['rf'] for d in details if 0 < d['rf'] < 999]
        min_rf = min(valid_rf) if valid_rf else 0
        n_fail = sum(1 for d in details if d['status'] == 'FAIL')

        return {'min_rf': min_rf, 'n_fail': n_fail, 'details': details, 'failing_pids': failing_pids}

    def _calculate_weight(self):
        weight = 0.0

        for pid in self.skin_properties:
            t = self.current_skin_thicknesses.get(pid, 0)
            rho = self.get_density(pid)
            if pid in self.prop_elements:
                area = sum(self.element_areas.get(eid, 0) for eid in self.prop_elements[pid])
                weight += area * t * rho

        for pid in self.bar_properties:
            dim1 = self.current_bar_thicknesses.get(pid, 0)  # Optimized dimension
            # Use BDF PBARL dim2 (original geometry), fallback to Excel
            if pid in self.pbarl_dims:
                dim2 = self.pbarl_dims[pid]['dim2']
            else:
                dim2 = self.bar_properties[pid].get('dim2', dim1)
            rho = self.get_density(pid)
            if pid in self.prop_elements:
                length = sum(self.bar_lengths.get(eid, 0) for eid in self.prop_elements[pid])
                # Cross-sectional area = dim1 * dim2 for rectangular bar
                weight += length * dim1 * dim2 * rho

        return weight

    def _calculate_weight_sensitivities(self):
        """Calculate weight sensitivity (dW/dt) for each property.

        Returns dict: {pid: dW/dt} where dW/dt is weight change per unit thickness change.

        For bars: dW/dt = total_length × dim2 × density
        For skins: dW/dt = total_area × density

        This tells us how much weight we add per mm of thickness increase.
        """
        sensitivities = {}

        # Skin properties: dW/dt = total_area × density
        for pid in self.skin_properties:
            rho = self.get_density(pid)
            if pid in self.prop_elements:
                total_area = sum(self.element_areas.get(eid, 0) for eid in self.prop_elements[pid])
                sensitivities[pid] = total_area * rho
            else:
                sensitivities[pid] = 0

        # Bar properties: dW/dt = total_length × dim2 × density
        for pid in self.bar_properties:
            # dim2 is the non-optimized dimension (stays constant)
            if pid in self.pbarl_dims:
                dim2 = self.pbarl_dims[pid]['dim2']
            else:
                dim2 = self.bar_properties[pid].get('dim2', 1.0)

            rho = self.get_density(pid)
            if pid in self.prop_elements:
                total_length = sum(self.bar_lengths.get(eid, 0) for eid in self.prop_elements[pid])
                sensitivities[pid] = total_length * dim2 * rho
            else:
                sensitivities[pid] = 0

        return sensitivities

    def _smart_thickness_update(self, result, step, bar_min, bar_max, skin_min, skin_max, target_rf, rf_tol):
        """Per-property smart thickness update based on RF sensitivity."""
        details = result.get('rf_details', [])
        failing_pids = result.get('failing_pids', set())

        # Group by PID and find min RF per property
        pid_data = {}
        for d in details:
            pid = d['pid']
            if pid is None:
                continue
            if pid not in pid_data:
                pid_data[pid] = {'min_rf': 999, 'max_stress': 0, 'req_thickness': None}
            if d['rf'] < pid_data[pid]['min_rf']:
                pid_data[pid]['min_rf'] = d['rf']
                pid_data[pid]['max_stress'] = d['stress']
                pid_data[pid]['req_thickness'] = d.get('req_thickness')

        updated = 0

        # Increase failing properties
        for pid in failing_pids:
            if pid in self.current_bar_thicknesses:
                current = self.current_bar_thicknesses[pid]
                req_t = pid_data.get(pid, {}).get('req_thickness')
                if req_t and req_t > current:
                    new_t = min(req_t * 1.05, bar_max)  # 5% margin
                else:
                    new_t = min(current + step, bar_max)
                if new_t > current:
                    self.current_bar_thicknesses[pid] = new_t
                    updated += 1

            elif pid in self.current_skin_thicknesses:
                current = self.current_skin_thicknesses[pid]
                req_t = pid_data.get(pid, {}).get('req_thickness')
                if req_t and req_t > current:
                    new_t = min(req_t * 1.05, skin_max)
                else:
                    new_t = min(current + step, skin_max)
                if new_t > current:
                    self.current_skin_thicknesses[pid] = new_t
                    updated += 1

        # Decrease over-designed properties (RF > target + tolerance + 0.2)
        reduce_threshold = target_rf + rf_tol + 0.2
        for pid, data in pid_data.items():
            if data['min_rf'] > reduce_threshold and pid not in failing_pids:
                if pid in self.current_bar_thicknesses:
                    current = self.current_bar_thicknesses[pid]
                    new_t = max(current - step/2, bar_min)
                    if new_t < current:
                        self.current_bar_thicknesses[pid] = new_t
                        updated += 1

                elif pid in self.current_skin_thicknesses:
                    current = self.current_skin_thicknesses[pid]
                    new_t = max(current - step/2, skin_min)
                    if new_t < current:
                        self.current_skin_thicknesses[pid] = new_t
                        updated += 1

        if updated:
            self.log(f"  Updated {updated} properties")
        else:
            self.log(f"  No property updates needed (failing: {len(failing_pids)}, at max: {sum(1 for p in failing_pids if (p in self.current_bar_thicknesses and self.current_bar_thicknesses[p] >= bar_max) or (p in self.current_skin_thicknesses and self.current_skin_thicknesses[p] >= skin_max))})")

    def _save_iteration(self, folder, result):
        try:
            with open(os.path.join(folder, "summary.csv"), 'w', newline='') as f:
                w = csv.writer(f)
                w.writerow(['Parameter', 'Value'])
                w.writerow(['Iteration', result['iteration']])
                w.writerow(['Min_RF', result['min_rf']])
                w.writerow(['N_Fail', result['n_fail']])
                w.writerow(['Weight', result['weight']])

            pd.DataFrame(result['rf_details']).to_csv(os.path.join(folder, "rf_details.csv"), index=False)

            bar_data = [{'PID': p, 'Thickness': t} for p, t in result['bar_thicknesses'].items()]
            skin_data = [{'PID': p, 'Thickness': t} for p, t in result['skin_thicknesses'].items()]
            pd.DataFrame(bar_data).to_csv(os.path.join(folder, "bar_thicknesses.csv"), index=False)
            pd.DataFrame(skin_data).to_csv(os.path.join(folder, "skin_thicknesses.csv"), index=False)

        except Exception as e:
            self.log(f"  Save error: {e}")

    def _save_results(self, folder):
        try:
            history = [{'iteration': r['iteration'], 'min_rf': r['min_rf'], 'n_fail': r['n_fail'], 'weight': r['weight']} for r in self.iteration_results]
            pd.DataFrame(history).to_csv(os.path.join(folder, "history.csv"), index=False)

            if self.best_solution:
                pd.DataFrame([{'iteration': self.best_solution['iteration'], 'min_rf': self.best_solution['min_rf'], 'weight': self.best_solution['weight']}]).to_csv(os.path.join(folder, "best.csv"), index=False)

        except Exception as e:
            self.log(f"Save error: {e}")

    def _generate_final_report(self, folder, algorithm_name, nastran_count=None, extra_info=None):
        """Generate comprehensive final optimization report."""
        if not self.best_solution:
            self.log("No best solution found - cannot generate report")
            return

        report_lines = []
        sep = "═" * 70

        # Header
        report_lines.append("")
        report_lines.append(sep)
        report_lines.append("              OPTIMIZATION FINAL REPORT")
        report_lines.append(sep)
        report_lines.append(f"  Algorithm: {algorithm_name}")
        report_lines.append(f"  Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(sep)

        # Best Solution Summary
        report_lines.append("")
        report_lines.append("  BEST SOLUTION FOUND:")
        report_lines.append("  " + "-" * 40)
        report_lines.append(f"    Iteration/Evaluation: {self.best_solution['iteration']}")
        report_lines.append(f"    Minimum RF: {self.best_solution['min_rf']:.4f}")
        report_lines.append(f"    Total Weight: {self.best_solution['weight']:.6f} tonnes")
        report_lines.append(f"    Failed Elements: {self.best_solution['n_fail']}")

        target_rf = float(self.target_rf.get())
        if self.best_solution['min_rf'] >= target_rf:
            report_lines.append(f"    Status: ✓ CONVERGED (RF >= {target_rf})")
        else:
            report_lines.append(f"    Status: ✗ NOT CONVERGED (RF < {target_rf})")

        # Bar Thickness Changes
        bar_min = float(self.bar_min_thickness.get())
        report_lines.append("")
        report_lines.append("  BAR THICKNESS RESULTS:")
        report_lines.append("  " + "-" * 40)
        report_lines.append(f"    {'PID':<10} {'Initial':>10} {'Final':>10} {'Change':>12}")
        report_lines.append("    " + "-" * 44)

        bar_thicknesses = self.best_solution.get('bar_thicknesses', {})
        total_bar_increase = 0
        for pid in sorted(bar_thicknesses.keys()):
            final_t = bar_thicknesses[pid]
            initial_t = bar_min
            change_pct = ((final_t - initial_t) / initial_t * 100) if initial_t > 0 else 0
            total_bar_increase += (final_t - initial_t)
            report_lines.append(f"    {pid:<10} {initial_t:>10.2f} {final_t:>10.2f} {change_pct:>+10.1f}%")

        if bar_thicknesses:
            avg_bar = sum(bar_thicknesses.values()) / len(bar_thicknesses)
            report_lines.append("    " + "-" * 44)
            report_lines.append(f"    {'Average':<10} {bar_min:>10.2f} {avg_bar:>10.2f}")

        # Skin Thickness Changes
        skin_min = float(self.skin_min_thickness.get())
        skin_thicknesses = self.best_solution.get('skin_thicknesses', {})
        if skin_thicknesses:
            report_lines.append("")
            report_lines.append("  SKIN THICKNESS RESULTS:")
            report_lines.append("  " + "-" * 40)
            report_lines.append(f"    {'PID':<10} {'Initial':>10} {'Final':>10} {'Change':>12}")
            report_lines.append("    " + "-" * 44)

            for pid in sorted(skin_thicknesses.keys()):
                final_t = skin_thicknesses[pid]
                initial_t = skin_min
                change_pct = ((final_t - initial_t) / initial_t * 100) if initial_t > 0 else 0
                report_lines.append(f"    {pid:<10} {initial_t:>10.2f} {final_t:>10.2f} {change_pct:>+10.1f}%")

            avg_skin = sum(skin_thicknesses.values()) / len(skin_thicknesses)
            report_lines.append("    " + "-" * 44)
            report_lines.append(f"    {'Average':<10} {skin_min:>10.2f} {avg_skin:>10.2f}")

        # RF Statistics
        rf_details = self.best_solution.get('rf_details', [])
        if rf_details:
            valid_rfs = [d['rf'] for d in rf_details if d['rf'] is not None and 0 < d['rf'] < 999]
            if valid_rfs:
                report_lines.append("")
                report_lines.append("  RF STATISTICS:")
                report_lines.append("  " + "-" * 40)
                report_lines.append(f"    Minimum RF: {min(valid_rfs):.4f}")
                report_lines.append(f"    Maximum RF: {max(valid_rfs):.4f}")
                report_lines.append(f"    Average RF: {sum(valid_rfs)/len(valid_rfs):.4f}")
                report_lines.append(f"    Total Elements Analyzed: {len(rf_details)}")

                # RF Distribution
                below_target = sum(1 for rf in valid_rfs if rf < target_rf)
                near_target = sum(1 for rf in valid_rfs if target_rf <= rf < target_rf + 0.2)
                above_target = sum(1 for rf in valid_rfs if rf >= target_rf + 0.2)

                report_lines.append("")
                report_lines.append(f"    RF Distribution:")
                report_lines.append(f"      Below target (RF < {target_rf}): {below_target} elements")
                report_lines.append(f"      Near target ({target_rf} ≤ RF < {target_rf+0.2}): {near_target} elements")
                report_lines.append(f"      Above target (RF ≥ {target_rf+0.2}): {above_target} elements")

                # Critical elements (lowest RF)
                report_lines.append("")
                report_lines.append("    Critical Elements (Lowest RF):")
                sorted_details = sorted([d for d in rf_details if d['rf'] and 0 < d['rf'] < 999],
                                       key=lambda x: x['rf'])[:5]
                for d in sorted_details:
                    report_lines.append(f"      EID {d['eid']}: RF={d['rf']:.4f}, PID={d['pid']}, Stress={d['stress']:.2f}")

        # Convergence Info
        report_lines.append("")
        report_lines.append("  CONVERGENCE INFO:")
        report_lines.append("  " + "-" * 40)

        if nastran_count:
            report_lines.append(f"    Total Nastran Evaluations: {nastran_count}")

        if self.iteration_results:
            first_result = self.iteration_results[0]
            report_lines.append(f"    Initial Min RF: {first_result['min_rf']:.4f}")
            report_lines.append(f"    Final Min RF: {self.best_solution['min_rf']:.4f}")
            rf_improvement = self.best_solution['min_rf'] - first_result['min_rf']
            report_lines.append(f"    RF Improvement: {rf_improvement:+.4f}")

            report_lines.append(f"    Initial Weight: {first_result['weight']:.6f} t")
            report_lines.append(f"    Final Weight: {self.best_solution['weight']:.6f} t")
            weight_change = ((self.best_solution['weight'] - first_result['weight']) / first_result['weight'] * 100) if first_result['weight'] > 0 else 0
            report_lines.append(f"    Weight Change: {weight_change:+.1f}%")

        # Extra algorithm-specific info
        if extra_info:
            report_lines.append("")
            report_lines.append("  ALGORITHM SPECIFIC:")
            report_lines.append("  " + "-" * 40)
            for key, value in extra_info.items():
                report_lines.append(f"    {key}: {value}")

        # Footer
        report_lines.append("")
        report_lines.append(sep)
        report_lines.append(f"  Report saved to: {folder}")
        report_lines.append(sep)
        report_lines.append("")

        # Join and display
        report_text = "\n".join(report_lines)
        self.log(report_text)

        # Save to file
        try:
            report_path = os.path.join(folder, "FINAL_REPORT.txt")
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
            self.log(f"Report saved to: {report_path}")

            # Also save detailed thickness CSV
            thickness_data = []
            for pid, t in bar_thicknesses.items():
                thickness_data.append({
                    'Type': 'BAR',
                    'PID': pid,
                    'Initial_mm': bar_min,
                    'Final_mm': t,
                    'Change_pct': ((t - bar_min) / bar_min * 100) if bar_min > 0 else 0
                })
            for pid, t in skin_thicknesses.items():
                thickness_data.append({
                    'Type': 'SKIN',
                    'PID': pid,
                    'Initial_mm': skin_min,
                    'Final_mm': t,
                    'Change_pct': ((t - skin_min) / skin_min * 100) if skin_min > 0 else 0
                })
            if thickness_data:
                pd.DataFrame(thickness_data).to_csv(os.path.join(folder, "final_thicknesses.csv"), index=False)

        except Exception as e:
            self.log(f"Error saving report: {e}")

    def _update_ui(self):
        if self.best_solution:
            self.result_summary.config(text=f"Best: Weight={self.best_solution['weight']:.6f}t, RF={self.best_solution['min_rf']:.4f}", foreground="green")

            self.best_text.delete(1.0, tk.END)
            txt = f"Best Solution (Iteration {self.best_solution['iteration']}):\n"
            txt += f"  Weight: {self.best_solution['weight']:.6f} tonnes\n"
            txt += f"  Min RF: {self.best_solution['min_rf']:.4f}\n"
            txt += f"  Failures: {self.best_solution['n_fail']}\n\n"

            txt += "Bar Thicknesses (sample):\n"
            for pid in list(self.best_solution['bar_thicknesses'].keys())[:10]:
                txt += f"  PID {pid}: {self.best_solution['bar_thicknesses'][pid]:.2f} mm\n"

            txt += "\nSkin Thicknesses (sample):\n"
            for pid in list(self.best_solution['skin_thicknesses'].keys())[:10]:
                txt += f"  PID {pid}: {self.best_solution['skin_thicknesses'][pid]:.2f} mm\n"

            self.best_text.insert(tk.END, txt)

    def export_results(self):
        if not self.iteration_results:
            messagebox.showerror("Error", "No results")
            return

        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = os.path.join(self.output_folder.get(), f"results_{ts}.xlsx")

            with pd.ExcelWriter(path, engine='openpyxl') as w:
                history = [{'iteration': r['iteration'], 'min_rf': r['min_rf'], 'n_fail': r['n_fail'], 'weight': r['weight']} for r in self.iteration_results]
                pd.DataFrame(history).to_excel(w, sheet_name='History', index=False)

                if self.best_solution:
                    bar_data = [{'PID': p, 'Thickness': t} for p, t in self.best_solution['bar_thicknesses'].items()]
                    skin_data = [{'PID': p, 'Thickness': t} for p, t in self.best_solution['skin_thicknesses'].items()]
                    pd.DataFrame(bar_data).to_excel(w, sheet_name='Bar_Thicknesses', index=False)
                    pd.DataFrame(skin_data).to_excel(w, sheet_name='Skin_Thicknesses', index=False)

            self.log(f"\nExported: {path}")
            messagebox.showinfo("Export", f"Saved to:\n{path}")

        except Exception as e:
            self.log(f"Export error: {e}")



def main():
    root = tk.Tk()
    app = IntegratedBDFRFTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
