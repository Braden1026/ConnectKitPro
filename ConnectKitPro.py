# -----------------------------------------------------------------------------
# Application: Connect Kit Pro
# Author:      Braden Yates
# Copyright:   (c) 2026 Braden Yates. All rights reserved.
# License:     Proprietary. Not for public distribution or modification.
# Description: A network diagnostic tool for MSP technicians.
# -----------------------------------------------------------------------------

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import threading
import socket
import smtplib
import ssl
import uuid
import os
import re
from email.mime.text import MIMEText
from email.utils import formatdate
import sys
import concurrent.futures
import ipaddress

__author__ = "Braden Yates"
__copyright__ = "Copyright 2026, Braden Yates"
__credits__ = ["CustomTkinter", "PySNMP", "SMBProtocol"]
__license__ = "Proprietary"
__version__ = "1.1"
__maintainer__ = "Braden Yates"
__email__ = "byates@dme.us.com"
__status__ = "Production"

# --- CONFIGURATION ---
ctk.set_appearance_mode("Dark")  # Options: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Options: "blue", "green", "dark-blue"

# --- LIBRARIES ---
try:
    from smbprotocol.connection import Connection
    from smbprotocol.session import Session
    from smbprotocol.tree import TreeConnect
    import smbclient 
    from pysnmp.hlapi import *
    from pysnmp.hlapi import UsmUserData, usmHMACSHAAuthProtocol, usmAesCfb128Protocol
    import openpyxl
    from openpyxl.styles import Font
    LIBRARIES_OK = True
except ImportError as e:
    print(f"Library Error: {e}")
    LIBRARIES_OK = False

class ConnectKit:
    def check_for_updates(self, manual_check=True):
        import urllib.request
        import json
        import threading
        
        # TODO: Replace 'YourGitHubUsername/YourRepoName' with your actual details
        github_repo = "YourGitHubUsername/ConnectKitPro" 
        api_url = f"https://api.github.com/repos/{github_repo}/releases/latest"
        
        def run_check():
            try:
                # GitHub API requires a User-Agent header
                req = urllib.request.Request(api_url, headers={'User-Agent': 'ConnectKitPro-App'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    data = json.loads(response.read().decode('utf-8'))
                    
                # GitHub release tags usually look like "v1.0.1". Strip the 'v' for math.
                remote_version = data.get("tag_name", "v0.0.0").lstrip('v')
                
                # Link to the actual GitHub Release page so they can download the .exe
                download_url = data.get("html_url", "") 
                
                # Convert version strings to tuples for accurate math (e.g., 1.0.10 > 1.0.9)
                current_v_tuple = tuple(map(int, __version__.split('.')))
                remote_v_tuple = tuple(map(int, remote_version.split('.')))
                
                if remote_v_tuple > current_v_tuple:
                    # Update available! Update UI from the main thread
                    self.root.after(0, lambda: self.prompt_update(remote_version, download_url))
                elif manual_check:
                    self.root.after(0, lambda: messagebox.showinfo("Up to Date", f"You are running the latest version ({__version__})."))
                    
            except Exception as e:
                if manual_check:
                    self.root.after(0, lambda: messagebox.showerror("Update Check Failed", f"Could not check GitHub for updates.\nError: {e}"))

        # Run the network request in a separate thread so it doesn't freeze the GUI
        threading.Thread(target=run_check, daemon=True).start()

    def prompt_update(self, new_version, url):
        import webbrowser
        msg = f"A new version of Connect Kit Pro is available!\n\nVersion: {new_version}\n\nWould you like to open the release page to download it?"
        if messagebox.askyesno("Update Available", msg):
            webbrowser.open(url)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def __init__(self, root):
        self.root = root
        self.root.title("Connect Kit Pro")
        self.root.geometry("900x700")  # Slightly larger to accommodate SNMP UI
        
        # --- SET WINDOW ICON ---
        try:
            icon_path = self.resource_path("icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon Warning: {e}")

        self.scanning_ports = False
        
        # SNMP State Variables
        self.is_scanning_snmp = False
        self.found_devices = []
        
        if not LIBRARIES_OK:
            messagebox.showerror("Missing Library", "Please run: pip install smbprotocol pysnmp openpyxl")
            self.root.destroy()
            return

        # --- MAIN TABVIEW ---
        self.tabview = ctk.CTkTabview(self.root, width=800, height=700)

        # 1. SMTP
        self.tabview.add("SMTP / Email")
        # 2. SMB
        self.tabview.add("SMB / Scan to Folder")
        # 3. SNMP (Moved Up)
        self.tabview.add("SNMP Scanner")
        # 4. Port Probe (Moved Down)
        self.tabview.add("Port Probe")

        # --- BOTTOM UTILITY BAR ---
        self.bottom_bar = ctk.CTkFrame(self.root, fg_color="transparent")
        self.bottom_bar.pack(fill="x", side="bottom", padx=20, pady=(0, 10))

        self.btn_update = ctk.CTkButton(
            self.bottom_bar, 
            text=f"v{__version__} - Check for Updates", 
            width=150, 
            fg_color="transparent", 
            border_width=1,
            text_color="gray",
            hover_color="#333333",
            command=self.check_for_updates
        )
        self.btn_update.pack(side="right")

        self.tabview.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.tab_smtp = self.tabview.tab("SMTP / Email")
        self.tab_smb = self.tabview.tab("SMB / Scan to Folder")
        self.tab_snmp = self.tabview.tab("SNMP Scanner")
        self.tab_ports = self.tabview.tab("Port Probe")
        
        self.setup_smtp_tab()
        self.setup_smb_tab()
        self.setup_snmp_tab()  # Setup SNMP 3rd
        self.setup_ports_tab() # Setup Ports 4th

        # Global Enter Key
        self.root.bind('<Return>', self.on_enter_press)

    # --- HELPERS ---
    def is_valid_ip(self, ip_str):
        hostname_regex = r"^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])$"
        ipv4_regex = r"^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
        if re.match(ipv4_regex, ip_str): return True
        elif re.match(hostname_regex, ip_str): return True 
        return False

    def is_valid_email(self, email_str):
        email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
        return re.match(email_regex, email_str) is not None

    def on_enter_press(self, event):
        current = self.tabview.get()
        if current == "SMTP / Email": self.start_smtp_test()
        elif current == "SMB / Scan to Folder": self.start_smb_check()
        elif current == "SNMP Scanner": self.toggle_snmp_scan()
        elif current == "Port Probe": self.toggle_port_scan()

    # ==========================================
    # TAB 1: SMTP DOCTOR
    # ==========================================
    def setup_smtp_tab(self):
        frame = self.tab_smtp
        
        # 1. Inputs Area
        input_frame = ctk.CTkFrame(frame)
        input_frame.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(input_frame, text="Server Configuration", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, columnspan=4, sticky="w", padx=15, pady=10)

        # Row 1
        ctk.CTkLabel(input_frame, text="Server Address:").grid(row=1, column=0, padx=15, pady=5, sticky="e")
        self.ent_server = ctk.CTkEntry(input_frame, width=250, placeholder_text="smtp.office365.com")
        self.ent_server.grid(row=1, column=1, padx=5, pady=5)

        ctk.CTkLabel(input_frame, text="Port:").grid(row=1, column=2, padx=15, pady=5, sticky="e")
        self.ent_port = ctk.CTkEntry(input_frame, width=80)
        self.ent_port.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.ent_port.insert(0, "587"); self.ent_port.configure(state="disabled")

        # Row 2
        ctk.CTkLabel(input_frame, text="Encryption:").grid(row=2, column=0, padx=15, pady=5, sticky="e")
        self.cmb_enc = ctk.CTkComboBox(input_frame, width=250, values=["STARTTLS (587)", "SSL/TLS (465)", "None (25)", "Custom"], state="readonly", command=self.update_smtp_port)
        self.cmb_enc.grid(row=2, column=1, padx=5, pady=5)
        self.cmb_enc.set("STARTTLS (587)")

        # Row 3
        ctk.CTkLabel(input_frame, text="Email/User:").grid(row=3, column=0, padx=15, pady=5, sticky="e")
        self.ent_user = ctk.CTkEntry(input_frame, width=250, placeholder_text="copier@company.com")
        self.ent_user.grid(row=3, column=1, padx=5, pady=5)

        ctk.CTkLabel(input_frame, text="Password:").grid(row=3, column=2, padx=15, pady=5, sticky="e")
        self.ent_pass = ctk.CTkEntry(input_frame, width=180, show="*")
        self.ent_pass.grid(row=3, column=3, padx=5, pady=5, sticky="w")

        # Row 4
        ctk.CTkLabel(input_frame, text="Send Test To:").grid(row=4, column=0, padx=15, pady=5, sticky="e")
        self.ent_to = ctk.CTkEntry(input_frame, width=250, placeholder_text="admin@company.com")
        self.ent_to.grid(row=4, column=1, padx=5, pady=5)

        # 2. Action Buttons
        self.btn_test = ctk.CTkButton(frame, text="RUN DIAGNOSTIC", height=40, font=ctk.CTkFont(size=13, weight="bold"), command=self.start_smtp_test)
        self.btn_test.pack(fill="x", padx=10, pady=10)

        # 3. Console Log
        self.log_smtp = self.create_console(frame)

    def update_smtp_port(self, choice):
        self.ent_port.configure(state="normal")
        self.ent_port.delete(0, tk.END)
        if "587" in choice: self.ent_port.insert(0, "587"); self.ent_port.configure(state="disabled")
        elif "465" in choice: self.ent_port.insert(0, "465"); self.ent_port.configure(state="disabled")
        elif "25" in choice: self.ent_port.insert(0, "25"); self.ent_port.configure(state="disabled")
        elif "Custom" == choice: pass 

    def start_smtp_test(self):
        if not self.validate_smtp(): return
        self.log_smtp.delete(1.0, tk.END)
        self.btn_test.configure(state="disabled", text="Connecting...")
        threading.Thread(target=self.run_smtp_test, daemon=True).start()

    def validate_smtp(self):
        send_to = self.ent_to.get().strip()
        user = self.ent_user.get().strip()
        pwd = self.ent_pass.get().strip()
        if not self.is_valid_email(send_to): messagebox.showerror("Error", "Invalid Destination Email"); return False
        if user and not self.is_valid_email(user): messagebox.showerror("Error", "Invalid Username Email"); return False
        if "None" not in self.cmb_enc.get() and (not user or not pwd): messagebox.showerror("Error", "Credentials Required"); return False
        return True

    def run_smtp_test(self):
        server = self.ent_server.get().strip()
        try: port = int(self.ent_port.get().strip())
        except: self.log_s("Port Error", "ERROR"); self.reset_smtp_btn(); return
        
        user = self.ent_user.get().strip()
        pwd = self.ent_pass.get().strip()
        to = self.ent_to.get().strip()
        enc = self.cmb_enc.get()

        try:
            self.log_s(f"Connecting to {server}:{port}...", "SENT")
            if "SSL/TLS" in enc:
                ctx = ssl.create_default_context()
                smtp = smtplib.SMTP_SSL(server, port, context=ctx, timeout=10)
            else:
                smtp = smtplib.SMTP(server, port, timeout=10)
            
            self.log_s(f"Connected. Hello: {smtp.ehlo()[1].decode()}", "RECV")

            if "STARTTLS" in enc:
                self.log_s("Starting TLS...", "SENT")
                smtp.starttls()
                smtp.ehlo()
                self.log_s("TLS Established.", "RECV")

            if user and pwd:
                self.log_s(f"Authenticating as {user}...", "SENT")
                smtp.login(user, pwd)
                self.log_s("Auth Success.", "RECV")

            self.log_s(f"Sending email to {to}...", "SENT")
            msg = MIMEText("Test from Connect Kit Pro.")
            msg['Subject'] = "Test Email"
            msg['From'] = user
            msg['To'] = to
            msg['Date'] = formatdate(localtime=True)
            smtp.sendmail(user, [to], msg.as_string())
            self.log_s(">> EMAIL ACCEPTED FOR DELIVERY <<", "RECV")
            smtp.quit()

        except Exception as e:
            self.log_s(f"ERROR: {e}", "ERROR")
        finally:
            self.root.after(0, self.reset_smtp_btn)

    def reset_smtp_btn(self):
        self.btn_test.configure(state="normal", text="RUN DIAGNOSTIC")
    
    def log_s(self, text, tag="RECV"):
        self.log_smtp.insert(tk.END, text + "\n", tag)
        self.log_smtp.see(tk.END)

    # ==========================================
    # TAB 2: SMB DOCTOR
    # ==========================================
    def setup_smb_tab(self):
        frame = self.tab_smb
        
        # 1. Inputs
        input_frame = ctk.CTkFrame(frame)
        input_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(input_frame, text="SMB Share Credentials", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=10)

        def add_row(label, ph, row):
            ctk.CTkLabel(input_frame, text=label).grid(row=row, column=0, padx=15, pady=5, sticky="e")
            entry = ctk.CTkEntry(input_frame, width=350, placeholder_text=ph)
            entry.grid(row=row, column=1, padx=5, pady=5, sticky="w")
            return entry

        self.ent_smb_path = add_row("Network Path:", "Ex: \\\\IP Address\\Folder\\Subfolder", 1)
        self.ent_smb_user = add_row("Username:", "PC Username", 2)
        self.ent_smb_pass = add_row("Password:", "PC Password", 3)
        self.ent_smb_pass.configure(show="*")
        self.ent_smb_domain = add_row("Domain:", "Optional", 4)

        # 2. Buttons
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        self.btn_smb = ctk.CTkButton(btn_frame, text="RUN DIAGNOSTIC", height=40, font=ctk.CTkFont(weight="bold"), fg_color="#28a745", hover_color="#1e7e34", command=self.start_smb_check)
        self.btn_smb.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_cleanup = ctk.CTkButton(btn_frame, text="CLEAN UP FILE", height=40, font=ctk.CTkFont(weight="bold"), fg_color="#555", hover_color="#333", state="disabled", command=self.cleanup_smb_file)
        self.btn_cleanup.pack(side="right", fill="x", expand=True, padx=(5, 0))

        # 3. Console
        self.log_smb = self.create_console(frame)
        self.log_smb.configure(state="disabled")

    def log_f(self, text, tag="INFO"):
        self.log_smb.configure(state='normal')
        self.log_smb.insert(tk.END, text + "\n", tag)
        self.log_smb.see(tk.END)
        self.log_smb.configure(state='disabled')

    def start_smb_check(self):
        self.btn_smb.configure(state="disabled", text="Running...")
        self.log_smb.configure(state='normal'); self.log_smb.delete(1.0, tk.END); self.log_smb.configure(state='disabled')
        threading.Thread(target=self.run_smb_check, daemon=True).start()

    def run_smb_check(self):
        try:
            full_input = self.ent_smb_path.get().strip().replace('/', '\\')
            user = self.ent_smb_user.get().strip()
            pwd = self.ent_smb_pass.get().strip()
            domain = self.ent_smb_domain.get().strip()

            if not full_input or not user or not pwd: self.log_f("Missing Fields", "ERROR"); return

            # Remove leading slashes and parse the path
            clean_path = full_input.lstrip('\\')
            parts = clean_path.split('\\', 1)
            
            if len(parts) < 2: 
                self.log_f("Invalid Path Format (Expected \\\\IP\\Share)", "ERROR"); return
                
            ip = parts[0]
            share_path = parts[1] 
            root_share = share_path.split('\\')[0] 

            if not self.is_valid_ip(ip): self.log_f(f"Invalid IP/Hostname: {ip}", "ERROR"); return

            smbclient.reset_connection_cache()
            
            self.log_f(f"1. Ping {ip}...", "INFO")
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM); s.settimeout(2)
            try:
                if s.connect_ex((socket.gethostbyname(ip), 445)) != 0: self.log_f("Port 445 Closed", "ERROR"); return
            except: self.log_f("Host not found", "ERROR"); return
            s.close(); self.log_f("Host Online.", "SUCCESS")

            full_user = f"{domain}\\{user}" if domain else user
            self.log_f(f"2. Auth as {full_user}...", "INFO")
            conn = Connection(uuid.uuid4(), ip, 445); conn.connect()
            session = Session(conn, username=full_user, password=pwd); session.connect()
            self.log_f("Authenticated.", "SUCCESS")

            self.log_f(f"3. Connecting to \\{root_share}...", "INFO")
            tree = TreeConnect(session, f"\\\\{ip}\\{root_share}"); tree.connect()
            self.log_f("Share Accessible.", "SUCCESS")

            self.log_f("4. Write Test...", "INFO")
            try:
                smbclient.register_session(ip, username=full_user, password=pwd)
                path = f"\\\\{ip}\\{share_path}\\Scan To Folder Test.txt"
                with smbclient.open_file(path, mode="w") as f: f.write("The test was successful! You can now delete this file.")
                self.log_f("File Written Successfully.", "SUCCESS")
                self.log_f("NOTE: File left for verification.", "WARN")
                self.root.after(0, lambda: self.btn_cleanup.configure(state="normal", fg_color="#d9534f"))
            except Exception as e:
                self.log_f(f"Write Failed: {e}", "ERROR")

        except Exception as e:
            self.log_f(f"Error: {e}", "ERROR")
        finally:
            try: conn.disconnect(True)
            except: pass
            self.root.after(0, lambda: self.btn_smb.configure(state="normal", text="RUN DIAGNOSTIC"))

    def cleanup_smb_file(self):
        full_input = self.ent_smb_path.get().strip().replace('/', '\\')
        user = self.ent_smb_user.get().strip()
        pwd = self.ent_smb_pass.get().strip()
        domain = self.ent_smb_domain.get().strip()
        full_user = f"{domain}\\{user}" if domain else user
        
        clean_path = full_input.lstrip('\\')
        parts = clean_path.split('\\', 1)
        if len(parts) < 2: return
        
        ip = parts[0]
        share_path = parts[1]
        path = f"\\\\{ip}\\{share_path}\\Scan To Folder Test.txt"
        
        try:
            smbclient.register_session(ip, username=full_user, password=pwd)
            smbclient.remove(path)
            messagebox.showinfo("Done", "File Deleted.")
            self.btn_cleanup.configure(state="disabled", fg_color="#555")
            self.log_f("File Deleted.", "SUCCESS")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ==========================================
    # TAB 3: SNMP SCANNER
    # ==========================================
    def setup_snmp_tab(self):
        frame = self.tab_snmp
        
        # UI Container
        self.frm_snmp_settings = ctk.CTkFrame(frame)
        self.frm_snmp_settings.pack(pady=10, padx=20, fill="x")

        # Row 0
        ctk.CTkLabel(self.frm_snmp_settings, text="SNMP Version:").grid(row=0, column=0, padx=15, pady=15, sticky="e")
        self.cmb_snmp_ver = ctk.CTkComboBox(self.frm_snmp_settings, values=["v1/v2c (Legacy)", "v3 (Secure)"], command=self.toggle_snmp_auth_inputs)
        self.cmb_snmp_ver.grid(row=0, column=1, padx=10, pady=15, sticky="w")
        self.cmb_snmp_ver.set("v1/v2c (Legacy)")

        ctk.CTkLabel(self.frm_snmp_settings, text="Community:").grid(row=0, column=2, padx=15, pady=15, sticky="e")
        self.ent_comm = ctk.CTkEntry(self.frm_snmp_settings, width=150)
        self.ent_comm.grid(row=0, column=3, padx=10, pady=15, sticky="w")
        self.ent_comm.insert(0, "public")

        # Row 1
        self.lbl_snmp_user = ctk.CTkLabel(self.frm_snmp_settings, text="v3 Username:", text_color="gray")
        self.lbl_snmp_user.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="e")
        self.ent_snmp_user = ctk.CTkEntry(self.frm_snmp_settings, width=150, state="disabled")
        self.ent_snmp_user.grid(row=1, column=1, padx=10, pady=(0, 15), sticky="w")

        self.lbl_snmp_pass = ctk.CTkLabel(self.frm_snmp_settings, text="v3 Password:", text_color="gray")
        self.lbl_snmp_pass.grid(row=1, column=2, padx=15, pady=(0, 15), sticky="e")
        self.ent_snmp_pass = ctk.CTkEntry(self.frm_snmp_settings, width=150, show="*", state="disabled")
        self.ent_snmp_pass.grid(row=1, column=3, padx=10, pady=(0, 15), sticky="w")

        # Row 2 – Target
        ctk.CTkLabel(self.frm_snmp_settings, text="Target Network:").grid(row=2, column=0, padx=15, pady=(0, 15), sticky="e")
        self.ent_target = ctk.CTkEntry(self.frm_snmp_settings, width=250, placeholder_text="Auto detects by default")
        self.ent_target.grid(row=2, column=1, columnspan=3, padx=10, pady=(0, 15), sticky="w")

        # Buttons
        frm_actions = ctk.CTkFrame(frame, fg_color="transparent")
        frm_actions.pack(pady=10, fill="x", padx=20)

        self.btn_snmp_scan = ctk.CTkButton(frm_actions, text="START SCAN", height=40, font=ctk.CTkFont(weight="bold"), command=self.toggle_snmp_scan)
        self.btn_snmp_scan.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_snmp_export = ctk.CTkButton(frm_actions, text="EXPORT TO EXCEL", height=40, font=ctk.CTkFont(weight="bold"), state="disabled", command=self.save_excel_snmp)
        self.btn_snmp_export.pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.snmp_progress = ctk.CTkProgressBar(frame, height=10)
        self.snmp_progress.pack(fill="x", padx=20, pady=(5, 15))
        self.snmp_progress.set(0)

        self.txt_snmp_console = self.create_console(frame)
        self.txt_snmp_console.tag_config("HEADER", foreground="#ffd740", font=("Consolas", 10, "bold"))
        self.txt_snmp_console.configure(state="disabled")

    def toggle_snmp_auth_inputs(self, choice):
        is_v3 = "v3" in choice
        self.ent_snmp_user.configure(state="normal" if is_v3 else "disabled")
        self.ent_snmp_pass.configure(state="normal" if is_v3 else "disabled")
        self.ent_comm.configure(state="disabled" if is_v3 else "normal")
        self.lbl_snmp_user.configure(text_color="#e0e0e0" if is_v3 else "gray")
        self.lbl_snmp_pass.configure(text_color="#e0e0e0" if is_v3 else "gray")

    def log_snmp_msg(self, msg, tag="INFO"):
        self.txt_snmp_console.configure(state="normal")
        self.txt_snmp_console.insert(tk.END, msg + "\n", tag)
        self.txt_snmp_console.see(tk.END)
        self.txt_snmp_console.configure(state="disabled")

    def toggle_snmp_scan(self):
        if self.is_scanning_snmp:
            self.is_scanning_snmp = False
            return
        self.start_snmp_scan()

    def get_local_ip(self):
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(("8.8.8.8", 1))
            return s.getsockname()[0]
        except:
            return "127.0.0.1"
        finally:
            s.close()

    def parse_target_input(self, text):
        text = text.strip()
        if not text: return None
        if "/" in text:
            net = ipaddress.ip_network(text, strict=False)
            return [str(ip) for ip in net.hosts()]
        if "-" in text:
            start, end = text.split("-", 1)
            start = start.strip(); end = end.strip()
            start_ip = ipaddress.ip_address(start)
            if "." not in end:
                base = start.rsplit(".", 1)[0]
                end = f"{base}.{end}"
            end_ip = ipaddress.ip_address(end)
            if start_ip.version != end_ip.version: raise ValueError("IP versions do not match")
            if int(end_ip) < int(start_ip): raise ValueError("Invalid IP range")
            return [str(ipaddress.ip_address(int(start_ip) + i)) for i in range(int(end_ip) - int(start_ip) + 1)]
        ipaddress.ip_address(text)
        return [text]

    def is_snmp_host_alive(self, ip):
        ports = [9100, 80, 443, 515, 631]
        for port in ports:
            try:
                s = socket.create_connection((ip, port), timeout=0.4)
                s.close(); return True
            except: continue
        return False

    def start_snmp_scan(self):
        try:
            self.parse_target_input(self.ent_target.get())
        except ValueError as e:
            messagebox.showerror("Invalid Target", str(e))
            return

        self.is_scanning_snmp = True
        self.found_devices.clear()
        self.txt_snmp_console.configure(state="normal")
        self.txt_snmp_console.delete(1.0, tk.END)
        self.txt_snmp_console.configure(state="disabled")

        self.btn_snmp_scan.configure(text="STOP SCAN")
        self.btn_snmp_export.configure(state="disabled")

        conf = {
            "type": "v3" if "v3" in self.cmb_snmp_ver.get() else "v1",
            "comm": self.ent_comm.get(),
            "user": self.ent_snmp_user.get(),
            "auth": self.ent_snmp_pass.get(),
        }

        threading.Thread(target=self.run_snmp_scan_thread, args=(conf,), daemon=True).start()

    def run_snmp_scan_thread(self, conf):
        try:
            targets = self.parse_target_input(self.ent_target.get())
            if targets is None:
                base = ".".join(self.get_local_ip().split(".")[:3])
                targets = [f"{base}.{i}" for i in range(1, 255)]
                self.log_snmp_msg(f"--- AUTO SCANNING {base}.x ---", "HEADER")
            else:
                self.log_snmp_msg(f"--- SCANNING {len(targets)} HOSTS ---", "HEADER")

            auth = self.build_snmp_auth(conf)
            total = len(targets)

            with concurrent.futures.ThreadPoolExecutor(max_workers=40) as ex:
                futures = {ex.submit(self.scan_snmp_host, ip, auth): ip for ip in targets}
                done = 0
                for f in concurrent.futures.as_completed(futures):
                    if not self.is_scanning_snmp: break
                    done += 1
                    self.snmp_progress.set(done / total)
                    result = f.result()
                    if result:
                        self.found_devices.append(result)
                        self.log_snmp_device(result)

            self.log_snmp_msg(f"--- SCAN COMPLETE ({len(self.found_devices)} DEVICES) ---", "HEADER")
        finally:
            self.is_scanning_snmp = False
            self.root.after(0, self.reset_snmp_ui)

    def reset_snmp_ui(self):
        self.btn_snmp_scan.configure(text="START SCAN")
        self.snmp_progress.set(0)
        if self.found_devices:
            self.btn_snmp_export.configure(state="normal")

    def build_snmp_auth(self, conf):
        if conf["type"] == "v3":
            return UsmUserData(
                conf["user"],
                conf["auth"],
                conf["auth"],
                authProtocol=usmHMACSHAAuthProtocol,
                privProtocol=usmAesCfb128Protocol
            )
        return CommunityData(conf["comm"], mpModel=1)

    def snmp_get_val(self, ip, auth, oid):
        try:
            it = getCmd(SnmpEngine(), auth, UdpTransportTarget((ip, 161), timeout=0.8, retries=0), ContextData(), ObjectType(ObjectIdentity(oid)))
            err, stat, _, binds = next(it)
            if not err and not stat: return binds[0][1].prettyPrint()
        except: pass
        return None

    def scan_snmp_host(self, ip, auth):
        if not self.is_snmp_host_alive(ip): return None
        sys_desc = self.snmp_get_val(ip, auth, "1.3.6.1.2.1.1.1.0")
        if not sys_desc: return None

        model = self.snmp_get_val(ip, auth, "1.3.6.1.2.1.25.3.2.1.3.1") or "Unknown Device"
        serial = self.snmp_get_val(ip, auth, "1.3.6.1.2.1.43.5.1.1.17.1") or "Unknown"
        meter = self.snmp_get_val(ip, auth, "1.3.6.1.2.1.43.10.2.1.4.1.1") or "0"

        data = {"IP": ip, "Model": model, "Serial": serial, "Total Meter": meter, "Black": "", "Cyan": "", "Magenta": "", "Yellow": ""}
        oid_desc = "1.3.6.1.2.1.43.11.1.1.6.1"
        oid_max = "1.3.6.1.2.1.43.11.1.1.8.1"
        oid_lvl = "1.3.6.1.2.1.43.11.1.1.9.1"

        for i in range(1, 9):
            desc = self.snmp_get_val(ip, auth, f"{oid_desc}.{i}")
            if not desc: continue
            max_v = self.snmp_get_val(ip, auth, f"{oid_max}.{i}")
            cur_v = self.snmp_get_val(ip, auth, f"{oid_lvl}.{i}")
            pct = "?"
            try:
                if max_v and cur_v and int(max_v) > 0: pct = f"{int((int(cur_v) / int(max_v)) * 100)}%"
            except: pass
            d = desc.lower()
            if "black" in d or "toner k" in d: data["Black"] = pct
            elif "cyan" in d or "toner c" in d: data["Cyan"] = pct
            elif "magenta" in d or "toner m" in d: data["Magenta"] = pct
            elif "yellow" in d or "toner y" in d: data["Yellow"] = pct
        return data

    def log_snmp_device(self, d):
        self.log_snmp_msg(f"[{d['IP']}] {d['Model']}", "SUCCESS")
        self.log_snmp_msg(f"   Serial: {d['Serial']} | Meter: {d['Total Meter']}", "INFO")
        levels = []
        if d["Black"]: levels.append(f"Black: {d['Black']}")
        if d["Cyan"]: levels.append(f"Cyan: {d['Cyan']}")
        if d["Magenta"]: levels.append(f"Magenta: {d['Magenta']}")
        if d["Yellow"]: levels.append(f"Yellow: {d['Yellow']}")
        if levels: self.log_snmp_msg(f"   Levels: {' | '.join(levels)}", "INFO")
        self.log_snmp_msg("-" * 40, "INFO")

    def save_excel_snmp(self):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not f: return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Audit Data"
        headers = ["IP", "Model", "Serial", "Total Meter", "Black", "Cyan", "Magenta", "Yellow"]
        ws.append(headers)
        for c in ws[1]: c.font = Font(bold=True)
        for d in self.found_devices: ws.append([d.get(h, "") for h in headers])
        wb.save(f)
        messagebox.showinfo("Export Complete", f"Saved to:\n{f}")

    # ==========================================
    # TAB 4: PORT PROBE
    # ==========================================
    def setup_ports_tab(self):
        frame = self.tab_ports
        
        # 1. Input Bar
        bar = ctk.CTkFrame(frame)
        bar.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(bar, text="Target IP:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=15)
        self.ent_ip = ctk.CTkEntry(bar, width=200, placeholder_text="192.168.1.1")
        self.ent_ip.pack(side="left", padx=5)
        
        self.btn_port_scan = ctk.CTkButton(bar, text="START SCAN", width=120, font=ctk.CTkFont(weight="bold"), fg_color="#28a745", hover_color="#1e7e34", command=self.toggle_port_scan)
        self.btn_port_scan.pack(side="right", padx=15, pady=10)

        # 2. Results Grid (Scrollable)
        self.grid_frame = ctk.CTkScrollableFrame(frame, label_text="Scan Results")
        self.grid_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Grid Headers
        for i, txt in enumerate(["Service Name", "Port", "Status"]):
            ctk.CTkLabel(self.grid_frame, text=txt, font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=i, sticky="w", padx=20, pady=5)

        self.scan_targets = [
            ("Web Interface", [80, 443, 8080]),
            ("Printing (LPR/RAW)", [9100, 515]),
            ("Scanning (SMB)", [445, 139]),
            ("Email (SMTP)", [25, 465, 587]),
            ("SNMP", [161]),
            ("FTP", [20, 21]),
            ("Fiery/PaperCut", [631, 9191, 9192])
        ]
        
        self.port_labels = {}
        r = 1
        for name, ports in self.scan_targets:
            ctk.CTkLabel(self.grid_frame, text=name, text_color="#aaa").grid(row=r, column=0, sticky="nw", padx=20, pady=2)
            for p in ports:
                ctk.CTkLabel(self.grid_frame, text=str(p)).grid(row=r, column=1, sticky="w", padx=20)
                lbl = ctk.CTkLabel(self.grid_frame, text="---", text_color="gray")
                lbl.grid(row=r, column=2, sticky="w", padx=20)
                self.port_labels[p] = lbl
                r += 1
            ctk.CTkFrame(self.grid_frame, height=1, fg_color="#444").grid(row=r, column=0, columnspan=3, sticky="ew", pady=5)
            r += 1

    def toggle_port_scan(self):
        if self.scanning_ports:
            self.scanning_ports = False
            self.btn_port_scan.configure(text="Stopping...", state="disabled")
        else:
            ip = self.ent_ip.get().strip()
            if not ip or not self.is_valid_ip(ip): messagebox.showerror("Error", "Invalid IP"); return
            
            self.scanning_ports = True
            self.btn_port_scan.configure(text="STOP SCAN", fg_color="#d9534f", hover_color="#c9302c")
            for p in self.port_labels: self.port_labels[p].configure(text="Waiting...", text_color="gray")
            threading.Thread(target=self.run_port_scan, args=(ip,), daemon=True).start()

    # --- UDP SNMP CHECKER ---
    def check_snmp(self, ip):
        """ Sends a v2c 'public' request. If we get ANY reply (even an auth error), the port is OPEN. """
        try:
            iterator = getCmd(
                SnmpEngine(),
                CommunityData('public', mpModel=1), # v2c
                UdpTransportTarget((ip, 161), timeout=1.0, retries=0),
                ContextData(),
                ObjectType(ObjectIdentity("1.3.6.1.2.1.1.1.0")) # sysDescr
            )
            errorIndication, errorStatus, errorIndex, varBinds = next(iterator)
            
            # If errorIndication is 'Request Timed Out', the port is Closed/Filtered.
            # If it is None, we got data. 
            if errorIndication:
                return False 
            return True
        except Exception:
            return False

    def run_port_scan(self, ip):
        # 1. FIX THE NAME ERROR: Define the list of ports first
        all_ports = []
        for service, ports in self.scan_targets:
            all_ports.extend(ports)

        # 2. START THE LOOP
        for port in all_ports:
            if not self.scanning_ports: break
            
            # Update UI to show we are working on this port
            self.root.after(0, lambda p=port: self.port_labels[p].configure(text="Scanning...", text_color="#007ACC"))

            # 3. SPECIAL HANDLING FOR SNMP (UDP 161)
            if port == 161:
                # Use the helper function to check UDP
                is_open = self.check_snmp(ip)
                
                if is_open:
                    txt = "OPEN"
                    clr = "#28a745" # Green
                else:
                    txt = "CLOSED"
                    clr = "#d9534f" # Red 
                
            else:
                # 4. STANDARD TCP SCAN (Web, Print, SMB, Email)
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(0.5)
                try:
                    res = s.connect_ex((ip, port))
                    s.close()
                    txt = "OPEN" if res == 0 else "CLOSED"
                    clr = "#28a745" if res == 0 else "#d9534f" # Green / Red
                except:
                    txt = "ERROR"
                    clr = "orange"
            
            # Update UI with the result
            self.root.after(0, lambda p=port, t=txt, c=clr: self.port_labels[p].configure(text=t, text_color=c))
        
        # Reset Button when done
        self.scanning_ports = False
        self.root.after(0, lambda: self.btn_port_scan.configure(text="START SCAN", fg_color="#28a745", hover_color="#1e7e34", state="normal"))

    # --- UI HELPERS ---
    def create_console(self, parent):
        container = ctk.CTkFrame(parent)
        container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        text = tk.Text(container, height=10, bg="#212121", fg="#e0e0e0", font=('Consolas', 10), relief="flat", highlightthickness=0)
        text.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        sb = ctk.CTkScrollbar(container, command=text.yview)
        sb.pack(side="right", fill="y", padx=(0, 5), pady=5)
        text.configure(yscrollcommand=sb.set)
        
        text.tag_config("SENT", foreground="#64b5f6")
        text.tag_config("RECV", foreground="#69f0ae")
        text.tag_config("ERROR", foreground="#ff5252")
        text.tag_config("SUCCESS", foreground="#69f0ae")
        text.tag_config("WARN", foreground="#ffd740")
        text.tag_config("INFO", foreground="#b0bec5")
        
        return text

if __name__ == "__main__":
    app = ctk.CTk()
    gui = ConnectKit(app)
    app.mainloop()