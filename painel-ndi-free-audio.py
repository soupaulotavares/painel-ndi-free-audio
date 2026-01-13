# -*- coding: utf-8 -*-
import os
import sys
import subprocess
import ctypes
import customtkinter as ctk
from tkinter import messagebox

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr is not None:
    sys.stderr.reconfigure(encoding='utf-8')

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class NDIPanelApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        if getattr(sys, 'frozen', False):
            self.root_path = os.path.dirname(sys.executable)
        else:
            self.root_path = os.path.dirname(os.path.abspath(__file__))

        self.nssm = os.path.join(self.root_path, "nssm.exe")
        self.app_exe = os.path.join(self.root_path, "NDIFreeAudio.exe")
        self.vbcable_exe = os.path.join(self.root_path, "VBCABLE", "VBCABLE_Setup_x64.exe")
        
        self.title("Painel NDI Free Audio | 2.2.1")
        self.geometry("900x950") 
        
        if not self.is_admin():
            messagebox.showerror("Erro", "Execute como ADMINISTRADOR!")
            self.destroy()
            return

        self.setup_ui()

    def is_admin(self):
        try: return ctypes.windll.shell32.IsUserAnAdmin()
        except: return False

    def get_ndi_services(self):
        try:
            # Comando PowerShell para pegar Nome e Status formatados
            cmd = 'powershell -NoProfile -Command "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Get-Service | Where-Object { $_.Name -like \'NDI Audio*\' } | ForEach-Object { \'[\' + $_.Status.ToString() + \'] \' + $_.Name }"'
            output = subprocess.check_output(cmd, shell=True, creationflags=subprocess.CREATE_NO_WINDOW).decode('utf-8', errors='replace')
            return [line.strip() for line in output.splitlines() if line.strip()]
        except:
            return []

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.left_frame = ctk.CTkFrame(self, width=450)
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        ctk.CTkLabel(self.left_frame, text="NDI Free Audio", font=("Segoe UI", 22, "bold"), text_color="#00FFFF").pack(pady=15)

        self.frame_manage = ctk.CTkFrame(self.left_frame)
        self.frame_manage.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(self.frame_manage, text="Serviços Ativos (Status + Nome)", font=("Segoe UI", 12, "bold")).pack(pady=5)
        self.service_list = self.get_ndi_services()
        self.combo_services = ctk.CTkComboBox(self.frame_manage, values=self.service_list if self.service_list else ["Nenhum serviço"], width=350)
        self.combo_services.pack(pady=5)

        btn_grid = ctk.CTkFrame(self.frame_manage, fg_color="transparent")
        btn_grid.pack(pady=10)
        ctk.CTkButton(btn_grid, text="REINICIAR", width=100, fg_color="#3B8ED0", command=lambda: self.quick_action("restart")).grid(row=0, column=0, padx=2)
        ctk.CTkButton(btn_grid, text="PARAR", width=100, fg_color="#E74C3C", command=lambda: self.quick_action("stop")).grid(row=0, column=1, padx=2)
        ctk.CTkButton(btn_grid, text="EXCLUIR", width=100, fg_color="#8B0000", command=self.delete_service).grid(row=0, column=2, padx=2)
        
        ctk.CTkButton(self.frame_manage, text="ATUALIZAR LISTA", width=310, command=self.refresh_list).pack(pady=5)
        ctk.CTkButton(self.frame_manage, text="REMOVER TODOS OS SERVIÇOS", width=310, fg_color="#FF0000", hover_color="#8B0000", command=self.delete_all_services).pack(pady=(0, 10))

        self.frame_create = ctk.CTkFrame(self.left_frame)
        self.frame_create.pack(fill="both", expand=True, padx=15, pady=10)
        
        ctk.CTkLabel(self.frame_create, text="CONFIGURAR NOVO FLUXO", font=("Segoe UI", 14, "bold"), text_color="green").pack(pady=10)

        self.mode_var = ctk.StringVar(value="input")
        mode_frame = ctk.CTkFrame(self.frame_create, fg_color="transparent")
        mode_frame.pack()
        ctk.CTkRadioButton(mode_frame, text="ENTRADA (PC Local -> NDI)", variable=self.mode_var, value="input", command=self.toggle_mode_fields).pack(side="left", padx=10)
        ctk.CTkRadioButton(mode_frame, text="SAÍDA (NDI -> Som Local)", variable=self.mode_var, value="output", command=self.toggle_mode_fields).pack(side="left", padx=10)

        self.ent_id = ctk.CTkEntry(self.frame_create, placeholder_text="ID do Dispositivo (Input/Output Device)", width=350)
        self.ent_id.pack(pady=5)
        
        self.ent_pc_host = ctk.CTkEntry(self.frame_create, placeholder_text="Nome do PC de Origem (Host)", width=350)
        
        self.ent_name = ctk.CTkEntry(self.frame_create, placeholder_text="Nome do Stream (Ex: MesaSom)", width=350)
        self.ent_name.pack(pady=5)
        
        self.ent_gain = ctk.CTkEntry(self.frame_create, placeholder_text="Ganho (Ex: +10 ou -5)", width=350)
        self.ent_gain.pack(pady=5)

        ctk.CTkButton(self.frame_create, text="CRIAR E INICIAR FLUXO", fg_color="green", height=40, font=("Segoe UI", 13, "bold"), command=self.create_service).pack(pady=20)

        self.btn_vbcable = ctk.CTkButton(self.left_frame, text="INSTALAR/REMOVER VB-CABLE", fg_color="#9B59B6", command=self.manage_vbcable)
        self.btn_vbcable.pack(pady=5, padx=15, fill="x")

        self.btn_firewall = ctk.CTkButton(self.left_frame, text="LIBERAR FIREWALL", fg_color="#27AE60", command=self.fix_firewall)
        self.btn_firewall.pack(pady=5, padx=15, fill="x")
        
        self.btn_services = ctk.CTkButton(self.left_frame, text="ABRIR SERVICES.MSC", fg_color="gray", command=lambda: os.startfile("services.msc"))
        self.btn_services.pack(pady=5, padx=15, fill="x")

        self.right_frame = ctk.CTkFrame(self)
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        ctk.CTkLabel(self.right_frame, text="LISTA DE DISPOSITIVOS DE ÁUDIO (IDs)", font=("Segoe UI", 14, "bold")).pack(pady=10)
        self.txt_devices = ctk.CTkTextbox(self.right_frame, font=("Consolas", 11))
        self.txt_devices.pack(fill="both", expand=True, padx=10, pady=5)
        ctk.CTkButton(self.right_frame, text="ESCANEAR DISPOSITIVOS", command=self.scan_devices).pack(pady=10)

    def manage_vbcable(self):
        if os.path.exists(self.vbcable_exe):
            try:
                subprocess.Popen([self.vbcable_exe], shell=True)
                messagebox.showinfo("VB-CABLE", "Instalador iniciado.")
            except:
                messagebox.showerror("Erro", "Não foi possível iniciar.")
        else:
            messagebox.showerror("Erro", "Arquivo não encontrado.")

    def toggle_mode_fields(self):
        if self.mode_var.get() == "output":
            self.ent_pc_host.pack(after=self.ent_id, pady=5)
        else:
            self.ent_pc_host.pack_forget()

    def scan_devices(self):
        self.txt_devices.delete("1.0", "end")
        self.txt_devices.insert("end", "Obtendo IDs...\n")
        try:
            result = subprocess.run([self.app_exe], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            full_text = result.stdout.decode('cp1252', errors='replace')
            if "Input Devices:" in full_text:
                filtered_text = full_text.split("Input Devices:", 1)[1]
                self.txt_devices.delete("1.0", "end")
                self.txt_devices.insert("end", "Input Devices:" + filtered_text)
            else:
                self.txt_devices.delete("1.0", "end")
                self.txt_devices.insert("end", full_text)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao escanear: {e}")

    def create_service(self):
        mode = self.mode_var.get()
        id_val = self.ent_id.get()
        name_val = self.ent_name.get()
        host_val = self.ent_pc_host.get()
        gain_val = self.ent_gain.get()

        if not id_val or not name_val:
            messagebox.showwarning("Erro", "Preencha ID e Nome.")
            return

        s_name = f"NDI Audio - ID {id_val} - {'Output' if mode == 'output' else 'Input'} - {name_val}"
        
        if mode == "output":
            args = f"-output {id_val} -output_name \"{host_val} ({name_val})\""
        else:
            args = f"-input {id_val} -input_name \"{name_val}\""
            
        if gain_val:
            args += f" {'-output_gain' if mode=='output' else '-input_gain'} {gain_val}{'dB' if 'dB' not in gain_val else ''}"

        try:
            subprocess.run([self.nssm, "install", s_name, self.app_exe, args], creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run([self.nssm, "set", s_name, "AppDirectory", self.root_path], creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run(f"powershell Start-Service -Name '{s_name}'", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            messagebox.showinfo("Sucesso", "Serviço criado!")
            self.refresh_list()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def fix_firewall(self):
        ps_cmd = 'New-NetFirewallRule -DisplayName "NDI Free Audio - Entrada" -Direction Inbound -Program "' + self.app_exe + '" -Action Allow; ' \
                 'New-NetFirewallRule -DisplayName "NDI Free Audio - Saída" -Direction Outbound -Program "' + self.app_exe + '" -Action Allow'
        subprocess.run(["powershell", ps_cmd], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("Firewall", "Regras atualizadas.")

    def extract_service_name(self, display_name):
        # Remove o prefixo '[Status] ' para obter apenas o nome real do serviço
        if "] " in display_name:
            return display_name.split("] ", 1)[1]
        return display_name

    def quick_action(self, action):
        selection = self.combo_services.get()
        if selection and selection != "Nenhum serviço":
            srv = self.extract_service_name(selection)
            subprocess.run(f"powershell {action}-Service -Name '{srv}'", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.refresh_list()
            messagebox.showinfo("Status", f"Serviço {action} executado.")

    def delete_service(self):
        selection = self.combo_services.get()
        if selection and selection != "Nenhum serviço":
            srv = self.extract_service_name(selection)
            if messagebox.askyesno("Confirmar", f"Excluir {srv}?"):
                subprocess.run(f'"{self.nssm}" stop "{srv}"', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.run(f'"{self.nssm}" remove "{srv}" confirm', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                self.refresh_list()

    def delete_all_services(self):
        # Obtém a lista bruta de serviços (apenas nomes)
        cmd = 'powershell -NoProfile -Command "Get-Service | Where-Object { $_.Name -like \'NDI Audio*\' } | Select-Object -ExpandProperty Name"'
        try:
            output = subprocess.check_output(cmd, shell=True, creationflags=subprocess.CREATE_NO_WINDOW).decode('utf-8', errors='replace')
            services = [line.strip() for line in output.splitlines() if line.strip()]
            
            if not services:
                messagebox.showinfo("Info", "Nenhum serviço NDI encontrado para remover.")
                return

            if messagebox.askyesno("AVISO CRÍTICO", f"Isso removerá TODOS os {len(services)} serviços NDI criados. Confirmar?"):
                for srv in services:
                    subprocess.run(f'"{self.nssm}" stop "{srv}"', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    subprocess.run(f'"{self.nssm}" remove "{srv}" confirm', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                
                self.refresh_list()
                messagebox.showinfo("Sucesso", "Todos os serviços foram removidos.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao remover todos: {e}")

    def refresh_list(self):
        new_list = self.get_ndi_services()
        self.combo_services.configure(values=new_list if new_list else ["Nenhum serviço"])
        if new_list: self.combo_services.set(new_list[0])
        else: self.combo_services.set("Nenhum serviço")

if __name__ == "__main__":
    app = NDIPanelApp()
    app.mainloop()