import os
import sys
import subprocess
import ctypes
import customtkinter as ctk
from tkinter import messagebox

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
        
        self.title("Painel NDI Free Audio")
        self.geometry("900x900") 
        
        if not self.is_admin():
            messagebox.showerror("Erro", "Execute como ADMINISTRADOR!")
            self.destroy()
            return

        self.setup_ui()

    def is_admin(self):
        try: return ctypes.windll.shell32.IsUserAnAdmin()
        except: return False

    def test_connection(self, host):
        response = subprocess.run(["ping", "-n", "1", host], 
                                  capture_output=True, 
                                  creationflags=subprocess.CREATE_NO_WINDOW)
        return response.returncode == 0

    def get_ndi_services(self):
        try:
            cmd = 'powershell "Get-Service | Where-Object { $_.Name -like \'NDI Audio*\' } | Select-Object -ExpandProperty Name"'
            output = subprocess.check_output(cmd, shell=True, creationflags=subprocess.CREATE_NO_WINDOW).decode('latin-1')
            return [line.strip() for line in output.splitlines() if line.strip()]
        except: return []

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.left_frame = ctk.CTkFrame(self, width=450)
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        ctk.CTkLabel(self.left_frame, text="NDI Free Audio", font=("Segoe UI", 22, "bold"), text_color="#00FFFF").pack(pady=15)

        self.frame_manage = ctk.CTkFrame(self.left_frame)
        self.frame_manage.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(self.frame_manage, text="Servicos Ativos", font=("Segoe UI", 12, "bold")).pack(pady=5)
        self.service_list = self.get_ndi_services()
        self.combo_services = ctk.CTkComboBox(self.frame_manage, values=self.service_list if self.service_list else ["Nenhum servico"], width=350)
        self.combo_services.pack(pady=5)

        btn_grid = ctk.CTkFrame(self.frame_manage, fg_color="transparent")
        btn_grid.pack(pady=10)
        ctk.CTkButton(btn_grid, text="REINICIAR", width=100, fg_color="#3B8ED0", command=lambda: self.quick_action("restart")).grid(row=0, column=0, padx=2)
        ctk.CTkButton(btn_grid, text="PARAR", width=100, fg_color="#E74C3C", command=lambda: self.quick_action("stop")).grid(row=0, column=1, padx=2)
        ctk.CTkButton(btn_grid, text="EXCLUIR", width=100, fg_color="#8B0000", command=self.delete_service).grid(row=0, column=2, padx=2)
        ctk.CTkButton(self.frame_manage, text="ATUALIZAR LISTA", width=310, command=self.refresh_list).pack(pady=(0, 10))

        self.frame_create = ctk.CTkFrame(self.left_frame)
        self.frame_create.pack(fill="both", expand=True, padx=15, pady=10)
        
        ctk.CTkLabel(self.frame_create, text="CONFIGURAR NOVO FLUXO", font=("Segoe UI", 14, "bold"), text_color="green").pack(pady=10)

        self.mode_var = ctk.StringVar(value="input")
        mode_frame = ctk.CTkFrame(self.frame_create, fg_color="transparent")
        mode_frame.pack()
        ctk.CTkRadioButton(mode_frame, text="ENTRADA (PC Local -> NDI)", variable=self.mode_var, value="input", command=self.toggle_mode_fields).pack(side="left", padx=10)
        ctk.CTkRadioButton(mode_frame, text="SAIDA (NDI -> Som Local)", variable=self.mode_var, value="output", command=self.toggle_mode_fields).pack(side="left", padx=10)

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
                messagebox.showinfo("VB-CABLE", "O instalador do VB-CABLE foi iniciado.\n\nSiga as instrucoes na janela do driver para Instalar ou Remover.")
            except Exception as e:
                messagebox.showerror("Erro", f"Nao foi possivel iniciar o instalador: {e}")
        else:
            messagebox.showerror("Erro", f"Arquivo nao encontrado em:\n{self.vbcable_exe}\n\nVerifique se a pasta VBCABLE existe no local do app.")

    def toggle_mode_fields(self):
        if self.mode_var.get() == "output":
            self.ent_pc_host.pack(after=self.ent_id, pady=5)
        else:
            self.ent_pc_host.pack_forget()

    def scan_devices(self):
        self.txt_devices.delete("1.0", "end")
        self.txt_devices.insert("end", "Obtendo IDs de audio...\n")
        try:
            result = subprocess.run([self.app_exe], capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            full_text = result.stdout if result.stdout else result.stderr
            if "Input Devices:" in full_text:
                filtered_text = full_text.split("Input Devices:", 1)[1]
                self.txt_devices.delete("1.0", "end")
                self.txt_devices.insert("end", "Input Devices:" + filtered_text)
            else:
                self.txt_devices.delete("1.0", "end")
                self.txt_devices.insert("end", full_text)
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel ler os dispositivos: {e}")

    def create_service(self):
        mode = self.mode_var.get()
        id_val = self.ent_id.get()
        name_val = self.ent_name.get()
        host_val = self.ent_pc_host.get()
        gain_val = self.ent_gain.get()

        if not id_val or not name_val:
            messagebox.showwarning("Erro", "Preencha ID e Nome do Stream.")
            return

        if mode == "output":
            if not host_val:
                messagebox.showwarning("Erro", "Para saida, informe o nome do PC de origem (Host).")
                return
            if not self.test_connection(host_val):
                messagebox.showwarning("Aviso de Conexao", f"O host {host_val} parece estar inacessivel. Verifique o nome e tente novamente.")
                return

            s_name = f"NDI Audio - ID {id_val} - Output - {name_val}"
            full_ndi_source = f"{host_val} ({name_val})"
            args = f"-output {id_val} -output_name \"{full_ndi_source}\""
            if gain_val: args += f" -output_gain {gain_val if 'dB' in gain_val else gain_val + 'dB'}"
        else:
            s_name = f"NDI Audio - ID {id_val} - Input - {name_val}"
            args = f"-input {id_val} -input_name \"{name_val}\""
            if gain_val: args += f" -input_gain {gain_val if 'dB' in gain_val else gain_val + 'dB'}"

        try:
            subprocess.run([self.nssm, "install", s_name, self.app_exe, args], creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run([self.nssm, "set", s_name, "AppDirectory", self.root_path], creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run(f"powershell Start-Service -Name '{s_name}'", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            messagebox.showinfo("Sucesso", "Fluxo criado e iniciado com sucesso!")
            self.refresh_list()
        except Exception as e:
            messagebox.showerror("Erro NSSM", str(e))

    def fix_firewall(self):
        ps_cmd = f'Remove-NetFirewallRule -DisplayName "NDI Free Audio - Entrada" -ErrorAction SilentlyContinue; ' \
                 f'Remove-NetFirewallRule -DisplayName "NDI Free Audio - Saida" -ErrorAction SilentlyContinue; ' \
                 f'New-NetFirewallRule -DisplayName "NDI Free Audio - Entrada" -Direction Inbound -Program "{self.app_exe}" -Action Allow; ' \
                 f'New-NetFirewallRule -DisplayName "NDI Free Audio - Saida" -Direction Outbound -Program "{self.app_exe}" -Action Allow'
        subprocess.run(["powershell", ps_cmd], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("Firewall", "Regras de rede atualizadas!")

    def quick_action(self, action):
        srv = self.combo_services.get()
        if srv and srv != "Nenhum servico":
            subprocess.run(f"powershell {action}-Service -Name '{srv}'", shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            messagebox.showinfo("Status", f"Servico {action} executado.")

    def delete_service(self):
        srv = self.combo_services.get()
        if srv and srv != "Nenhum servico":
            if messagebox.askyesno("Confirmar", f"Excluir o servico {srv}?"):
                subprocess.run(f'"{self.nssm}" stop "{srv}"', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.run(f'"{self.nssm}" remove "{srv}" confirm', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                self.refresh_list()

    def refresh_list(self):
        new_list = self.get_ndi_services()
        self.combo_services.configure(values=new_list if new_list else ["Nenhum servico"])
        if new_list: self.combo_services.set(new_list[0])
        else: self.combo_services.set("Nenhum servico")

if __name__ == "__main__":
    app = NDIPanelApp()
    app.mainloop()