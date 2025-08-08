import tkinter as tk
from tkinter import ttk
import os
import shutil
from datetime import datetime, timedelta
import locale
import zipfile
import subprocess
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Agendador e Trabalhador de NFEs")
        self.geometry("800x600")

        self.create_widgets()

    def create_widgets(self):
        # Notebook para abas
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Aba de Configurações
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Configurações")
        self.create_settings_tab(self.settings_frame)

        # Aba de Execução
        self.execution_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.execution_frame, text="Execução")
        self.create_execution_tab(self.execution_frame)

        # Aba de Agendamento
        self.schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.schedule_frame, text="Agendamento")
        self.create_schedule_tab(self.schedule_frame)

    def create_settings_tab(self, parent_frame):
        # Exemplo de campo de configuração
        ttk.Label(parent_frame, text="Pasta Origem:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.pasta_origem_entry = ttk.Entry(parent_frame, width=50)
        self.pasta_origem_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.pasta_origem_entry.insert(0, "C:\\Mobility_POS\\Xml_IO") # Valor padrão do script

        ttk.Label(parent_frame, text="Pasta Destino Base:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.pasta_destino_base_entry = ttk.Entry(parent_frame, width=50)
        self.pasta_destino_base_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.pasta_destino_base_entry.insert(0, "C:\\CopiaNotasFiscais")

        ttk.Label(parent_frame, text="Caminho Rclone:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.rclone_path_entry = ttk.Entry(parent_frame, width=50)
        self.rclone_path_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.rclone_path_entry.insert(0, "C:\\Ferramentas\\rclone\\rclone.exe")

        ttk.Label(parent_frame, text="Nome Remoto Rclone:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.rclone_remote_name_entry = ttk.Entry(parent_frame, width=50)
        self.rclone_remote_name_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.rclone_remote_name_entry.insert(0, "MeuGoogleDrive")

        ttk.Label(parent_frame, text="Pasta Base Drive:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.pasta_base_drive_entry = ttk.Entry(parent_frame, width=50)
        self.pasta_base_drive_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.pasta_base_drive_entry.insert(0, "CLIENTES")

        ttk.Label(parent_frame, text="Nome Cliente Específico:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.nome_cliente_especifico_entry = ttk.Entry(parent_frame, width=50)
        self.nome_cliente_especifico_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.nome_cliente_especifico_entry.insert(0, "PIE HOUSE")

        # Configurações de E-mail
        ttk.Label(parent_frame, text="--- Configurações de E-mail ---").grid(row=6, column=0, columnspan=2, padx=5, pady=10, sticky="w")

        ttk.Label(parent_frame, text="Servidor SMTP:").grid(row=7, column=0, padx=5, pady=5, sticky="w")
        self.smtp_server_entry = ttk.Entry(parent_frame, width=50)
        self.smtp_server_entry.grid(row=7, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_server_entry.insert(0, "smtp.gmail.com")

        ttk.Label(parent_frame, text="Porta SMTP:").grid(row=8, column=0, padx=5, pady=5, sticky="w")
        self.smtp_port_entry = ttk.Entry(parent_frame, width=50)
        self.smtp_port_entry.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_port_entry.insert(0, "587")

        ttk.Label(parent_frame, text="Usuário SMTP:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
        self.smtp_username_entry = ttk.Entry(parent_frame, width=50)
        self.smtp_username_entry.grid(row=9, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_username_entry.insert(0, "solracinformatica@gmail.com")

        ttk.Label(parent_frame, text="Senha SMTP (App): ").grid(row=10, column=0, padx=5, pady=5, sticky="w")
        self.smtp_password_entry = ttk.Entry(parent_frame, width=50, show="*") # Senha oculta
        self.smtp_password_entry.grid(row=10, column=1, padx=5, pady=5, sticky="ew")
        self.smtp_password_entry.insert(0, "yuue eaqj kitd ohrq")

        ttk.Label(parent_frame, text="E-mail De:").grid(row=11, column=0, padx=5, pady=5, sticky="w")
        self.email_from_entry = ttk.Entry(parent_frame, width=50)
        self.email_from_entry.grid(row=11, column=1, padx=5, pady=5, sticky="ew")
        self.email_from_entry.insert(0, "solracinformatica@gmail.com")

        ttk.Label(parent_frame, text="E-mail Para:").grid(row=12, column=0, padx=5, pady=5, sticky="w")
        self.email_to_entry = ttk.Entry(parent_frame, width=50)
        self.email_to_entry.grid(row=12, column=1, padx=5, pady=5, sticky="ew")
        self.email_to_entry.insert(0, "solracinformatica@gmail.com")

        # Opções de ativação/desativação de funcionalidades
        ttk.Label(parent_frame, text="--- Opções de Funcionalidades ---").grid(row=13, column=0, columnspan=2, padx=5, pady=10, sticky="w")

        self.enable_email_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent_frame, text="Ativar Notificações por E-mail", variable=self.enable_email_var).grid(row=14, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        self.enable_rclone_upload_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent_frame, text="Ativar Upload para Google Drive (rclone)", variable=self.enable_rclone_upload_var).grid(row=15, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        self.enable_prerequisites_check_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent_frame, text="Ativar Verificação de Pré-requisitos", variable=self.enable_prerequisites_check_var).grid(row=16, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        parent_frame.columnconfigure(1, weight=1) # Faz a coluna 1 expandir

    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def get_settings(self):
        settings = {
            "pasta_origem": self.pasta_origem_entry.get(),
            "pasta_destino_base": self.pasta_destino_base_entry.get(),
            "rclone_path": self.rclone_path_entry.get(),
            "rclone_remote_name": self.rclone_remote_name_entry.get(),
            "pasta_base_drive": self.pasta_base_drive_entry.get(),
            "nome_cliente_especifico": self.nome_cliente_especifico_entry.get(),
            "smtp_server": self.smtp_server_entry.get(),
            "smtp_port": int(self.smtp_port_entry.get()),
            "smtp_username": self.smtp_username_entry.get(),
            "smtp_password": self.smtp_password_entry.get(),
            "email_from": self.email_from_entry.get(),
            "email_to": self.email_to_entry.get(),
            "enable_email": self.enable_email_var.get(),
            "enable_rclone_upload": self.enable_rclone_upload_var.get(),
            "enable_prerequisites_check": self.enable_prerequisites_check_var.get(),
        }
        return settings

    def send_script_email(self, subject, body, attachment_path=None):
        if not self.enable_email_var.get():
            self.log_message("Envio de e-mail desativado nas configurações.")
            return

        settings = self.get_settings()
        try:
            msg = MIMEMultipart()
            msg["From"] = settings["email_from"]
            msg["To"] = settings["email_to"]
            msg["Subject"] = subject

            msg.attach(MIMEText(body, "plain"))

            if attachment_path and os.path.exists(attachment_path):
                file_size_mb = os.path.getsize(attachment_path) / (1024 * 1024)
                if file_size_mb <= 25: # Limite do Gmail
                    with open(attachment_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
                    msg.attach(part)
                else:
                    self.log_message(f"AVISO: Arquivo de log muito grande ({file_size_mb:.2f}MB) para anexar ao e-mail.")
                    body += f"\n\nNOTA: Arquivo de log muito grande para anexar. Verifique o log localmente em: {attachment_path}"
                    msg = MIMEMultipart()
                    msg["From"] = settings["email_from"]
                    msg["To"] = settings["email_to"]
                    msg["Subject"] = subject
                    msg.attach(MIMEText(body, "plain"))

            with smtplib.SMTP(settings["smtp_server"], settings["smtp_port"]) as server:
                server.starttls()
                server.login(settings["smtp_username"], settings["smtp_password"])
                server.sendmail(settings["email_from"], settings["email_to"], msg.as_string())
            self.log_message(f"E-mail de notificação enviado com sucesso: {subject}")
        except Exception as e:
            self.log_message(f"ERRO ao enviar e-mail: {e}")

    def execute_backup(self):
        self.log_text.delete(1.0, tk.END) # Limpa o log anterior
        self.log_message("--- INICIANDO EXECUÇÃO DA LÓGICA DO TRABALHADOR ---")
        settings = self.get_settings()

        script_status = "SUCESSO"
        error_messages = []
        start_time = datetime.now()

        # Configurações de caminho e nome de arquivo de log
        log_file = os.path.join(settings["pasta_destino_base"], "log_copia_nfe.log")
        if not os.path.exists(os.path.dirname(log_file)):
            os.makedirs(os.path.dirname(log_file))

        try:
            # Pré-requisitos
            if settings["enable_prerequisites_check"]:
                self.log_message("Verificando pré-requisitos...")
                if not os.path.exists(settings["rclone_path"]):
                    error_messages.append(f"rclone não encontrado em \'{settings['rclone_path']}\'")
                if not os.path.exists(settings["pasta_origem"]):
                    error_messages.append(f"Pasta origem não encontrada: \'{settings['pasta_origem']}\'")
                
                # Adicionar verificação de rclone.conf e remote
                config_path = os.path.join(os.path.dirname(settings["rclone_path"]), "rclone.conf")
                if os.path.exists(settings["rclone_path"]) and not os.path.exists(config_path):
                    error_messages.append(f"Arquivo de configuração do rclone não encontrado: \'{config_path}\'")
                
                if os.path.exists(settings["rclone_path"]) and os.path.exists(config_path):
                    try:
                        # Use subprocess.run para capturar a saída
                        result = subprocess.run([settings["rclone_path"], "listremotes", "--config", config_path], capture_output=True, text=True, check=True)
                        if f"{settings['rclone_remote_name']}:\n" not in result.stdout:
                            error_messages.append(f"Remote \'{settings['rclone_remote_name']}\' não encontrado na configuração do rclone")
                    except subprocess.CalledProcessError as e:
                        error_messages.append(f"Erro ao verificar configuração do rclone: {e.stderr}")
                    except FileNotFoundError:
                        error_messages.append(f"Erro: rclone não encontrado no caminho especificado: {settings['rclone_path']}")

                if error_messages:
                    script_status = "FALHA"
                    raise Exception("Pré-requisitos não atendidos. Verifique o log para detalhes.")
                self.log_message("Pré-requisitos verificados com sucesso.")

            # Lógica de cópia de arquivos
            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8') # Para nomes de meses em português
            mes_de_referencia = datetime.now().replace(day=1) # Mês atual
            # Para o mês anterior, descomente a linha abaixo e comente a de cima
            # mes_de_referencia = (datetime.now().replace(day=1) - timedelta(days=1)).replace(day=1)

            primeiro_dia_mes_referencia = mes_de_referencia.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            ultimo_dia_mes_referencia = (mes_de_referencia.replace(month=mes_de_referencia.month % 12 + 1, day=1) - timedelta(microseconds=1))

            self.log_message(f"Período de busca dos arquivos: de {primeiro_dia_mes_referencia.strftime('%d/%m/%Y')} a {ultimo_dia_mes_referencia.strftime('%d/%m/%Y')}.")

            nome_mes_referencia = mes_de_referencia.strftime('%B').upper()
            nome_pasta_destino_local = f"{mes_de_referencia.strftime('%Y-%m')}_{nome_mes_referencia}"
            pasta_destino_completa = os.path.join(settings["pasta_destino_base"], nome_pasta_destino_local)

            if not os.path.exists(pasta_destino_completa):
                self.log_message(f"Criando pasta de destino local: {pasta_destino_completa}")
                os.makedirs(pasta_destino_completa)

            self.log_message(f"Procurando arquivos .xml em \'{settings['pasta_origem']}\'...")
            arquivos_para_copiar = []
            for root, _, files in os.walk(settings["pasta_origem"]):
                for file in files:
                    if file.endswith(".xml"):
                        file_path = os.path.join(root, file)
                        mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                        if primeiro_dia_mes_referencia <= mod_time <= ultimo_dia_mes_referencia:
                            arquivos_para_copiar.append(file_path)

            if arquivos_para_copiar:
                self.log_message(f"Foram encontrados {len(arquivos_para_copiar)} arquivos para copiar.")
                arquivos_copiados_com_sucesso = 0
                for arquivo in arquivos_para_copiar:
                    try:
                        shutil.copy(arquivo, pasta_destino_completa)
                        arquivos_copiados_com_sucesso += 1
                        self.log_message(f"Copiado: {os.path.basename(arquivo)}")
                    except Exception as e:
                        error_messages.append(f"Erro ao copiar \'{os.path.basename(arquivo)}\': {e}")
                        self.log_message(f"ERRO ao copiar \'{os.path.basename(arquivo)}\': {e}")

                if arquivos_copiados_com_sucesso == 0:
                    script_status = "FALHA"
                    error_messages.append("Nenhum arquivo foi copiado com sucesso")
                    raise Exception("Nenhum arquivo foi copiado com sucesso")

                self.log_message(f"Total de arquivos copiados com sucesso: {arquivos_copiados_com_sucesso} de {len(arquivos_para_copiar)}")

                # Lógica de compactação
                nome_arquivo_zip = f"NFEs_{mes_de_referencia.strftime('%b').upper()}_{mes_de_referencia.strftime('%Y')}.zip"
                caminho_arquivo_zip = os.path.join(settings["pasta_destino_base"], nome_arquivo_zip)

                if os.path.exists(caminho_arquivo_zip):
                    self.log_message(f"Arquivo .zip \'{caminho_arquivo_zip}\' já existe. Removendo para criar um novo.")
                    os.remove(caminho_arquivo_zip)

                self.log_message(f"Iniciando a compactação para \'{caminho_arquivo_zip}\'...")
                try:
                    with zipfile.ZipFile(caminho_arquivo_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for folderName, subfolders, filenames in os.walk(pasta_destino_completa):
                            for filename in filenames:
                                if filename.endswith(".xml"):
                                    file_path = os.path.join(folderName, filename)
                                    zipf.write(file_path, os.path.basename(file_path)) # Adiciona apenas o nome do arquivo no zip
                    self.log_message("Compactação concluída com sucesso.")
                except Exception as e:
                    script_status = "FALHA"
                    error_messages.append(f"Falha na compactação: {e}")
                    raise Exception(f"Erro na compactação: {e}")

                # Lógica de upload rclone
                if settings["enable_rclone_upload"]:
                    if os.path.exists(caminho_arquivo_zip):
                        file_size = round(os.path.getsize(caminho_arquivo_zip) / (1024 * 1024), 2)
                        self.log_message(f"Arquivo compactado criado ({file_size} MB). Iniciando upload para o Google Drive...")

                        caminho_destino_drive = f"{settings['pasta_base_drive']}/{settings['nome_cliente_especifico']}/{nome_pasta_destino_local}/"
                        self.log_message(f"Destino no Drive: \'{settings['rclone_remote_name']}:{caminho_destino_drive}\'")

                        config_path = os.path.join(os.path.dirname(settings["rclone_path"]), "rclone.conf")
                        argumentos = [
                            "copy",
                            caminho_arquivo_zip,
                            f"{settings['rclone_remote_name']}:{caminho_destino_drive}",
                            "--config",
                            config_path,
                            "--progress",
                            "--transfers", "1",
                            "--checkers", "1"
                        ]

                        self.log_message(f"Executando rclone: {settings['rclone_path']} {' '.join(argumentos)}")
                        try:
                            process = subprocess.run([settings["rclone_path"]] + argumentos, capture_output=True, text=True, check=True)
                            self.log_message(f"SUCESSO: Upload do arquivo \'{nome_arquivo_zip}\' concluído.")
                            self.log_message(process.stdout)
                        except subprocess.CalledProcessError as e:
                            script_status = "FALHA"
                            error_messages.append(f"Upload com rclone falhou (código: {e.returncode}). Erro: {e.stderr}")
                            self.log_message(f"ERRO: Upload com rclone falhou. Código: {e.returncode}. Erro: {e.stderr}")
                        except FileNotFoundError:
                            script_status = "FALHA"
                            error_messages.append(f"Erro: rclone não encontrado no caminho especificado: {settings['rclone_path']}")
                            self.log_message(f"ERRO: rclone não encontrado no caminho especificado: {settings['rclone_path']}")
                    else:
                        script_status = "FALHA"
                        error_messages.append("Falha ao criar o arquivo compactado")
                        self.log_message("ERRO: Falha ao criar o arquivo compactado. O upload não será realizado.")
                else:
                    self.log_message("Upload para Google Drive (rclone) desativado nas configurações.")

            else:
                self.log_message("Nenhum arquivo .xml do mês anterior foi encontrado para cópia.")

        except Exception as e:
            script_status = "FALHA"
            error_messages.append(f"ERRO INESPERADO: {e}")
            self.log_message(f"ERRO INESPERADO: {e}")
        finally:
            end_time = datetime.now()
            duration = end_time - start_time
            self.log_message(f"Duração total da execução: {duration}")
            self.log_message("----------------- EXECUÇÃO DA LÓGICA DO TRABALHADOR FINALIZADA ---")

            # Enviar e-mail de notificação
            body_details = f"""
Computador: {os.environ.get('COMPUTERNAME', 'N/A')}
Usuário: {os.environ.get('USERNAME', 'N/A')}
Horário de início: {start_time.strftime('%d/%m/%Y %H:%M:%S')}
Horário de término: {end_time.strftime('%d/%m/%Y %H:%M:%S')}
Duração: {duration}
Cliente: {settings['nome_cliente_especifico']}

"""

            if script_status == "SUCESSO":
                subject = f"SUCESSO: Copia de NFEs - {settings['nome_cliente_especifico']} - {os.environ.get('COMPUTERNAME', 'N/A')}"
                body_success = body_details + "A cópia e upload de NFEs foi concluída com sucesso!"
                self.send_script_email(subject, body_success)
            else:
                subject = f"FALHA: Copia de NFEs - {settings['nome_cliente_especifico']} - {os.environ.get('COMPUTERNAME', 'N/A')}"
                body_failure = body_details + "FALHAS DETECTADAS:\n\n" + ("\n".join(error_messages))
                self.send_script_email(subject, body_failure, log_file)

    def create_execution_tab(self, parent_frame):
        ttk.Label(parent_frame, text="Status da Execução:").pack(padx=5, pady=5, anchor="nw")
        self.log_text = tk.Text(parent_frame, wrap="word", height=20, width=80)
        self.log_text.pack(padx=5, pady=5, fill="both", expand=True)
        self.log_text.insert(tk.END, "Logs de execução aparecerão aqui...")

        ttk.Button(parent_frame, text="Executar Backup Agora", command=self.execute_backup).pack(pady=10)

    def create_schedule_tab(self, parent_frame):
        ttk.Label(parent_frame, text="Configurações de Agendamento:").pack(padx=5, pady=5, anchor="nw")
        self.enable_schedule_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(parent_frame, text="Ativar Agendamento Mensal", variable=self.enable_schedule_var).pack(padx=5, pady=5, anchor="nw")
        ttk.Button(parent_frame, text="Agendar Tarefa", command=self.schedule_task).pack(pady=10)

if __name__ == "__main__":
    app = App()
    app.mainloop()



    def schedule_task(self):
        settings = self.get_settings()
        task_name = "Copia Mensal de NFEs para Nuvem"
        script_path = os.path.abspath(__file__) # Caminho do script Python atual

        # Comando para executar o script Python com o interpretador Python
        # Usamos `sys.executable` para garantir que o mesmo interpretador Python que está executando o app seja usado.
        # O parâmetro `-SomenteExecutar` no script PowerShell original indicava para não agendar novamente.
        # No nosso caso, o app Python sempre executa a lógica de backup/upload e o agendamento é uma função separada.
        # Então, o comando agendado deve chamar o app Python para executar a função de backup.
        # Para simplificar, vamos criar um script wrapper PowerShell que chama o app Python.
        # Ou, podemos agendar diretamente o comando python.exe -m PyInstaller --onefile --windowed app.py
        # Para o agendamento, vamos criar um script PowerShell que executa a função de backup do nosso app.
        # Isso exigiria que o app estivesse em um estado que pudesse ser chamado via linha de comando para a função de backup.
        # Uma abordagem mais simples para agendamento é criar um script PowerShell separado que chama o app Python.
        # Ou, se o app for empacotado, o executável pode ser agendado.

        # Por enquanto, vamos simular o agendamento de um comando simples para fins de demonstração.
        # A implementação real dependeria de como o aplicativo será empacotado e distribuído.
        # Para um aplicativo Python, o ideal seria agendar a execução do executável gerado pelo PyInstaller.
        # Ou, se for um script, agendar 'pythonw.exe <caminho_do_script.py>'

        # Para este exemplo, vamos agendar um comando PowerShell que executa um script Python.
        # Este script PowerShell precisaria ser criado e salvo em algum lugar.
        # Vamos criar um script PowerShell temporário para isso.

        powershell_script_content = f"""
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$pythonAppPath = Join-Path $scriptDir "app.py" # Assumindo que app.py estará no mesmo diretório
$logFile = "{os.path.join(settings["pasta_destino_base"], "log_copia_nfe.log")}"

# Redirecionar stdout e stderr para o log
Start-Process -FilePath "python.exe" -ArgumentList "$pythonAppPath", "--execute-backup-from-scheduler" -NoNewWindow -RedirectStandardOutput $logFile -RedirectStandardError $logFile -Wait
"""
        
        # Salvar o script PowerShell temporário
        temp_ps_script_path = os.path.join(settings["pasta_destino_base"], "run_backup.ps1")
        with open(temp_ps_script_path, "w") as f:
            f.write(powershell_script_content)

        command_to_schedule = f"powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"{temp_ps_script_path}\""

        try:
            # Verifica se a tarefa já existe
            check_task_command = f"schtasks /query /TN \"{task_name}\""
            result = subprocess.run(check_task_command, capture_output=True, text=True, check=False, shell=True)

            if task_name in result.stdout:
                self.log_message(f"AVISO: A tarefa \'{task_name}\' já existe. Não será criada novamente.")
                self.log_message(f"Para recriar a tarefa, delete-a primeiro com: schtasks /delete /tn \"{task_name}\" /f")
            else:
                self.log_message(f"Criando tarefa \'{task_name}\' para executar \'{command_to_schedule}\'...")
                
                # CORREÇÃO CRÍTICA: Comando com parâmetro -SomenteExecutar para evitar loop infinito
                # No nosso caso, o script Python não agendaria novamente, então não precisamos do -SomenteExecutar
                # A tarefa agendada chamará diretamente a função de backup do app Python.

                # Criar tarefa usando schtasks
                create_task_command = f"schtasks /Create /TN \"{task_name}\" /TR \"{command_to_schedule}\" /SC MONTHLY /D 1 /ST \"10:00\" /RU \"SYSTEM\" /RL HIGHEST /F"
                result = subprocess.run(create_task_command, capture_output=True, text=True, check=False, shell=True)
                
                if result.returncode == 0:
                    self.log_message(f"SUCESSO: Tarefa \'{task_name}\' foi agendada com sucesso.")
                    self.log_message("Ela será executada todo dia 1 de cada mês às 10:00.")
                    self.log_message(f"Para testar a tarefa manualmente: schtasks /run /tn \"{task_name}\"")
                else:
                    self.log_message(f"ERRO: Falha ao criar a tarefa. Código de saída: {result.returncode}")
                    self.log_message(f"Detalhes: {result.stderr}")
        except Exception as e:
            self.log_message(f"ERRO ao tentar agendar a tarefa: {e}")

    def schedule_task(self):
        settings = self.get_settings()
        task_name = "Copia Mensal de NFEs para Nuvem"
        script_path = os.path.abspath(__file__)

        powershell_script_content = f"""
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$pythonAppPath = Join-Path $scriptDir "app.py"
$logFile = "{os.path.join(settings["pasta_destino_base"], "log_copia_nfe.log")}"

Start-Process -FilePath "python.exe" -ArgumentList "$pythonAppPath", "--execute-backup-from-scheduler" -NoNewWindow -RedirectStandardOutput $logFile -RedirectStandardError $logFile -Wait
"""

        temp_ps_script_path = os.path.join(settings["pasta_destino_base"], "run_backup.ps1")
        with open(temp_ps_script_path, "w") as f:
            f.write(powershell_script_content)

        command_to_schedule = f"powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"{temp_ps_script_path}\""

        try:
            check_task_command = f"schtasks /query /TN \"{task_name}\""
            result = subprocess.run(check_task_command, capture_output=True, text=True, check=False, shell=True)

            if task_name in result.stdout:
                self.log_message(f"AVISO: A tarefa '{task_name}' já existe. Não será criada novamente.")
                self.log_message(f"Para recriar a tarefa, delete-a primeiro com: schtasks /delete /tn \"{task_name}\" /f")
            else:
                self.log_message(f"Criando tarefa '{task_name}' para executar '{command_to_schedule}'...")

                create_task_command = f"schtasks /Create /TN \"{task_name}\" /TR \"{command_to_schedule}\" /SC MONTHLY /D 1 /ST \"10:00\" /RU \"SYSTEM\" /RL HIGHEST /F"
                result = subprocess.run(create_task_command, capture_output=True, text=True, check=False, shell=True)

                if result.returncode == 0:
                    self.log_message(f"SUCESSO: Tarefa '{task_name}' foi agendada com sucesso.")
                    self.log_message("Ela será executada todo dia 1 de cada mês às 10:00.")
                    self.log_message(f"Para testar a tarefa manualmente: schtasks /run /tn \"{task_name}\"")
                else:
                    self.log_message(f"ERRO: Falha ao criar a tarefa. Código de saída: {result.returncode}")
                    self.log_message(f"Detalhes: {result.stderr}")
        except Exception as e:
            self.log_message(f"ERRO ao tentar agendar a tarefa: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
