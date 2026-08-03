import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
from src.core.processors import DataProcessor, PickupProcessor
from src.visualization.visualizers import ChartGenerator, TableGenerator
from src.infrastructure.exporters import FileExporter

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Revinn Estratégias - Gerador de Relatórios")
        self.geometry("700x650")
        
        # Tema e Estilo
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Injeção de dependências
        self.data_proc = DataProcessor()
        self.pickup_proc = PickupProcessor()
        self.exporter = FileExporter()
        
        self._setup_ui()

    def _setup_ui(self):
        # Grid Configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Header - Heurística #8: Design Estético e Minimalista
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(pady=(20, 10), padx=30, fill="x")
        
        ctk.CTkLabel(header_frame, text="Auditoria Hoteleira BI", font=("Poppins", 28, "bold")).pack(side="left")
        
        # Mode Toggle - Heurística #7: Flexibilidade e Eficiência
        self.mode_switch = ctk.CTkSwitch(header_frame, text="Modo Escuro", command=self.toggle_appearance, font=("Inter", 12))
        self.mode_switch.select()
        self.mode_switch.pack(side="right")

        # Main Container
        container = ctk.CTkScrollableFrame(self)
        container.pack(pady=10, padx=30, fill="both", expand=True)

        # Seção de Arquivos - Heurística #6: Reconhecimento em vez de recordação
        self._create_section_label(container, "Fontes de Dados (CSV/XLSX)")
        
        self.file1_entry = self._create_file_field(container, "Check-outs *", self.select_file1, "Selecione o arquivo de check-outs")
        self.file2_entry = self._create_file_field(container, "Emitidas *", self.select_file2, "Selecione o arquivo de reservas emitidas")
        self.pickup_entry = self._create_file_field(container, "Pickup (Opcional)", self.select_pickup_file, "Selecione o arquivo de pickup para gráficos")
        
        # Seção de Configurações - Heurística #2: Correspondência entre o sistema e o mundo real
        self._create_section_label(container, "Informações do Relatório")
        
        self.hotel_entry = ctk.CTkEntry(container, placeholder_text="Nome da Propriedade (ex: Hotel Central)", font=("Inter", 12))
        self.hotel_entry.pack(pady=5, padx=10, fill="x")

        self.ref_entry = ctk.CTkEntry(container, placeholder_text="Referência Temporal (ex: Maio 2026)", font=("Inter", 12))
        self.ref_entry.pack(pady=5, padx=10, fill="x")

        self.month_var = ctk.StringVar(value="Maio")
        self.month_select = ctk.CTkComboBox(container, values=[
            "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
        ], variable=self.month_var, font=("Inter", 12))
        self.month_select.pack(pady=5, padx=10, fill="x")

        # Seções Estáticas Booking e Expedia
        self.create_static_sections(container)

        # Destino - Heurística #5: Prevenção de erros
        self._create_section_label(container, "Destino e Saída")
        self.out_entry = self._create_file_field(container, "Pasta Destino *", self.select_output_folder, "Onde os relatórios serão salvos")
        
        self.name_entry = ctk.CTkEntry(container, placeholder_text="Nome do Arquivo Final *", font=("Inter", 12))
        self.name_entry.insert(0, "Relatorio_Auditoria")
        self.name_entry.pack(pady=5, padx=10, fill="x")

        # Barra de Status/Progresso - Heurística #1: Visibilidade do status do sistema
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(pady=(10, 5), padx=30, fill="x")
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(self, text="Pronto para iniciar", font=("Inter", 12))
        self.status_label.pack(pady=(0, 10))

        # Ações
        self.process_btn = ctk.CTkButton(self, text="GERAR RELATÓRIOS", font=("Inter", 16, "bold"), 
                                        command=self.start_processing, height=50, fg_color="#159F92", hover_color="#005151")
        self.process_btn.pack(pady=(0, 20), padx=30, fill="x")

        # Log Expansível - Heurística #9: Ajuda os usuários a reconhecerem e recuperarem-se de erros
        self.log_text = ctk.CTkTextbox(self, height=100, font=("Consolas", 12))
        self.log_text.pack(pady=(0, 20), padx=30, fill="x")

    def _create_section_label(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=("Inter", 14, "bold"), text_color="#159F92").pack(pady=(15, 5), padx=10, anchor="w")

    def _create_file_field(self, parent, label_text, command, tooltip):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(pady=2, padx=5, fill="x")
        
        lbl = ctk.CTkLabel(frame, text=label_text, width=120, anchor="w", font=("Inter", 12))
        lbl.pack(side="left", padx=5)
        
        entry = ctk.CTkEntry(frame, placeholder_text=tooltip, font=("Inter", 12))
        entry.pack(side="left", padx=5, expand=True, fill="x")
        
        ctk.CTkButton(frame, text="Procurar", width=80, command=command, font=("Inter", 12), fg_color="#159F92", hover_color="#005151").pack(side="left", padx=2)
        return entry

    def toggle_appearance(self):
        if self.mode_switch.get():
            ctk.set_appearance_mode("dark")
        else:
            ctk.set_appearance_mode("light")

    def log(self, msg, is_error=False):
        color = "red" if is_error else "gray"
        self.log_text.insert("end", f"> {msg}\n")
        self.log_text.see("end")
        self.status_label.configure(text=msg, text_color="white" if not is_error else "#ff4d4d")

    def select_file1(self): self._fill_entry(self.file1_entry, filedialog.askopenfilename(filetypes=[("CSV", "*.csv")]))
    def select_file2(self): self._fill_entry(self.file2_entry, filedialog.askopenfilename(filetypes=[("CSV", "*.csv")]))
    def select_pickup_file(self): self._fill_entry(self.pickup_entry, filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]))
    def select_output_folder(self): self._fill_entry(self.out_entry, filedialog.askdirectory())
    
    def _fill_entry(self, entry, val):
        if val: entry.delete(0, "end"); entry.insert(0, val)

    def start_processing(self):
        # Coleta os dados das seções adicionais
        additional_sec_data = []
        for sec in self.additional_sections:
            title = sec["title_entry"].get().strip()
            desc = sec["desc_textbox"].get("1.0", "end-1c").strip()
            img_path = sec["image_entry"].get().strip()
            if title or desc or img_path:
                additional_sec_data.append({
                    "key": sec.get("key", ""),
                    "title": title,
                    "description": desc,
                    "image_path": img_path
                })

        args = (
            self.file1_entry.get().strip(), 
            self.file2_entry.get().strip(), 
            self.pickup_entry.get().strip(), 
            self.out_entry.get().strip(), 
            self.name_entry.get().strip(),
            self.hotel_entry.get().strip(),
            self.ref_entry.get().strip(),
            self.month_var.get().strip()
        )
        
        # Heurística #5 e #9: Prevenção e Recuperação de erros (Aviso específico)
        missing = []
        if not args[0]: missing.append("Check-outs *")
        if not args[1]: missing.append("Emitidas *")
        if not args[3]: missing.append("Pasta Destino *")
        if not args[4]: missing.append("Nome do Arquivo Final *")

        if missing:
            msg = "Por favor, preencha os seguintes campos obrigatórios:\n\n" + "\n".join(f"- {f}" for f in missing)
            messagebox.showwarning("Campos Obrigatórios", msg)
            return
            
        self.process_btn.configure(state="disabled", text="PROCESSANDO...")
        self.progress_bar.set(0)
        self.progress_bar.configure(determinate_speed=0.5)
        self.progress_bar.start()
        
        threading.Thread(target=self.process, args=args + (additional_sec_data,), daemon=True).start()

    def process(self, f1, f2, p_file, out, name, hotel, reference, selected_month, additional_sections):
        try:
            self.log("Carregando bases de dados...")
            df_co = self.data_proc.process_audit_data(f1)
            self.progress_bar.set(0.2)
            
            df_re = self.data_proc.process_audit_data(f2)
            self.progress_bar.set(0.4)
            
            self.log("Gerando tabelas comparativas...")
            rel1, rel2, rel3 = TableGenerator.generate_audit_reports(df_co, df_re)
            sums = [self.data_proc.get_summary_metrics(df_co), self.data_proc.get_summary_metrics(df_re)]
            self.progress_bar.set(0.6)

            evol_df, closing_df, charts = None, None, []
            if p_file and os.path.exists(p_file):
                self.log("Analisando Pickup e buscando planilhas válidas...")
                evol_df, closing_df = self.pickup_proc.process_pickup_evolution(p_file, selected_month)
                
                if evol_df is not None and not evol_df.empty:
                    self.log(f"Pickup: {len(evol_df)} dias de evolução encontrados.")
                    self.log("Gerando gráficos de tendência...")
                    charts = ChartGenerator.generate_pickup_charts(evol_df, closing_df)
                else:
                    self.log("Aviso: Nenhuma aba de evolução válida encontrada no Excel de Pickup.", is_error=True)
            
            self.progress_bar.set(0.8)
            base_path = os.path.join(out, name)
            
            self.log("Finalizando exportação dos documentos...")
            self.exporter.export_to_excel([rel1, rel2, rel3], ['Cat_CO', 'Canal_CO', 'Canal_RE'], base_path + ".xlsx")
            
            # Exporta Word
            self.exporter.export_to_docx(
                [rel1, rel2, rel3], sums, evol_df, closing_df, charts, base_path + ".docx", 
                property_name=hotel if hotel else name,
                reference=reference
            )
            
            # Exporta PDF Lâmina de Performance
            try:
                self.log("Gerando Lâmina de Performance PDF...")
                self.exporter.export_to_pdf(
                    [rel1, rel2, rel3], sums, evol_df, closing_df, base_path + ".pdf",
                    property_name=hotel if hotel else name,
                    reference=reference,
                    additional_sections=additional_sections
                )
            except Exception as pdf_err:
                self.log(f"Não foi possível gerar o PDF: {str(pdf_err)}", is_error=True)
            
            self.progress_bar.stop()
            self.progress_bar.set(1.0)
            self.log("Relatórios gerados com sucesso!")
            
            if messagebox.askyesno("Sucesso", "Relatórios gerados! Deseja abrir a pasta de destino?"):
                if os.path.exists(out):
                    os.startfile(out)
                    
        except Exception as e:
            self.log(f"Erro no processamento: {str(e)}", is_error=True)
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{str(e)}")
            self.progress_bar.stop()
            self.progress_bar.set(0)
        finally:
            self.process_btn.configure(state="normal", text="GERAR RELATÓRIOS")

    def create_static_sections(self, container):
        self._create_section_label(container, "Canais: Análises de Performance")
        self.sections_container = ctk.CTkFrame(container, fg_color="transparent")
        self.sections_container.pack(pady=5, padx=10, fill="x")
        self.additional_sections = []
        
        # 1. Seção Booking
        self._add_static_section("Booking", "Notas & Ranking (Last 90) | Booking.com")
        
        # 2. Seção Expedia
        self._add_static_section("Expedia", "Performance (Last 90) | Expedia")

    def _add_static_section(self, label_text, default_title):
        card = ctk.CTkFrame(self.sections_container, border_width=1, border_color="#159F92")
        card.pack(pady=10, padx=5, fill="x")
        
        header_frame = ctk.CTkFrame(card, fg_color="transparent")
        header_frame.pack(pady=5, padx=10, fill="x")
        
        lbl = ctk.CTkLabel(header_frame, text=f"Seção: {label_text}", font=("Inter", 12, "bold"))
        lbl.pack(side="left")
        
        title_entry = ctk.CTkEntry(card, placeholder_text="Título da Seção", font=("Inter", 12))
        title_entry.insert(0, default_title)
        title_entry.pack(pady=5, padx=10, fill="x")
        
        lbl_desc = ctk.CTkLabel(card, text="Descrição / Conteúdo:", font=("Inter", 10))
        lbl_desc.pack(anchor="w", padx=10)
        
        desc_textbox = ctk.CTkTextbox(card, height=70, font=("Inter", 12))
        desc_textbox.pack(pady=(0, 5), padx=10, fill="x")
        
        img_frame = ctk.CTkFrame(card, fg_color="transparent")
        img_frame.pack(pady=5, padx=10, fill="x")
        
        img_entry = ctk.CTkEntry(img_frame, placeholder_text="Selecione imagens para anexar (opcional, separadas por ';')", font=("Inter", 12))
        img_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        def seek_image(ent=img_entry):
            fpaths = filedialog.askopenfilenames(filetypes=[("Imagens", "*.png *.jpg *.jpeg *.gif *.bmp")])
            if fpaths:
                ent.delete(0, "end")
                ent.insert(0, "; ".join(fpaths))
                
        seek_btn = ctk.CTkButton(img_frame, text="Procurar", width=80, command=seek_image, font=("Inter", 12), fg_color="#159F92", hover_color="#005151")
        seek_btn.pack(side="right")
        
        self.additional_sections.append({
            "key": label_text.lower(),
            "frame": card,
            "title_entry": title_entry,
            "desc_textbox": desc_textbox,
            "image_entry": img_entry
        })
