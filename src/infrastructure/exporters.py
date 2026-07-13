import pandas as pd
import os
import sys
# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.enum.text import PP_ALIGN
# from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import Inches as DocxInches, Pt as DocxPt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

class FileExporter:
    """Classe responsável por exportar dados para diversos formatos (Excel, PPT, Word)."""
    
    def _set_cell_background(self, cell, fill):
        """Define a cor de fundo de uma célula."""
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), fill)
        cell._tc.get_or_add_tcPr().append(shading_elm)

    def _set_cell_border(self, cell):
        """Adiciona bordas pretas sólidas a uma célula e centraliza verticalmente."""
        tcPr = cell._tc.get_or_add_tcPr()
        
        # Bordas
        tcBorders = OxmlElement('w:tcBorders')
        for edge in ['top', 'left', 'bottom', 'right']:
            edge_elm = OxmlElement(f'w:{edge}')
            edge_elm.set(qn('w:val'), 'single')
            edge_elm.set(qn('w:sz'), '4')  # 1/2 pt
            edge_elm.set(qn('w:space'), '0')
            edge_elm.set(qn('w:color'), '000000')
            tcBorders.append(edge_elm)
        tcPr.append(tcBorders)

        # Alinhamento vertical
        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), 'center')
        tcPr.append(vAlign)

    def _get_resource_path(self, relative_path):
        """ Retorna o caminho absoluto para recursos, funcionando em dev e no PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    @staticmethod
    def format_value(val, col_name):
        if pd.api.types.is_datetime64_any_dtype(pd.Series([val])):
            return val.strftime('%d/%m/%Y')
        if 'ADS' in str(col_name):
            return f"{int(val)}"
        if any(k in str(col_name) for k in ['ADR', 'Receita']):
            return f"R$ {val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        if '%' in str(col_name):
            return f"{val:.2f}%"
        if 'Reservas' in str(col_name):
            return f"{int(val)}"
        if isinstance(val, (float, int)):
            return f"{val:.2f}"
        return str(val)

    def export_to_excel(self, dataframes, sheet_names, output_path):
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for df, name in zip(dataframes, sheet_names):
                df.to_excel(writer, sheet_name=name, index=False)

    def _add_page_number(self, run):
        """Helper to add dynamic page numbering field."""
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

    def export_to_docx(self, dataframes, summaries, evol_df, closing_df, charts, output_path, property_name="[NOME DA PROPRIEDADE]", reference=""):
        doc = Document()
        
        # --- Global Style (Inter Light) ---
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Inter Light'
        font.size = DocxPt(11)

        # --- Heading Styles (Inter Light, Size 14, Bold, Centered) ---
        for level in range(1, 3):
            h_style = doc.styles[f'Heading {level}']
            h_style.font.name = 'Inter Light'
            h_style.font.size = DocxPt(14)
            h_style.font.bold = True
            h_style.font.color.rgb = None  # Mantém cor padrão (geralmente preto ou azul escuro)
            h_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # --- Footer Configuration ---
        section = doc.sections[0]
        section.different_first_page_header_footer = True
        
        # Footer for Page 1 (Cover) - Only Copyright
        f_first = section.first_page_footer
        p_first = f_first.paragraphs[0]
        p_first.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_f1 = p_first.add_run('Revinn Estratégias © Copyright – Todos os direitos reservados.')
        run_f1.font.name = 'Inter Light'
        run_f1.font.size = DocxPt(9)

        # Footer for Page 2 onwards - Copyright + Page Number
        f_rest = section.footer
        p_rest = f_rest.paragraphs[0]
        p_rest.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_f2 = p_rest.add_run('Revinn Estratégias © Copyright – Todos os direitos reservados. | Página ')
        run_f2.font.name = 'Inter Light'
        run_f2.font.size = DocxPt(9)
        
        run_page = p_rest.add_run()
        run_page.font.name = 'Inter Light'
        run_page.font.size = DocxPt(9)
        self._add_page_number(run_page)

        # --- Cover ---
        # Logo at top
        logo_path = self._get_resource_path('image.png')
        if os.path.exists(logo_path):
            p_logo = doc.add_paragraph()
            p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_logo = p_logo.add_run()
            run_logo.add_picture(logo_path, width=DocxInches(1.5))

        for _ in range(2): doc.add_paragraph() # Reduzido de 4 para 2

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run('Revinn Estratégias')
        run.font.name = 'Inter Light'
        run.font.size = DocxPt(14)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run('Departamento de Consultoria')
        run.font.name = 'Inter Light'
        run.font.size = DocxPt(11)

        for _ in range(3): doc.add_paragraph() # Reduzido de 7 para 3

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Limpa o nome removendo underscores
        clean_name = property_name.replace('_', ' ').upper()
        run = p.add_run(clean_name)
        run.font.name = 'Inter Light'
        run.font.size = DocxPt(18)
        run.font.bold = True

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle = 'RELATÓRIO DE KPI MENSAL'
        if reference:
            subtitle += f" - {reference}"
        run = p.add_run(subtitle)
        run.font.name = 'Inter Light'
        run.font.size = DocxPt(14)

        # Push São Paulo / 2026 to bottom
        for _ in range(4): doc.add_paragraph() # Reduzido de 8 para 4

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run('São Paulo\n2026')
        run.font.name = 'Inter Light'
        run.font.size = DocxPt(11)

        doc.add_page_break()

        # --- Audit Tables ---
        titles = ["Check-outs por Categoria", "Check-outs por Canal", "Reservas Emitidas por Canal"]
        # Relaciona cada tabela com seu respectivo resumo (sums[0] para CO, sums[1] para RE)
        table_sums = [summaries[0], summaries[0], summaries[1]]
        
        for i, df in enumerate(dataframes):
            summary_data = table_sums[i]
            heading = doc.add_heading(titles[i], level=1)
            
            self._add_docx_table(doc, df, summary=summary_data)
            
            # Adiciona descritivo abaixo da tabela
            revenue_str = self.format_value(summary_data.get('total_receita', 0), 'Receita')
            nights = summary_data.get('total_noites', 0)
            
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(f"RN's: {nights}; Renda Total {revenue_str}")
            run.font.name = 'Inter Light'
            run.font.bold = True
            run.font.size = DocxPt(9)
            
            doc.add_page_break()

        # --- Climber | YoY Mensal (Análise de Pickup) ---
        if evol_df is not None and not evol_df.empty:
            heading = doc.add_heading('Climber | YoY Mensal', level=1)
            
            if len(charts) >= 3:
                heading = doc.add_heading('Evolução Diária (Linhas)', level=2)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(charts[0], width=DocxInches(6.5))
                
                heading = doc.add_heading('Comparativo de Fechamento', level=2)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(charts[1], width=DocxInches(3.1))
                p.add_run().add_picture(charts[2], width=DocxInches(3.1))
                
                doc.add_page_break()
                self._add_closing_docx_table(doc, closing_df)
                doc.add_page_break()
                self._add_evolution_docx_table(doc, evol_df)
            else:
                doc.add_paragraph("Dados de evolução insuficientes para gerar gráficos.")
                if closing_df is not None and not closing_df.empty:
                    self._add_closing_docx_table(doc, closing_df)
        else:
            # Caso não haja dados de pickup, apenas adiciona a página vazia padrão
            doc.add_page_break()
            heading = doc.add_heading('Climber | YoY Mensal', level=1)

        # --- Additional Empty Pages ---
        extra_pages = [
            "Notas & Ranking (Last 90) | Booking.com",
            "Performance (Last 90) | Expedia"
        ]
        for title in extra_pages:
            doc.add_page_break()
            heading = doc.add_heading(title, level=1)

        # --- Back Cover ---
        doc.add_page_break()
        for _ in range(10): doc.add_paragraph()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run('Revinn Estratégias')
        run.font.name = 'Inter Light'
        run.font.size = DocxPt(14)
        run.font.bold = True
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run('Consultoria e Estratégia Hoteleira')
        run.font.name = 'Inter Light'
        run.font.size = DocxPt(11)
        
        doc.save(output_path)

    def set_table_autofit(self, table):
        """Ajusta a tabela para se auto-ajustar ao conteúdo (AutoFit to Contents)."""
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblW = tblPr.xpath('w:tblW')
        if tblW:
            tblW[0].set(qn('w:type'), 'auto')
            tblW[0].set(qn('w:w'), '0')

    def _add_docx_table(self, doc, df, summary=None):
        rows_count = len(df) + 1
        if summary: rows_count += 1
        
        table = doc.add_table(rows=rows_count, cols=len(df.columns))
        self.set_table_autofit(table)
        
        # Header (Cor: #159F92, Texto: Branco)
        for i, col in enumerate(df.columns): 
            cell = table.rows[0].cells[i]
            cell.text = col
            self._set_cell_background(cell, "159F92")
            self._set_cell_border(cell)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Inter Light'
                    run.font.size = DocxPt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)
        
        # Data Rows
        for i, row in enumerate(df.values):
            bg_color = "E8E8E8" if i % 2 == 0 else "FFFFFF"
            for j, val in enumerate(row):
                cell = table.rows[i+1].cells[j]
                cell.text = self.format_value(val, df.columns[j])
                self._set_cell_border(cell)
                
                if j == 0:
                    self._set_cell_background(cell, "159F92")
                else:
                    self._set_cell_background(cell, bg_color)
                
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Inter Light'
                        run.font.size = DocxPt(9)
                        if j == 0:
                            run.font.color.rgb = RGBColor(255, 255, 255)
        
        # Summary Row (TOTAL)
        if summary:
            last_row = table.rows[-1].cells
            last_row[0].text = "TOTAL"
            self._set_cell_background(last_row[0], "159F92")
            
            for j, cell_item in enumerate(last_row):
                self._set_cell_border(cell_item)
                if j > 0:
                    self._set_cell_background(cell_item, "E8E8E8" if len(df) % 2 == 0 else "FFFFFF")
                
                if j == 0: 
                    for p in last_row[0].paragraphs:
                        for run in p.runs: 
                            run.font.name = 'Inter Light'
                            run.font.color.rgb = RGBColor(255, 255, 255)
                    continue
                
                col_name = df.columns[j]
                
                if 'Receita' in col_name or 'ADR' in col_name or 'ADS' in col_name:
                    if 'Receita' in col_name: 
                        val = summary.get('total_receita', 0)
                        last_row[j].text = self.format_value(val, col_name)
                    elif 'ADR' in col_name:
                        rev = summary.get('total_receita', 0)
                        nights = summary.get('total_noites', 1)
                        val = rev / nights if nights > 0 else 0
                        last_row[j].text = self.format_value(val, col_name)
                    elif 'ADS' in col_name:
                        val = df[col_name].mean()
                        last_row[j].text = f"{val:.1f}".replace('.', ',')
                elif 'Reservas' in col_name:
                    val = summary.get('total_reservas', 0)
                    last_row[j].text = str(val)
                elif 'LOS' in col_name:
                    val = summary.get('total_noites', 0) / summary.get('total_reservas', 1)
                    last_row[j].text = f"{val:.2f}"
                elif 'Share' in col_name:
                    last_row[j].text = "100,00%"

            for j, cell in enumerate(last_row):
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Inter Light'
                        run.font.bold = True
                        run.font.size = DocxPt(9)
                        if j == 0:
                            run.font.color.rgb = RGBColor(255, 255, 255)

    def _add_closing_docx_table(self, doc, df):
        heading = doc.add_heading("Tabela de Fechamento Consolidado", level=2)
        
        metrics = [
            ('Ocupação (%)', 'Ocupação Presente', 'Ocupação Passado', '{:.2f}%'),
            ('ADR (R$)', 'ADR Presente', 'ADR Passado', 'R$ {:,.2f}'),
            ('Receita (R$)', 'Receita Presente', 'Receita Passado', 'R$ {:,.2f}')
        ]
        table = doc.add_table(rows=1, cols=4)
        self.set_table_autofit(table)
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = 'Métrica', 'Ano Atual', 'Ano Passado', 'Variação (%)'
        for cell in hdr:
            self._set_cell_background(cell, "159F92")
            self._set_cell_border(cell)
            for p in cell.paragraphs:
                for run in p.runs: 
                    run.font.name = 'Inter Light'
                    run.font.size = DocxPt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)
        
        for i, (label, pres, past, fmt) in enumerate(metrics):
            row = table.add_row().cells
            p_val, pas_val = df[pres].iloc[0], df[past].iloc[0]
            var = ((p_val - pas_val) / pas_val * 100) if pas_val != 0 else 0
            row[0].text = label
            row[1].text = fmt.format(p_val).replace(',', 'X').replace('.', ',').replace('X', '.')
            row[2].text = fmt.format(pas_val).replace(',', 'X').replace('.', ',').replace('X', '.')
            row[3].text = f"{var:+.2f}%"
            
            bg_color = "E8E8E8" if i % 2 == 0 else "FFFFFF"
            for j, cell in enumerate(row):
                self._set_cell_border(cell)
                if j == 0:
                    self._set_cell_background(cell, "159F92")
                else:
                    self._set_cell_background(cell, bg_color)
                
                for p in cell.paragraphs:
                    for run in p.runs: 
                        run.font.name = 'Inter Light'
                        run.font.size = DocxPt(9)
                        if j == 0:
                            run.font.color.rgb = RGBColor(255, 255, 255)
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run("* Nota: A Tabela de Fechamento Consolidado baseia-se na comparação do mesmo período (mês/dias equivalentes) entre o ano atual e o ano anterior (Year-over-Year).")
        run.font.name = 'Inter Light'
        run.font.italic = True
        run.font.size = DocxPt(8)

    def _add_evolution_docx_table(self, doc, df):
        heading = doc.add_heading("Detalhamento Diário YoY", level=2)
        
        cols = ['Data', 'Tipo', 'Ocupação (%)', 'ADR (R$)', 'Receita (R$)']
        table = doc.add_table(rows=1, cols=len(cols))
        self.set_table_autofit(table)
        for i, col in enumerate(cols): 
            cell = table.rows[0].cells[i]
            cell.text = col
            self._set_cell_background(cell, "159F92")
            self._set_cell_border(cell)
            for p in cell.paragraphs:
                for run in p.runs: 
                    run.font.name = 'Inter Light'
                    run.font.size = DocxPt(10)
                    run.font.color.rgb = RGBColor(255, 255, 255)
        
        for i, (_, row) in enumerate(df.iterrows()):
            r1 = table.add_row().cells
            r1[0].text, r1[1].text, r1[2].text = row['Data'].strftime('%d/%m/%Y'), "Ano Atual", f"{row['Ocupação Presente']:.2f}%"
            r1[3].text = self.format_value(row['ADR Presente'], 'ADR')
            r1[4].text = self.format_value(row['Receita Presente'], 'Receita')
            
            r2 = table.add_row().cells
            r2[0].text, r2[1].text, r2[2].text = row['Data'].strftime('%d/%m/%Y'), "Ano Passado", f"{row['Ocupação Passado']:.2f}%"
            r2[3].text = self.format_value(row['ADR Passado'], 'ADR')
            r2[4].text = self.format_value(row['Receita Passado'], 'Receita')

            bg_color_pair = "E8E8E8" if i % 2 == 0 else "FFFFFF"
            
            for row_cells in [r1, r2]:
                for j, cell in enumerate(row_cells):
                    self._set_cell_border(cell)
                    if j == 0:
                        self._set_cell_background(cell, "159F92")
                    else:
                        self._set_cell_background(cell, bg_color_pair)
                    
                    for p in cell.paragraphs:
                        for run in p.runs: 
                            run.font.name = 'Inter Light'
                            run.font.size = DocxPt(9)
                            if j == 0:
                                run.font.color.rgb = RGBColor(255, 255, 255)

    def format_decimal(self, val):
        if pd.isna(val):
            return "0,00"
        try:
            return f"{float(val):.2f}".replace('.', ',')
        except (ValueError, TypeError):
            return str(val)

    def export_to_pdf(self, dataframes, summaries, evol_df, closing_df, output_path, property_name="[NOME DA PROPRIEDADE]", reference="", additional_sections=None):
        # 1. Configurar caminhos do DLL do GTK3 no Windows para evitar falhas do WeasyPrint
        if os.name == 'nt':
            possible_gtk_paths = [
                r"C:\Program Files\GTK3-Runtime Win64\bin",
                r"C:\Program Files (x86)\GTK3-Runtime Win64\bin",
                r"C:\msys64\mingw64\bin",
                r"C:\tools\msys64\mingw64\bin"
            ]
            for path in possible_gtk_paths:
                if os.path.exists(path):
                    try:
                        os.add_dll_directory(path)
                    except AttributeError:
                        os.environ['PATH'] = path + os.path.pathsep + os.environ['PATH']
                    break

        # Importa weasyprint localmente para evitar erros de importação na inicialização da aplicação
        try:
            from weasyprint import HTML
        except OSError as e:
            raise ImportError(
                "Não foi possível inicializar o WeasyPrint. Certifique-se de que o GTK+3 Runtime está instalado no Windows.\n"
                "Para resolver, acesse: https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#windows\n"
                f"Erro detalhado: {str(e)}"
            )

        # 2. Obter valores para os cartões de KPI
        receita_val = "R$ 0,00"
        receita_var = "0,00%"
        receita_var_class = "positive"
        
        ocupacao_val = "0,00%"
        ocupacao_var = "0,00%"
        ocupacao_var_class = "positive"
        
        adr_val = "R$ 0,00"
        adr_var = "0,00%"
        adr_var_class = "positive"
        
        revpar_val = "R$ 0,00"
        revpar_var = "0,00%"
        revpar_var_class = "positive"

        if closing_df is not None and not closing_df.empty:
            # Dados consolidados de fechamento do Pickup
            pres_rec = closing_df['Receita Presente'].iloc[0]
            past_rec = closing_df['Receita Passado'].iloc[0]
            receita_val = self.format_value(pres_rec, 'Receita')
            if past_rec != 0:
                var = (pres_rec - past_rec) / past_rec * 100
                receita_var = f"{var:+.2f}%".replace('.', ',')
                receita_var_class = "positive" if var >= 0 else "negative"
            else:
                receita_var = "0,00%"
                receita_var_class = "positive"

            pres_occ = closing_df['Ocupação Presente'].iloc[0]
            past_occ = closing_df['Ocupação Passado'].iloc[0]
            ocupacao_val = f"{pres_occ:.2f}%".replace('.', ',')
            if past_occ != 0:
                var = (pres_occ - past_occ) / past_occ * 100
                ocupacao_var = f"{var:+.2f}%".replace('.', ',')
                ocupacao_var_class = "positive" if var >= 0 else "negative"
            else:
                ocupacao_var = "0,00%"
                ocupacao_var_class = "positive"

            pres_adr = closing_df['ADR Presente'].iloc[0]
            past_adr = closing_df['ADR Passado'].iloc[0]
            adr_val = self.format_value(pres_adr, 'ADR')
            if past_adr != 0:
                var = (pres_adr - past_adr) / past_adr * 100
                adr_var = f"{var:+.2f}%".replace('.', ',')
                adr_var_class = "positive" if var >= 0 else "negative"
            else:
                adr_var = "0,00%"
                adr_var_class = "positive"

            pres_rev = pres_occ * pres_adr / 100
            past_rev = past_occ * past_adr / 100
            revpar_val = self.format_value(pres_rev, 'ADR')
            if past_rev != 0:
                var = (pres_rev - past_rev) / past_rev * 100
                revpar_var = f"{var:+.2f}%".replace('.', ',')
                revpar_var_class = "positive" if var >= 0 else "negative"
            else:
                revpar_var = "0,00%"
                revpar_var_class = "positive"
        else:
            # Fallback para os dados de check-outs (CO)
            if summaries and len(summaries) > 0:
                co_sum = summaries[0]
                tot_rec = co_sum.get('total_receita', 0)
                tot_nights = co_sum.get('total_noites', 0)
                receita_val = self.format_value(tot_rec, 'Receita')
                receita_var = "N/D"
                receita_var_class = "positive"
                
                ocupacao_val = "N/D"
                ocupacao_var = "N/D"
                ocupacao_var_class = "positive"
                
                calc_adr = tot_rec / tot_nights if tot_nights > 0 else 0
                adr_val = self.format_value(calc_adr, 'ADR')
                adr_var = "N/D"
                adr_var_class = "positive"
                
                revpar_val = "N/D"
                revpar_var = "N/D"
                revpar_var_class = "positive"

        # 3. Construir linhas das tabelas em HTML
        # Tabela 1: Desempenho por Categoria de Acomodação (Check-outs)
        tabela_categorias_rows = []
        df_cat = dataframes[0]
        for _, row in df_cat.iterrows():
            tabela_categorias_rows.append(f"""
            <tr>
                <td>{row['Categoria']}</td>
                <td>{self.format_value(row['ADR'], 'ADR')}</td>
                <td>{self.format_decimal(row['LOS'])}</td>
                <td>{self.format_value(row['Receita'], 'Receita')}</td>
            </tr>
            """)
        if len(summaries) > 0:
            co_sum = summaries[0]
            tot_rec = co_sum.get('total_receita', 0)
            tot_nights = co_sum.get('total_noites', 0)
            tot_res = co_sum.get('total_reservas', 1)
            tot_adr = tot_rec / tot_nights if tot_nights > 0 else 0
            tot_los = tot_nights / tot_res if tot_res > 0 else 0
            tabela_categorias_rows.append(f"""
            <tr class="total-row">
                <td>TOTAL</td>
                <td>{self.format_value(tot_adr, 'ADR')}</td>
                <td>{self.format_decimal(tot_los)}</td>
                <td>{self.format_value(tot_rec, 'Receita')}</td>
            </tr>
            """)
        tabela_categorias_html = "\n".join(tabela_categorias_rows)

        # Tabela 2: Atribuição de Performance por Canais de Emissão (Check-outs)
        tabela_emissoes_passadas_rows = []
        df_canal_co = dataframes[1]
        for _, row in df_canal_co.iterrows():
            tabela_emissoes_passadas_rows.append(f"""
            <tr>
                <td>{row['Canal']}</td>
                <td>{self.format_value(row['Receita'], 'Receita')}</td>
                <td>{int(row['Reservas'])}</td>
                <td>{self.format_decimal(row['LOS'])}</td>
                <td>{self.format_value(row['ADR'], 'ADR')}</td>
                <td>{int(row['ADS'])}</td>
                <td>{self.format_decimal(row['Share (%)'])}%</td>
            </tr>
            """)
        if len(summaries) > 0:
            co_sum = summaries[0]
            tot_rec = co_sum.get('total_receita', 0)
            tot_res = co_sum.get('total_reservas', 0)
            tot_nights = co_sum.get('total_noites', 0)
            tot_los = tot_nights / tot_res if tot_res > 0 else 0
            tot_adr = tot_rec / tot_nights if tot_nights > 0 else 0
            tot_ads = df_canal_co['ADS'].mean() if not df_canal_co.empty else 0
            tabela_emissoes_passadas_rows.append(f"""
            <tr class="total-row">
                <td>TOTAL</td>
                <td>{self.format_value(tot_rec, 'Receita')}</td>
                <td>{tot_res}</td>
                <td>{self.format_decimal(tot_los)}</td>
                <td>{self.format_value(tot_adr, 'ADR')}</td>
                <td>{f"{tot_ads:.1f}".replace('.', ',')}</td>
                <td>100,00%</td>
            </tr>
            """)
        tabela_emissoes_passadas_html = "\n".join(tabela_emissoes_passadas_rows)

        # Tabela 3: Carteira de Captação Futura por Canais de Emissão (Pacing / OTB)
        tabela_emissoes_futuras_rows = []
        df_canal_re = dataframes[2]
        for _, row in df_canal_re.iterrows():
            tabela_emissoes_futuras_rows.append(f"""
            <tr>
                <td>{row['Canal']}</td>
                <td>{self.format_value(row['Receita'], 'Receita')}</td>
                <td>{int(row['Reservas'])}</td>
                <td>{self.format_value(row['ADR'], 'ADR')}</td>
                <td>{int(row['ADS'])}</td>
                <td>{self.format_decimal(row['Share (%)'])}%</td>
            </tr>
            """)
        if len(summaries) > 1:
            re_sum = summaries[1]
            tot_rec = re_sum.get('total_receita', 0)
            tot_res = re_sum.get('total_reservas', 0)
            tot_nights = re_sum.get('total_noites', 0)
            tot_adr = tot_rec / tot_nights if tot_nights > 0 else 0
            tot_ads = df_canal_re['ADS'].mean() if not df_canal_re.empty else 0
            tabela_emissoes_futuras_rows.append(f"""
            <tr class="total-row">
                <td>TOTAL</td>
                <td>{self.format_value(tot_rec, 'Receita')}</td>
                <td>{tot_res}</td>
                <td>{self.format_value(tot_adr, 'ADR')}</td>
                <td>{f"{tot_ads:.1f}".replace('.', ',')}</td>
                <td>100,00%</td>
            </tr>
            """)
        tabela_emissoes_futuras_html = "\n".join(tabela_emissoes_futuras_rows)

        # 3.5 Construir HTML para as seções adicionais dinâmicas
        additional_sections_html = ""
        if additional_sections:
            for sec in additional_sections:
                title = sec.get("title", "").strip()
                desc = sec.get("description", "").strip()
                img_path = sec.get("image_path", "").strip()
                
                additional_sections_html += f"""
                <div style="page-break-before: always;">
                    <div class="table-row header-container">
                        <div class="table-cell" style="width: 50%;">
                            <div class="brand-title">revinn</div>
                            <div class="brand-subtitle">estratégias de receita</div>
                        </div>
                        <div class="table-cell doc-info" style="width: 50%;">
                            <h1>{property_name}</h1>
                            <p>Lâmina de Performance Mensal &bull; {reference}</p>
                        </div>
                    </div>
                """
                
                if title:
                    additional_sections_html += f"""
                    <div class="section-header">{title}</div>
                    """
                
                if desc:
                    formatted_desc = desc.replace("\n", "<br>")
                    additional_sections_html += f"""
                    <div style="font-size: 9.5pt; color: #2c3e50; line-height: 1.5; margin-bottom: 20px; text-align: justify; padding: 10px; background: #f8f9fa; border-left: 3px solid #008080; border: 1px solid #e2e8f0; border-left-width: 3px; border-left-color: #008080; border-radius: 0 4px 4px 0;">
                        {formatted_desc}
                    </div>
                    """
                
                if img_path and os.path.exists(img_path):
                    clean_path = os.path.abspath(img_path).replace("\\", "/")
                    if not clean_path.startswith("/"):
                        clean_path = "/" + clean_path
                    file_url = f"file://{clean_path}"
                    
                    additional_sections_html += f"""
                    <div style="text-align: center; margin-top: 15px; margin-bottom: 15px; page-break-inside: avoid;">
                        <img src="{file_url}" style="max-width: 100%; max-height: 420px; border: 1px solid #e2e8f0; border-radius: 4px; padding: 4px; background: #ffffff;" />
                    </div>
                    """
                
                additional_sections_html += """
                </div>
                """

        # 4. Definir HTML Template e substituir variáveis
        html_template = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Revinn - Lâmina de Performance Mensal</title>
    <style>
        @page {{
            size: A4;
            margin: 15mm 12mm;
            @bottom-right {{
                content: "Pág. " counter(page);
                font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
                font-size: 8pt;
                color: #7f8c8d;
            }}
        }}
        
        *, *::before, *::after {{
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            color: #2c3e50;
            line-height: 1.4;
            margin: 0;
            padding: 0;
            background-color: #ffffff;
            font-size: 10pt;
        }}
        
        /* Estrutura de colunas via display: table para compatibilidade WeasyPrint */
        .table-row {{
            display: table;
            width: 100%;
            table-layout: fixed;
            margin-bottom: 15px;
        }}
        
        .table-cell {{
            display: table-cell;
            vertical-align: top;
        }}
        
        /* Cabeçalho Corporativo */
        .header-container {{
            border-bottom: 2px solid #008080;
            padding-bottom: 8px;
            margin-bottom: 20px;
        }}
        
        .brand-title {{
            font-size: 22pt;
            font-weight: bold;
            color: #008080;
            text-transform: lowercase;
            margin: 0;
            letter-spacing: -0.5px;
        }}
        
        .brand-subtitle {{
            font-size: 8pt;
            font-weight: bold;
            color: #7f8c8d;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            margin-top: -2px;
        }}
        
        .doc-info {{
            text-align: right;
        }}
        
        .doc-info h1 {{
            font-size: 14pt;
            margin: 0;
            color: #2c3e50;
            text-transform: uppercase;
            font-weight: 700;
        }}
        
        .doc-info p {{
            font-size: 9pt;
            margin: 3px 0 0 0;
            color: #7f8c8d;
            font-weight: 600;
        }}

        /* Blocos de Indicadores Gerais (KPIs) */
        .kpi-container {{
            display: table;
            width: 100%;
            table-layout: fixed;
            margin-bottom: 20px;
            border-collapse: separate;
            border-spacing: 8px 0;
            margin-left: -8px;
            margin-right: -8px;
        }}
        
        .kpi-card {{
            display: table-cell;
            background: #f8f9fa;
            border: 1px solid #e2e8f0;
            padding: 10px;
            border-radius: 4px;
            text-align: left;
            vertical-align: middle;
        }}
        
        .kpi-label {{
            font-size: 7.5pt;
            color: #7f8c8d;
            text-transform: uppercase;
            font-weight: 700;
            margin-bottom: 4px;
            display: block;
        }}
        
        .kpi-value {{
            font-size: 14pt;
            font-weight: 700;
            color: #2c3e50;
            margin: 0;
        }}
        
        .kpi-sub {{
            font-size: 7.5pt;
            font-weight: 600;
            margin-top: 3px;
            display: block;
        }}
        
        .negative {{ color: #c0392b; }}
        .positive {{ color: #27ae60; }}

        /* Headers de Seção */
        .section-header {{
            font-size: 10pt;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #ffffff;
            background-color: #2c3e50;
            padding: 5px 8px;
            margin-top: 18px;
            margin-bottom: 10px;
            font-weight: 700;
            border-radius: 2px;
            page-break-inside: avoid;
            page-break-after: avoid;
        }}
        
        /* Tabelas Financeiras */
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 15px;
            font-size: 8.5pt;
            page-break-inside: avoid;
        }}
        
        th {{
            font-weight: 700;
            text-transform: uppercase;
            font-size: 8pt;
            color: #34495e;
            border-bottom: 2px solid #2c3e50;
            padding: 5px 6px;
            text-align: right;
            background-color: #fcfcfc;
        }}
        
        th:first-child, td:first-child {{
            text-align: left;
        }}
        
        td {{
            padding: 5px 6px;
            border-bottom: 1px solid #e2e8f0;
            text-align: right;
            color: #2c3e50;
        }}
        
        .total-row td {{
            font-weight: 700;
            background-color: #f8f9fa;
            border-top: 1px solid #2c3e50;
            border-bottom: 2px solid #2c3e50;
        }}
    </style>
</head>
<body>

<div class="container">

    <div class="table-row header-container">
        <div class="table-cell" style="width: 50%;">
            <div class="brand-title">revinn</div>
            <div class="brand-subtitle">estratégias de receita</div>
        </div>
        <div class="table-cell doc-info" style="width: 50%;">
            <h1>{property_name}</h1>
            <p>Lâmina de Performance Mensal &bull; {reference}</p>
        </div>
    </div>

    <div class="kpi-container">
        <div class="kpi-card">
            <span class="kpi-label">Receita Líquida</span>
            <p class="kpi-value">{receita_val}</p>
            <span class="kpi-sub {receita_var_class}">{receita_var}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Taxa de Ocupação</span>
            <p class="kpi-value">{ocupacao_val}</p>
            <span class="kpi-sub {ocupacao_var_class}">{ocupacao_var}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Diária Média (ADR)</span>
            <p class="kpi-value">{adr_val}</p>
            <span class="kpi-sub {adr_var_class}">{adr_var}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">RevPAR Consolidado</span>
            <p class="kpi-value">{revpar_val}</p>
            <span class="kpi-sub {revpar_var_class}">{revpar_var}</span>
        </div>
    </div>

    <div class="section-header">Desempenho por Categoria de Acomodação (Check-outs)</div>
    <table>
        <thead>
            <tr>
                <th>Categoria de Quarto</th>
                <th>ADR Realizado</th>
                <th>LOS Média (Noites)</th>
                <th>Métricas de Alocação Absoluta</th>
            </tr>
        </thead>
        <tbody>
            {tabela_categorias_html}
        </tbody>
    </table>

    <div class="section-header">Atribuição de Performance por Canais de Emissão (Check-outs)</div>
    <table>
        <thead>
            <tr>
                <th>Canal de Distribuição</th>
                <th>Receita Bruta</th>
                <th>Reservas</th>
                <th>LOS</th>
                <th>ADR</th>
                <th>ADS (Janela)</th>
                <th>Share (%)</th>
            </tr>
        </thead>
        <tbody>
            {tabela_emissoes_passadas_html}
        </tbody>
    </table>

    <div class="section-header">Carteira de Captação Futura por Canais de Emissão (Pacing / OTB)</div>
    <table>
        <thead>
            <tr>
                <th>Canal Temático</th>
                <th>Volume Captado</th>
                <th>Reservas</th>
                <th>ADR Médio</th>
                <th>ADS (Antecedência)</th>
                <th>Share OTB (%)</th>
            </tr>
        </thead>
        <tbody>
            {tabela_emissoes_futuras_html}
        </tbody>
    </table>

    {additional_sections_html}

</div>

</body>
</html>"""

        # 5. Escrever PDF
        temp_html_path = output_path + ".temp.html"
        with open(temp_html_path, "w", encoding="utf-8") as f:
            f.write(html_template)
            
        try:
            HTML(temp_html_path).write_pdf(output_path)
        finally:
            if os.path.exists(temp_html_path):
                os.remove(temp_html_path)
