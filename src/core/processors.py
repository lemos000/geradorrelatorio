import pandas as pd
import os
from typing import Tuple, Dict, Any
import io

class DataProcessor:
    """Classe responsável pelo carregamento e limpeza de dados de auditoria (CSV)."""
    
    REQUIRED_COLS = {
        'textBox17': 'Canal',
        'textBox13': 'Categoria',
        'textBox20': 'CheckIn',
        'textBox19': 'CheckOut',
        'textBox22': 'Receita',
        'textBox3': 'Status',
        'textBox11': 'DataReserva'
    }

    @staticmethod
    def clean_currency(value: Any) -> float:
        if pd.isna(value) or not isinstance(value, str):
            return 0.0
        clean_val = value.replace(' BRL', '').replace('.', '').replace(',', '.')
        try:
            return float(clean_val)
        except ValueError:
            return 0.0

    def process_audit_data(self, file_path: str) -> pd.DataFrame:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

        # Realiza a leitura tolerante a encoding
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='latin1') as f:
                lines = f.readlines()
        except Exception as e:
            raise ValueError(f"Erro ao ler CSV {os.path.basename(file_path)}: {str(e)}")

        # Limpeza e correção estrutural de cada linha
        clean_lines = []
        for i, line in enumerate(lines):
            clean_line = line.strip()
            
            # Remove o delimitador residual no final de cada linha
            if clean_line.endswith(';;;;'):
                clean_line = clean_line[:-4]
                
            # Remove as aspas globais e normaliza as aspas duplas internas para manter valores seguros (ex: "2.172,74")
            if i > 0 and clean_line.startswith('"') and clean_line.endswith('"'):
                clean_line = clean_line[1:-1].replace('""', '"')
                
            clean_lines.append(clean_line)

        # Transforma o buffer limpo em DataFrame
        try:
            df = pd.read_csv(io.StringIO('\n'.join(clean_lines)))
        except Exception as e:
            raise ValueError(f"Erro no parsing estrutural do CSV {os.path.basename(file_path)}: {str(e)}")

        missing = [col for col in self.REQUIRED_COLS.keys() if col not in df.columns]
        if missing:
            raise KeyError(f"Colunas ausentes em {os.path.basename(file_path)}: {', '.join(missing)}")

        df = df[list(self.REQUIRED_COLS.keys())].rename(columns=self.REQUIRED_COLS)
        df = df.dropna(subset=['Canal', 'Receita'])
        df = df[~df['Canal'].astype(str).str.contains('Total', case=False, na=False)]
        
        if df.empty:
            raise ValueError(f"O arquivo {os.path.basename(file_path)} está vazio após a limpeza.")

        df['Receita'] = df['Receita'].apply(self.clean_currency)
        df['CheckIn'] = pd.to_datetime(df['CheckIn'], format='%d/%m/%Y', errors='coerce')
        df['CheckOut'] = pd.to_datetime(df['CheckOut'], format='%d/%m/%Y', errors='coerce')
        df = df.dropna(subset=['CheckIn', 'CheckOut'])
        
        if df.empty:
            raise ValueError(f"Nenhuma data válida em {os.path.basename(file_path)}.")

        df['LOS_Calc'] = (df['CheckOut'] - df['CheckIn']).dt.days
        df['LOS_Calc'] = df['LOS_Calc'].apply(lambda x: x if x > 0 else 1)
        
        df['DataReserva'] = pd.to_datetime(df['DataReserva'], format='%d/%m/%Y', errors='coerce')
        df['LeadTime'] = (df['CheckIn'] - df['DataReserva']).dt.days.fillna(0)
        df['LeadTime'] = df['LeadTime'].apply(lambda x: x if x >= 0 else 0)

        df['IsCancelled'] = df['Status'].astype(str).str.contains('cancelada', case=False, na=False).astype(int)
        
        return df

    def get_summary_metrics(self, df: pd.DataFrame) -> Dict[str, Any]:
        return {
            'total_reservas': len(df),
            'total_receita': df['Receita'].sum(),
            'total_noites': df['LOS_Calc'].sum(),
            'total_canceladas': df['IsCancelled'].sum()
        }

class PickupProcessor:
    """Classe responsável pelo processamento de dados de Pickup (Excel)."""
    
    def process_pickup_evolution(self, file_path: str, selected_month: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        if not os.path.exists(file_path):
            return None, None
        
        month_map = {
            "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
            "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
        }
        month_num = month_map.get(selected_month, 5)
        
        try:
            xl = pd.ExcelFile(file_path)
            
            # 1. Encontrar planilha de evolução
            evol_sheet = None
            for name in xl.sheet_names:
                if 'evolution_table' in name:
                    evol_sheet = name
                    break
            if not evol_sheet:
                evol_sheet = xl.sheet_names[0]
                
            df_evol_raw = xl.parse(evol_sheet)
            
            # 2. Encontrar planilha de pickup (fechamento)
            pickup_sheet = None
            for name in xl.sheet_names:
                if 'pickup_table' in name and not name.endswith('_piv'):
                    pickup_sheet = name
                    break
            if not pickup_sheet:
                pickup_sheet = xl.sheet_names[-1]
                
            df_closing_raw = xl.parse(pickup_sheet)
            
        except Exception as e:
            print(f"Erro ao ler Excel {file_path}: {e}")
            return None, None
            
        # --- PROCESSAR EVOLUÇÃO ---
        # Filtrar segment == vazio (NaN ou string em branco)
        if 'segment' in df_evol_raw.columns:
            df_evol_raw = df_evol_raw[df_evol_raw['segment'].isna() | (df_evol_raw['segment'].astype(str).str.strip() == '')]
            
        # Colunas mapeadas
        evol_cols = {
            'date.present.record_date': 'Data',
            'kpis_occupancy%.past.after.value': 'Ocupação Passado',
            'kpis_occupancy%.present.after.value': 'Ocupação Presente',
            'kpis_adr%.past.after.value': 'ADR Passado',
            'kpis_adr%.present.after.value': 'ADR Presente',
            'kpis_revpar%.past.after.value': 'RevPAR Passado',
            'kpis_revpar%.present.after.value': 'RevPAR Presente'
        }
        
        # Filtra as colunas existentes no raw
        existing_evol_cols = {raw_col: new_col for raw_col, new_col in evol_cols.items() if raw_col in df_evol_raw.columns}
        df_evol = df_evol_raw[list(existing_evol_cols.keys())].rename(columns=existing_evol_cols)
        
        # Garantir colunas
        for col in evol_cols.values():
            if col not in df_evol.columns:
                df_evol[col] = 0.0
                
        df_evol['Data'] = pd.to_datetime(df_evol['Data'], errors='coerce')
        df_evol = df_evol.dropna(subset=['Data'])
        
        # Filtrar evolução pelo mês selecionado
        df_evol = df_evol[df_evol['Data'].dt.month == month_num]
        
        if not df_evol.empty:
            agg_rules = {col: 'mean' for col in df_evol.columns if col != 'Data'}
            df_evol = df_evol.groupby('Data').agg(agg_rules).reset_index().sort_values('Data')
        else:
            df_evol = pd.DataFrame(columns=evol_cols.values())
        
        # --- PROCESSAR FECHAMENTO ---
        if 'segment' in df_closing_raw.columns:
            df_closing_raw = df_closing_raw[df_closing_raw['segment'].isna() | (df_closing_raw['segment'].astype(str).str.strip() == '')]
            
        closing_cols = {
            'date.present.calendar_date': 'DataFechamento',
            'kpis_occupancy%.past.after.value': 'Ocupação Passado',
            'kpis_occupancy%.present.after.value': 'Ocupação Presente',
            'kpis_adr%.past.after.value': 'ADR Passado',
            'kpis_adr%.present.after.value': 'ADR Presente',
            'kpis_revenue.past.after.value': 'Receita Passado',
            'kpis_revenue.present.after.value': 'Receita Presente',
            'kpis_revpar%.past.after.value': 'RevPAR Passado',
            'kpis_revpar%.present.after.value': 'RevPAR Presente'
        }
        
        existing_closing_cols = {raw_col: new_col for raw_col, new_col in closing_cols.items() if raw_col in df_closing_raw.columns}
        df_closing = df_closing_raw[list(existing_closing_cols.keys())].rename(columns=existing_closing_cols)
        
        # Garantir colunas
        for col in closing_cols.values():
            if col not in df_closing.columns:
                df_closing[col] = 0.0
                
        df_closing['DataFechamento'] = pd.to_datetime(df_closing['DataFechamento'], errors='coerce')
        
        # Filtrar pelo mês selecionado
        df_closing_month = df_closing[df_closing['DataFechamento'].dt.month == month_num]
        
        if df_closing_month.empty:
            # Fallback se não achar exatamente o mês: faz média/soma de todo o período do closing
            agg_rules_closing = {col: ('sum' if 'Receita' in col else 'mean') for col in df_closing.columns if col != 'DataFechamento'}
            df_closing_final = df_closing.drop(columns=['DataFechamento']).agg(agg_rules_closing).to_frame().T
        else:
            df_closing_final = df_closing_month.iloc[[0]].copy()
            if 'DataFechamento' in df_closing_final.columns:
                df_closing_final = df_closing_final.drop(columns=['DataFechamento'])
                
        return df_evol, df_closing_final
