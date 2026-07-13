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
    
    EVOL_MAP = {
        'date.present.record_date': 'Data',
        'kpis_occupancy%.past.after.value': 'Ocupação Passado',
        'kpis_occupancy%.present.after.value': 'Ocupação Presente',
        'kpis_adr%.past.after.value': 'ADR Passado',
        'kpis_adr%.present.after.value': 'ADR Presente',
        'kpis_revenue.past.after.value': 'Receita Passado',
        'kpis_revenue.present.after.value': 'Receita Presente'
    }

    CLOSING_MAP = {
        'kpis_occupancy%.past.after.value': 'Ocupação Passado',
        'kpis_occupancy%.present.after.value': 'Ocupação Presente',
        'kpis_adr%.past.after.value': 'ADR Passado',
        'kpis_adr%.present.after.value': 'ADR Presente',
        'kpis_revenue.past.after.value': 'Receita Passado',
        'kpis_revenue.present.after.value': 'Receita Presente'
    }

    def process_pickup_evolution(self, file_path: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        if not os.path.exists(file_path):
            return None, None
        
        try:
            xl = pd.ExcelFile(file_path)
            valid_sheets = []
            for sheet_name in xl.sheet_names:
                df_tmp = xl.parse(sheet_name)
                if not df_tmp.empty and len(df_tmp.columns) > 0:
                    valid_sheets.append(df_tmp)
            
            if len(valid_sheets) == 0:
                return None, None
            
            # Evolution: busca a primeira planilha que tenha a coluna de Data de Registro
            df_evol_raw = None
            for df in valid_sheets:
                if 'date.present.record_date' in df.columns:
                    df_evol_raw = df
                    break
            
            if df_evol_raw is None:
                df_evol_raw = valid_sheets[0]

            # Closing: busca a última planilha ou uma que tenha os KPIs e não seja a de evolução
            df_closing_raw = valid_sheets[-1]
            for df in valid_sheets:
                if 'kpis_occupancy%.present.after.value' in df.columns and not df.equals(df_evol_raw):
                    df_closing_raw = df
                    break

        except Exception as e:
            print(f"Erro ao ler Excel {file_path}: {e}")
            return None, None
        
        # Evolution: Processar e Agrupar por Data
        df_evol = df_evol_raw[[c for c in self.EVOL_MAP.keys() if c in df_evol_raw.columns]].rename(columns=self.EVOL_MAP)
        for col in self.EVOL_MAP.values():
            if col not in df_evol.columns: df_evol[col] = 0.0
        
        df_evol['Data'] = pd.to_datetime(df_evol['Data'], errors='coerce')
        df_evol = df_evol.dropna(subset=['Data'])
        
        if df_evol.empty:
            return None, None
            
        # Agregação: Soma Receita, Média para Ocupação e ADR
        agg_rules = {col: ('sum' if 'Receita' in col else 'mean') for col in df_evol.columns if col != 'Data'}
        df_evol = df_evol.groupby('Data').agg(agg_rules).reset_index().sort_values('Data')

        # Closing: Processar e Consolidar
        df_closing = df_closing_raw[[c for c in self.CLOSING_MAP.keys() if c in df_closing_raw.columns]].rename(columns=self.CLOSING_MAP)
        for col in self.CLOSING_MAP.values():
            if col not in df_closing.columns: df_closing[col] = 0.0
            
        agg_rules_closing = {col: ('sum' if 'Receita' in col else 'mean') for col in df_closing.columns}
        df_closing_final = df_closing.agg(agg_rules_closing).to_frame().T
        
        return df_evol, df_closing_final
