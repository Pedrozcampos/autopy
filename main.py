import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

# Configurações de Design
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AuditProcessor:
    def __init__(self, file_path, et_value=100000):
        self.file_path = file_path
        self.et_value = et_value
        self.keywords = ['ajuste', 'estorno', 'erro', 'manual', 'urgente', 'socio', 'conforme']

    def process_audit(self, output_path):
        # Carregamento inteligente (CSV ou Excel)
        if self.file_path.endswith('.csv'):
            df = pd.read_csv(self.file_path)
        else:
            df = pd.read_excel(self.file_path, engine='openpyxl')
        

        # Nota: O pandas as vezes lê a coluna sem nome como 'Unnamed: 2'
        df.columns = [c if 'Unnamed' not in str(c) else 'Historico' for c in df.columns]
        
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
        df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
        df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)
        df['Valor_Bruto'] = df['Débito'] + df['Crédito']

        # --- Procedimentos ---
        df['Proc_10x_Media'] = df['Valor_Bruto'] > (df.groupby('Cta.C.Part.')['Valor_Bruto'].transform('mean') * 10)
        df['Proc_Excede_ET'] = df['Valor_Bruto'] > self.et_value
        df['Proc_Redondo'] = df['Valor_Bruto'].apply(lambda x: x > 0 and x % 100 == 0)
        df['Proc_Sem_Hist'] = df['Historico'].apply(lambda x: len(str(x)) < 5 or pd.isna(x))
        df['Proc_FDS'] = df['Data'].dt.dayofweek.isin([5, 6])
        df['Proc_Keywords'] = df['Historico'].str.contains('|'.join(self.keywords), case=False, na=False)

        output_path = "Razao_Auditado_Final.xlsx"
        df.to_excel(output_path, index=False)
        
        # Gerar estatísticas para o Painel
        stats = {
            '10x Média': df['Proc_10x_Media'].sum(),
            'Excede ET': df['Proc_Excede_ET'].sum(),
            'Vlr Redondo': df['Proc_Redondo'].sum(),
            'Sem Histórico': df['Proc_Sem_Hist'].sum(),
            'Fim de Semana': df['Proc_FDS'].sum(),
            'Palavras-Chave': df['Proc_Keywords'].sum()
        }
        
        return output_path, stats

class DashboardWindow(ctk.CTkToplevel):
    def __init__(self, stats):
        super().__init__()
        self.title("Painel de Análise de Riscos")
        self.geometry("700x500")
        
        label = ctk.CTkLabel(self, text="Resumo de Ocorrências por Procedimento", font=("Roboto", 18, "bold"))
        label.pack(pady=10)

        # Criando o gráfico com Matplotlib
        fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
        fig.patch.set_facecolor('#2b2b2b') # Cor de fundo combinando com Dark Mode
        ax.set_facecolor('#2b2b2b')
        
        names = list(stats.keys())
        values = list(stats.values())
        
        bars = ax.bar(names, values, color='#1f538d')
        ax.set_xticklabels(names, rotation=45, ha='right', color='white', fontsize=10)
        plt.subplots_adjust(bottom=0.25)
        ax.tick_params(axis='y', colors='white')
        ax.spines['bottom'].set_color('white')
        ax.spines['left'].set_color('white')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        plt.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)

        # Adiciona os números em cima das barras
        for bar in bars:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', color='white')

        canvas = FigureCanvasTkAgg(fig, master=self)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações de Janela
        self.title("Auditor Contábil")
        self.geometry("500x550")
        
        # 1. Cor de Fundo Profunda
        self.configure(fg_color="#121212") 

        # 2. Container Principal
        self.main_container = ctk.CTkFrame(
            self, 
            fg_color="#1e1e1e",      
            corner_radius=25,        
            border_width=1,           
            border_color="#333333"    
        )
        self.main_container.pack(pady=40, padx=40, fill="both", expand=True)

        # Título
        self.label = ctk.CTkLabel(
            self.main_container, 
            text="Auditoria Contábil", 
            font=("Roboto", 27, "bold"),
            text_color="#ffffff"
        )
        self.label.pack(pady=(35, 20))

        # Campo de Entrada ET
        self.et_label = ctk.CTkLabel(self.main_container, text="Valor do Erro Tolerável (ET):", font=("Roboto", 15))
        self.et_label.pack(pady=(10, 5))
        
        self.et_entry = ctk.CTkEntry(
            self.main_container, 
            width=220, 
            height=40,
            border_color="#444444",
            fg_color="#2a2a2a",
            placeholder_text="Ex: 100",
            justify="center"
        )
        self.et_entry.insert(0, "100")
        self.et_entry.pack(pady=10)

        # 3. Botão Principal 
        self.btn_run = ctk.CTkButton(
            self.main_container, 
            text="Selecionar Razão (.xlsx)", 
            command=self.run_process, 
            height=50,
            width=220,
            corner_radius=12,
            fg_color="#3b82f6",       
            hover_color="#2563eb",    
            font=("Roboto", 14, "bold"),
            border_width=2,
            border_color="#60a5fa"    
        )
        self.btn_run.pack(pady=25)

        # 4. Botão de Painel (Invisível)
        self.btn_dash = ctk.CTkButton(
            self.main_container, 
            text="Ver Painel de Análise", 
            command=self.open_dashboard, 
            fg_color="transparent", 
            border_width=2, 
            border_color="#3b82f6",
            text_color="#3b82f6",
            hover_color="#1e293b",
            height=40
        )
        
    def run_process(self):
        path_entrada = filedialog.askopenfilename(
            title="Selecione o arquivo para auditoria",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )

        if path_entrada:
            path_saida = filedialog.asksaveasfilename(
                title="Salvar arquivo auditado como...",
                defaultextension=".xlsx",
                initialfile="Razao_Auditado_Final.xlsx",
                filetypes=[("Excel files", "*.xlsx")]   
            )
 
            if path_saida:
                try:
                    et = float(self.et_entry.get())
                    proc = AuditProcessor(path_entrada, et)
                    _, self.stats = proc.process_audit(path_saida)
                    
                    
                    mensagem = f"Sucesso! Arquivo processado e salvo em:\n{path_saida}"
                    messagebox.showinfo("Sucesso", mensagem)
                    self.btn_dash.pack(pady=10) # Revela o botão do painel
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro: {str(e)}")

    def open_dashboard(self):
        if self.stats:
            DashboardWindow(self.stats)

if __name__ == "__main__":
    app = App()
    app.mainloop()