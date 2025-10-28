import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import warnings
from datetime import datetime
import os
import sys

warnings.filterwarnings('ignore', category=UserWarning)

class PowerBIInterativo:
    def __init__(self, root, source_file=None):
        self.root = root
        self.root.title("Power BI em Python - Relatório Interativo")
        self.root.geometry("1500x900")
        self.source_file = source_file
        
        self.setup_styles()
        try:
            if source_file:
                self.process_csv_data(source_file)
                self.setup_ui()
            else:
                messagebox.showerror("Erro", "Nenhum arquivo CSV informado.")
                root.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            root.destroy()

    def process_csv_data(self, file_path):
        """Ler e preparar dados do CSV (normalização robusta de nomes de coluna)"""
        import unicodedata
        import re

        def normalize_col_name(s):
            if s is None:
                return ''
            s = str(s).strip().lower()
            s = unicodedata.normalize('NFKD', s)
            s = ''.join(ch for ch in s if not unicodedata.combining(ch))
            s = re.sub(r'[^0-9a-z]+', ' ', s)
            s = re.sub(r'\s+', ' ', s).strip()
            return s

        print(f"Carregando dados de: {file_path}")
        try:
            df = pd.read_csv(file_path, encoding='utf-8', sep=';', engine='python')
        except Exception:
            df = pd.read_csv(file_path, encoding='latin1', sep=';', engine='python')

        print("Colunas originais:", list(df.columns))

        normalized_map = {col: normalize_col_name(col) for col in df.columns}
        df.rename(columns=normalized_map, inplace=True)

        print("Colunas normalizadas:", list(df.columns))

        # Detecta automaticamente colunas
        data_col = next((c for c in df.columns if ('data' in c and ('emiss' in c or 'emissao' in c or 'emissão' in c))), None)
        total_col = next((c for c in df.columns if ('total' in c and ('liquido' in c or 'liquid' in c or 'valor' in c or 'total' == c))), None)
        cliente_col = next((c for c in df.columns if ('cliente' in c and 'veiculo' not in c)), None)
        campanha_col = next((c for c in df.columns if 'campanha' in c), None)

        print("Detectado -> data_col:", data_col, " total_col:", total_col, " cliente_col:", cliente_col, " campanha_col:", campanha_col)

        if not data_col or not total_col:
            raise ValueError(
                "Colunas de data e total líquido não foram encontradas.\n"
                "Verifique os nomes das colunas no CSV. Colunas normalizadas: "
                + ", ".join(list(df.columns))
            )

        def parse_valor(v):
            if pd.isna(v):
                return 0.0
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v).strip()
            s = s.replace('r$', '').replace(' ', '').replace('.', '').replace(',', '.')
            try:
                return float(s)
            except:
                m = re.search(r'[-+]?\d+(\.\d+)?', s)
                return float(m.group(0)) if m else 0.0

        df['valor'] = df[total_col].apply(parse_valor)
        df['data'] = pd.to_datetime(df[data_col], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['data'])

        df['ano'] = df['data'].dt.year
        df['mes'] = df['data'].dt.month
        df['anomes'] = df['data'].dt.strftime('%Y-%m')

        df['cliente'] = df[cliente_col] if cliente_col in df.columns else 'Desconhecido'
        df['campanha'] = df[campanha_col] if campanha_col in df.columns else 'Sem Campanha'
        df['cliente'] = df['cliente'].fillna('Desconhecido')
        df['campanha'] = df['campanha'].fillna('Sem Campanha')

        self.df = df
        self.available_years = sorted(df['ano'].unique().tolist())
        self.available_clientes = sorted(df['cliente'].dropna().unique().tolist())
        self.available_campanhas = sorted(df['campanha'].dropna().unique().tolist())

        print(f"Arquivo carregado com {len(df)} linhas.")
        print(f"Anos disponíveis: {self.available_years}")
        print(f"Clientes únicos: {len(self.available_clientes)}")
        print(f"Campanhas únicas: {len(self.available_campanhas)}")

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Header.TLabel', font=('Arial', 11, 'bold'))
        style.configure('Custom.TButton', font=('Arial', 10))

    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(control_frame, text="Ano:", style='Header.TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.ano_var = tk.StringVar(value='Todos')
        self.ano_combo = ttk.Combobox(control_frame, textvariable=self.ano_var, width=10, state="readonly")
        self.ano_combo['values'] = ['Todos'] + [str(a) for a in self.available_years]
        self.ano_combo.pack(side=tk.LEFT, padx=(0, 20))
        self.ano_combo.bind('<<ComboboxSelected>>', lambda e: self.update_plot())

        ttk.Label(control_frame, text="Cliente:", style='Header.TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.cliente_var = tk.StringVar(value='Todos')
        self.cliente_combo = ttk.Combobox(control_frame, textvariable=self.cliente_var, width=40, state="readonly")
        self.cliente_combo['values'] = ['Todos'] + self.available_clientes
        self.cliente_combo.pack(side=tk.LEFT, padx=(0, 20))
        self.cliente_combo.bind('<<ComboboxSelected>>', lambda e: self.update_plot())

        ttk.Label(control_frame, text="Campanha:", style='Header.TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.campanha_var = tk.StringVar(value='Todas')
        self.campanha_combo = ttk.Combobox(control_frame, textvariable=self.campanha_var, width=40, state="readonly")
        self.campanha_combo['values'] = ['Todas'] + self.available_campanhas
        self.campanha_combo.pack(side=tk.LEFT, padx=(0, 20))
        self.campanha_combo.bind('<<ComboboxSelected>>', lambda e: self.update_plot())

        ttk.Button(control_frame, text="Salvar Gráfico", command=self.save_plot, style='Custom.TButton').pack(side=tk.LEFT, padx=(10, 20))

        self.stats_label = ttk.Label(control_frame, text="", style='Header.TLabel')
        self.stats_label.pack(side=tk.LEFT, padx=(20, 0))

        self.plot_frame = ttk.Frame(main_frame)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)
        self.create_plot()

    def create_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(14, 8))
        self.update_plot()

    def filtrar_dados(self):
        df = self.df.copy()
        ano = self.ano_var.get()
        cliente = self.cliente_var.get()
        campanha = self.campanha_var.get()

        if ano != 'Todos':
            df = df[df['ano'] == int(ano)]
        if cliente != 'Todos':
            df = df[df['cliente'] == cliente]
        if campanha != 'Todas':
            df = df[df['campanha'] == campanha]

        return df

    def update_plot(self):
        df_plot = self.filtrar_dados()
        self.ax.clear()

        if df_plot.empty:
            self.ax.text(0.5, 0.5, "Sem dados para exibir", ha='center', va='center', fontsize=14, color='gray')
            self.redraw()
            return

        # 🔍 Filtrar PIs recorrentes
        if 'pi' in df_plot.columns:
            recorrentes = df_plot['pi'].value_counts()
            pis_recor = recorrentes[recorrentes >= 2].index.tolist()
            df_plot = df_plot[df_plot['pi'].isin(pis_recor)]
        else:
            pis_recor = []
            print("Coluna 'pi' não encontrada — exibindo todos os registros.")

        if df_plot.empty:
            self.ax.text(0.5, 0.5, "Nenhuma PI recorrente encontrada", ha='center', va='center', fontsize=14, color='gray')
            self.redraw()
            return

        # 🎨 Cores por cliente
        clientes_unicos = df_plot['cliente'].unique().tolist()
        cmap = plt.get_cmap('tab20')
        cores = {cliente: cmap(i % 20) for i, cliente in enumerate(clientes_unicos)}

        # Guardar info para clique
        self.scatter_data = []
        self.text_labels = []

        # 📊 Plotar linhas e bolinhas por cliente
        for cliente in clientes_unicos:
            df_cliente = df_plot[df_plot['cliente'] == cliente]
            resumo = df_cliente.groupby('anomes')['valor'].sum().reset_index()

            linha, = self.ax.plot(
                resumo['anomes'],
                resumo['valor'],
                marker='o',
                linestyle='-',
                linewidth=2,
                color=cores[cliente],
                alpha=0.8,
                label=cliente
            )

            pontos = self.ax.scatter(
                resumo['anomes'],
                resumo['valor'],
                s=100,
                color=cores[cliente],
                edgecolor='black',
                alpha=0.9,
                picker=True  # 🔹 habilita clique
            )

            self.scatter_data.append((pontos, linha, cliente, resumo))

        # 🧮 Estatísticas gerais
        total = df_plot['valor'].sum()
        media = df_plot['valor'].mean()
        pico = df_plot['valor'].max()
        mes_pico = df_plot.loc[df_plot['valor'].idxmax(), 'anomes']

        # 🏷️ Título dinâmico com contagem de PIs
        self.ax.set_title(
            f"Desempenho de Clientes com PIs Recorrentes (Total de {len(pis_recor)} PIs)",
            fontsize=14,
            fontweight='bold',
            pad=25
        )

        # Limpar texto de cliente ativo
        if hasattr(self, "cliente_texto") and self.cliente_texto is not None:
            try:
                self.cliente_texto.remove()
            except Exception:
                pass
        self.cliente_texto = None

        self.ax.set_xlabel("Mês")
        self.ax.set_ylabel("Valor (R$)")
        self.ax.grid(True, linestyle='--', alpha=0.3)
        plt.xticks(rotation=45)
        self.ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"R$ {x:,.0f}"))

        legend = self.ax.legend(
            title="Cliente",
            bbox_to_anchor=(1.02, 1),
            loc='upper left',
            borderaxespad=0,
            fontsize=9
        )
        legend.get_title().set_fontsize('10')
        self.fig.subplots_adjust(right=0.78)

        self.stats_label.config(
            text=f"Total: R$ {total:,.0f} | Média: R$ {media:,.0f} | Pico: R$ {pico:,.0f} ({mes_pico})"
        )

        self.fig.canvas.mpl_connect('pick_event', self.on_point_click)
        self.redraw()


    # ==========================================================
    # 🔵 Clique nas bolinhas: destacar e mostrar valores
    # ==========================================================
    def on_point_click(self, event):
        artist = event.artist

        # Remover textos antigos (valores e nome de cliente)
        for txt in getattr(self, "text_labels", []):
            txt.remove()
        self.text_labels = []

        if hasattr(self, "cliente_texto") and self.cliente_texto:
            self.cliente_texto.remove()
            self.cliente_texto = None

        for pontos, linha, cliente, resumo in self.scatter_data:
            if artist == pontos:
                # Apagar destaques anteriores
                for p, l, _, _ in self.scatter_data:
                    l.set_linewidth(2)
                    l.set_alpha(0.3)
                    p.set_sizes([100])
                    p.set_alpha(0.3)

                # Destacar o cliente clicado
                linha.set_linewidth(3.5)
                linha.set_alpha(1)
                pontos.set_sizes([200])
                pontos.set_alpha(1)

                # 💰 Mostrar os valores em cada ponto
                for x, y in zip(resumo['anomes'], resumo['valor']):
                    self.text_labels.append(
                        self.ax.text(
                            x, y,
                            f"R$ {y:,.0f}",
                            ha='center',
                            va='bottom',
                            fontsize=8,
                            color='black',
                            fontweight='bold',
                            rotation=45
                        )
                    )

                # 🧾 Mostrar nome do cliente abaixo do título
                self.cliente_texto = self.ax.text(
                    0.5, 1.01,
                    f"🔹 Cliente: {cliente}",
                    transform=self.ax.transAxes,
                    ha='center',
                    va='bottom',
                    fontsize=10,
                    fontweight='bold',
                    color='navy'
                )

                self.redraw()
                break




    def redraw(self):
        if hasattr(self, 'canvas'):
            self.canvas.get_tk_widget().destroy()
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        if hasattr(self, 'toolbar'):
            self.toolbar.destroy()
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.plot_frame)
        self.toolbar.update()

    def save_plot(self):
        try:
            filename = "grafico_powerbi_python.png"
            self.fig.savefig(filename, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Sucesso", f"Gráfico salvo como '{filename}'")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar gráfico: {str(e)}")

def main():
    root = tk.Tk()
    default_path = r"C:\Users\karen.takara\OneDrive - Essie Publicidade e Comunicacao Ltda\Documentos\bi\resultado_filtrado.csv"

    if len(sys.argv) > 1:
        source_file = sys.argv[1]
    else:
        source_file = default_path

    if not os.path.exists(source_file):
        messagebox.showerror("Erro", f"Arquivo CSV não encontrado:\n{source_file}")
        return

    app = PowerBIInterativo(root, source_file=source_file)
    root.mainloop()

if __name__ == "__main__":
    main()
