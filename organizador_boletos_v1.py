import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import re
import fitz  # PyMuPDF
import pandas as pd
from typing import Dict, Optional, List, Tuple
import queue
from datetime import datetime, timedelta

# --- BIBLIOTECAS PARA QR CODE ---
try:
    from PIL import Image
    from pyzbar.pyzbar import decode
    import io
    QR_CODE_DISPONIVEL = True
except ImportError:
    QR_CODE_DISPONIVEL = False

class ExtratorBoletosGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🧾 Extrator de Dados de Boletos - v1")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        self.root.configure(bg='#f0f0f0')
        
        # Variáveis
        self.pasta_selecionada = tk.StringVar()
        self.progresso = tk.DoubleVar()
        self.status_atual = tk.StringVar(value="Pronto para processar boletos")
        self.queue = queue.Queue()
        
        self.criar_interface()
        self.verificar_dependencias()
        
    def verificar_dependencias(self):
        """Verifica se as bibliotecas necessárias estão instaladas"""
        if not QR_CODE_DISPONIVEL:
            messagebox.showwarning(
                "Dependências", 
                "Algumas bibliotecas para QR Code não estão instaladas.\n"
                "Para funcionalidade completa, execute no terminal:\n"
                "pip install pyzbar Pillow"
            )
    
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="🧾 Extrator de Dados de Boletos v1", 
                           font=('Arial', 16, 'bold'))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Frame de seleção de pasta
        pasta_frame = ttk.LabelFrame(main_frame, text="Seleção de Pasta", padding="10")
        pasta_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        pasta_frame.columnconfigure(1, weight=1)
        
        ttk.Label(pasta_frame, text="Pasta:").grid(row=0, column=0, padx=(0, 10))
        self.entry_pasta = ttk.Entry(pasta_frame, textvariable=self.pasta_selecionada, state='readonly')
        self.entry_pasta.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.btn_selecionar = ttk.Button(pasta_frame, text="📁 Selecionar", 
                                         command=self.selecionar_pasta)
        self.btn_selecionar.grid(row=0, column=2)
        
        # Frame de opções avançadas
        opcoes_frame = ttk.LabelFrame(main_frame, text="Opções de Processamento", padding="10")
        opcoes_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        opcoes_frame.columnconfigure(0, weight=1)
        
        self.var_multiplas_paginas = tk.BooleanVar(value=True)
        ttk.Checkbutton(opcoes_frame, text="🔍 Análise inteligente em múltiplas páginas", 
                        variable=self.var_multiplas_paginas).grid(row=0, column=0, sticky=tk.W)
        
        self.var_backup_dados = tk.BooleanVar(value=True)
        ttk.Checkbutton(opcoes_frame, text="💾 Salvar dados brutos para debug", 
                        variable=self.var_backup_dados).grid(row=1, column=0, sticky=tk.W)
        
        # Frame de controles
        controles_frame = ttk.Frame(main_frame)
        controles_frame.grid(row=3, column=0, columnspan=3, pady=(0, 15))
        
        self.btn_processar = ttk.Button(controles_frame, text="🚀 Processar Boletos", 
                                        command=self.iniciar_processamento, style='Accent.TButton')
        self.btn_processar.grid(row=0, column=0, padx=(0, 10))
        self.btn_limpar = ttk.Button(controles_frame, text="🗑️ Limpar", 
                                     command=self.limpar_resultados)
        self.btn_limpar.grid(row=0, column=1, padx=(0, 10))
        self.btn_abrir_pasta = ttk.Button(controles_frame, text="📂 Abrir Pasta", 
                                          command=self.abrir_pasta_resultado)
        self.btn_abrir_pasta.grid(row=0, column=2, padx=(0, 10))
        self.btn_exportar = ttk.Button(controles_frame, text="📊 Exportar Relatório", 
                                       command=self.exportar_relatorio_detalhado)
        self.btn_exportar.grid(row=0, column=3)
        
        # Frame de progresso
        progresso_frame = ttk.LabelFrame(main_frame, text="Progresso", padding="10")
        progresso_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        progresso_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progresso_frame, variable=self.progresso, 
                                            maximum=100, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self.label_status = ttk.Label(progresso_frame, textvariable=self.status_atual)
        self.label_status.grid(row=1, column=0, sticky=tk.W)
        
        # Frame de resultados
        resultados_frame = ttk.LabelFrame(main_frame, text="Resultados", padding="10")
        resultados_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        resultados_frame.columnconfigure(0, weight=1)
        resultados_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Treeview para mostrar os resultados
        colunas = ('Arquivo', 'Linha Digitável', 'Valor', 'Vencimento', 'QR Code', 'Status', 'Páginas')
        self.tree = ttk.Treeview(resultados_frame, columns=colunas, show='headings', height=12)
        
        self.tree.heading('Arquivo', text='Arquivo')
        self.tree.heading('Linha Digitável', text='Linha Digitável')
        self.tree.heading('Valor', text='Valor (R$)')
        self.tree.heading('Vencimento', text='Vencimento')
        self.tree.heading('QR Code', text='QR Code')
        self.tree.heading('Status', text='Status')
        self.tree.heading('Páginas', text='Páginas')
        
        self.tree.column('Arquivo', width=150, minwidth=100)
        self.tree.column('Linha Digitável', width=180, minwidth=150)
        self.tree.column('Valor', width=100, minwidth=80, anchor='e')
        self.tree.column('Vencimento', width=100, minwidth=80, anchor='center')
        self.tree.column('QR Code', width=80, minwidth=60, anchor='center')
        self.tree.column('Status', width=100, minwidth=80, anchor='center')
        self.tree.column('Páginas', width=60, minwidth=50, anchor='center')
        
        scrollbar_v = ttk.Scrollbar(resultados_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_h = ttk.Scrollbar(resultados_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_v.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_h.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Tags para colorir linhas baseado no status
        self.tree.tag_configure('sucesso', background='#d4edda')
        self.tree.tag_configure('parcial', background='#fff3cd')
        self.tree.tag_configure('erro', background='#f8d7da')
        
        self.tree.bind("<Button-3>", self.mostrar_menu_contexto)
        self.tree.bind("<Double-1>", self.copiar_item_duplo_clique)
        
        self.menu_contexto = tk.Menu(self.root, tearoff=0)
        self.menu_contexto.add_command(label="📋 Copiar Linha Digitável", command=lambda: self.copiar_dados('linha'))
        self.menu_contexto.add_command(label="📱 Copiar QR Code", command=lambda: self.copiar_dados('qrcode'))
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="📄 Ver Todos os Dados", command=self.mostrar_dados_completos)
        self.menu_contexto.add_command(label="🔍 Debug - Ver Texto Extraído", command=self.mostrar_debug_texto)
        
        self.tooltip = None
        self.tree.bind("<Motion>", self.mostrar_tooltip)
        self.tree.bind("<Leave>", self.ocultar_tooltip)
        
        stats_frame = ttk.Frame(main_frame)
        stats_frame.grid(row=6, column=0, columnspan=3, pady=(10, 0), sticky=tk.W)
        self.label_stats = ttk.Label(stats_frame, text="", font=('Arial', 9))
        self.label_stats.grid(row=0, column=0)
        
        self.configurar_estilo()
        self.dados_completos = {}
        self.verificar_queue()
    
    def configurar_estilo(self):
        style = ttk.Style()
        style.theme_use('vista')
        style.configure('Accent.TButton', font=('Arial', 10, 'bold'))
    
    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os boletos PDF")
        if pasta:
            self.pasta_selecionada.set(pasta)
            try:
                pdfs = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
                self.status_atual.set(f"Pasta selecionada: {len(pdfs)} arquivos PDF encontrados")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível acessar a pasta:\n{e}")

    def iniciar_processamento(self):
        if not self.pasta_selecionada.get():
            messagebox.showerror("Erro", "Por favor, selecione uma pasta primeiro!")
            return
        if not os.path.isdir(self.pasta_selecionada.get()):
            messagebox.showerror("Erro", "A pasta selecionada não existe!")
            return
        self.btn_processar.config(state='disabled')
        self.limpar_resultados()
        thread = threading.Thread(target=self.processar_boletos_thread, daemon=True)
        thread.start()

    def processar_boletos_thread(self):
        try:
            pasta = self.pasta_selecionada.get()
            arquivos_pdf = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
            total_arquivos = len(arquivos_pdf)
            if total_arquivos == 0:
                self.queue.put(('erro', "Nenhum arquivo PDF encontrado na pasta!"))
                self.queue.put(('fim', None))
                return
            
            boletos_processados = []
            for i, nome_arquivo in enumerate(arquivos_pdf):
                caminho_completo = os.path.join(pasta, nome_arquivo)
                progresso = ((i + 1) / total_arquivos) * 100
                self.queue.put(('progresso', progresso, f"Processando ({i+1}/{total_arquivos}): {nome_arquivo}"))
                
                dados = self.extrair_dados_boleto_avancado(caminho_completo)
                if dados:
                    boletos_processados.append(dados)
                    self.queue.put(('resultado', dados))
            
            self.queue.put(('progresso', 100, f"Processamento concluído! {len(boletos_processados)} de {total_arquivos} boletos analisados."))
            
            if boletos_processados:
                self.salvar_resultados(pasta, boletos_processados)
                
        except Exception as e:
            self.queue.put(('erro', f"Erro durante processamento: {str(e)}"))
        finally:
            self.queue.put(('fim', None))

    def salvar_resultados(self, pasta: str, dados: List[Dict]) -> None:
        """Salva os resultados em múltiplos formatos"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Excel principal
        df = pd.DataFrame(dados)
        caminho_excel = os.path.join(pasta, f"resumo_boletos_{timestamp}.xlsx")
        df_limpo = df.drop(columns=['Texto_Bruto'], errors='ignore')
        df_limpo.to_excel(caminho_excel, index=False)
        
        # CSV para backup
        caminho_csv = os.path.join(pasta, f"resumo_boletos_{timestamp}.csv")
        df_limpo.to_csv(caminho_csv, index=False, encoding='utf-8-sig')
        
        if self.var_backup_dados.get():
            # Salva dados brutos para debug
            dados_debug = []
            for item in dados:
                if 'Texto_Bruto' in item:
                    dados_debug.append({
                        'Arquivo': item['Arquivo'],
                        'Texto_Extraido': item['Texto_Bruto'],
                        'Total_Paginas': item.get('Total_Paginas', 'N/A'),
                        'Status_Processamento': item.get('Status', 'N/A')
                    })
            
            if dados_debug:
                df_debug = pd.DataFrame(dados_debug)
                caminho_debug = os.path.join(pasta, f"debug_textos_{timestamp}.xlsx")
                df_debug.to_excel(caminho_debug, index=False)
        
        self.queue.put(('excel_salvo', caminho_excel))

    def verificar_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                tipo = item[0]
                if tipo == 'progresso':
                    self.progresso.set(item[1])
                    self.status_atual.set(item[2])
                elif tipo == 'resultado':
                    self.adicionar_resultado(item[1])
                elif tipo == 'excel_salvo':
                    messagebox.showinfo("Sucesso", f"Resumo salvo em:\n{item[1]}")
                elif tipo == 'erro':
                    messagebox.showerror("Erro", item[1])
                    self.status_atual.set("Erro no processamento")
                elif tipo == 'fim':
                    self.btn_processar.config(state='normal')
        except queue.Empty:
            pass
        self.root.after(100, self.verificar_queue)
    
    def adicionar_resultado(self, dados):
        linha_str = str(dados.get('Linha Digitável', '')).replace('\n', ' ').replace('\r', '')
        linha_display = linha_str[:40] + '...' if len(linha_str) > 40 else linha_str
        valor_str = f"{dados.get('Valor', 'N/A'):.2f}" if isinstance(dados.get('Valor'), (int, float)) else dados.get('Valor', 'N/A')
        
        # Determina o status e a tag de cor
        status = dados.get('Status', 'Incompleto')
        if status == 'Completo':
            tag = 'sucesso'
        elif status == 'Parcial':
            tag = 'parcial'
        else:
            tag = 'erro'
        
        valores_exibicao = (
            dados.get('Arquivo', ''), 
            linha_display, 
            valor_str,
            dados.get('Vencimento', ''), 
            'Sim' if dados.get('QR Code') != 'Não encontrado' else 'Não',
            status,
            str(dados.get('Total_Paginas', 'N/A'))
        )
        
        item_id = self.tree.insert('', tk.END, values=valores_exibicao, tags=(tag,))
        self.dados_completos[item_id] = dados
        self.atualizar_estatisticas()

    def atualizar_estatisticas(self):
        total_items = len(self.dados_completos)
        items_completos = sum(1 for dados in self.dados_completos.values() if dados.get('Status') == 'Completo')
        items_parciais = sum(1 for dados in self.dados_completos.values() if dados.get('Status') == 'Parcial')
        items_com_qr = sum(1 for dados in self.dados_completos.values() if dados.get('QR Code') != 'Não encontrado')
        soma_valores = sum(dados.get('Valor', 0) for dados in self.dados_completos.values() if isinstance(dados.get('Valor'), (int, float)))
        
        self.label_stats.config(text=f"Total: {total_items} | ✅ Completos: {items_completos} | ⚠️ Parciais: {items_parciais} | 📱 Com QR: {items_com_qr} | 💰 Soma: R$ {soma_valores:.2f}")

    def limpar_resultados(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.dados_completos.clear()
        self.progresso.set(0)
        self.status_atual.set("Pronto para processar")
        self.label_stats.config(text="")
    
    def mostrar_menu_contexto(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.menu_contexto.tk_popup(event.x_root, event.y_root)
    
    def copiar_item_duplo_clique(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.copiar_dados('linha')
    
    def copiar_dados(self, tipo):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um item primeiro!")
            return
        item_id = selection[0]
        dados = self.dados_completos.get(item_id)
        if not dados:
            messagebox.showerror("Erro", "Dados não encontrados!")
            return
        if tipo == 'linha':
            texto = dados.get('Linha Digitável', '')
            tipo_nome = "Linha Digitável"
        elif tipo == 'qrcode':
            texto = dados.get('QR Code', '')
            tipo_nome = "QR Code (PIX Copia e Cola)"
            if texto == 'Não encontrado':
                messagebox.showwarning("Aviso", "QR Code não encontrado neste boleto!")
                return
        if texto and texto != 'Não encontrado':
            self.root.clipboard_clear()
            self.root.clipboard_append(texto)
            self.root.update()
            preview = texto[:70] + '...' if len(texto) > 70 else texto
            messagebox.showinfo("✅ Copiado!", f"{tipo_nome} copiado para área de transferência:\n\n{preview}")
        else:
            messagebox.showwarning("Aviso", f"{tipo_nome} não encontrado!")
    
    def mostrar_dados_completos(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um item primeiro!")
            return
        item = selection[0]
        dados = self.dados_completos.get(item)
        if not dados:
            messagebox.showerror("Erro", "Dados não encontrados!")
            return
        janela_detalhes = tk.Toplevel(self.root)
        janela_detalhes.title(f"📄 Detalhes - {dados.get('Arquivo', 'Arquivo')}")
        janela_detalhes.geometry("800x600")
        janela_detalhes.configure(bg='#f0f0f0')
        main_frame = ttk.Frame(janela_detalhes, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        titulo = ttk.Label(main_frame, text=f"📄 {dados.get('Arquivo', 'Arquivo')}", font=('Arial', 14, 'bold'))
        titulo.pack(pady=(0, 15), anchor=tk.W)
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=('Consolas', 10), bg='white', relief='sunken', bd=1, height=20)
        scrollbar_texto = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar_texto.set)
        conteudo = "📊 DADOS EXTRAÍDOS DO BOLETO\n" + "="*60 + "\n\n"
        for campo, valor in dados.items():
            if campo != 'Texto_Bruto':  # Não mostra o texto bruto aqui
                conteudo += f"🏷️  {campo}:\n{valor}\n\n"
        text_widget.insert(tk.END, conteudo)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_texto.pack(side=tk.RIGHT, fill=tk.Y)
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.pack(pady=(15, 0))
        ttk.Button(botoes_frame, text="📋 Copiar Linha Digitável",
                   command=lambda: self.copiar_dado_janela(dados.get('Linha Digitável', ''), 'Linha Digitável')).pack(side=tk.LEFT, padx=(0, 10))
        if dados.get('QR Code') != 'Não encontrado':
            ttk.Button(botoes_frame, text="📱 Copiar QR Code",
                       command=lambda: self.copiar_dado_janela(dados.get('QR Code', ''), 'QR Code')).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(botoes_frame, text="❌ Fechar",
                   command=janela_detalhes.destroy).pack(side=tk.RIGHT)
    
    def mostrar_debug_texto(self):
        """Mostra o texto bruto extraído do PDF para debug"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um item primeiro!")
            return
        item = selection[0]
        dados = self.dados_completos.get(item)
        if not dados or 'Texto_Bruto' not in dados:
            messagebox.showwarning("Aviso", "Texto bruto não disponível para este item! Ative a opção 'Salvar dados brutos para debug' antes de processar.")
            return
            
        janela_debug = tk.Toplevel(self.root)
        janela_debug.title(f"🔍 Debug - {dados.get('Arquivo', 'Arquivo')}")
        janela_debug.geometry("900x700")
        janela_debug.configure(bg='#f0f0f0')
        
        main_frame = ttk.Frame(janela_debug, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        titulo = ttk.Label(main_frame, text=f"🔍 Texto Extraído - {dados.get('Arquivo', 'Arquivo')}", 
                           font=('Arial', 14, 'bold'))
        titulo.pack(pady=(0, 15), anchor=tk.W)
        
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=('Consolas', 9), bg='#f8f9fa', 
                              relief='sunken', bd=1, height=25)
        scrollbar_texto = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar_texto.set)
        
        texto_bruto = dados.get('Texto_Bruto', 'Texto não disponível')
        text_widget.insert(tk.END, texto_bruto)
        text_widget.config(state=tk.DISABLED)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_texto.pack(side=tk.RIGHT, fill=tk.Y)
        
        ttk.Button(main_frame, text="❌ Fechar", 
                   command=janela_debug.destroy).pack(pady=(15, 0))
    
    def exportar_relatorio_detalhado(self):
        """Exporta um relatório mais detalhado"""
        if not self.dados_completos:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar!")
            return
            
        pasta = self.pasta_selecionada.get()
        if not pasta:
            messagebox.showerror("Erro", "Pasta não selecionada!")
            return
            
        dados_para_export = list(self.dados_completos.values())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        try:
            # Cria um relatório mais detalhado
            df = pd.DataFrame(dados_para_export)
            caminho_relatorio = os.path.join(pasta, f"relatorio_detalhado_{timestamp}.xlsx")
            
            with pd.ExcelWriter(caminho_relatorio, engine='openpyxl') as writer:
                # Aba principal
                df_main = df.drop(columns=['Texto_Bruto'], errors='ignore')
                df_main.to_excel(writer, sheet_name='Resumo', index=False)
                
                # Aba com estatísticas
                valores_encontrados = [d.get('Valor', 0) for d in dados_para_export if isinstance(d.get('Valor'), (int, float))]
                
                stats_data = {
                    'Métrica': ['Total de Boletos', 'Boletos Completos', 'Boletos Parciais', 
                                'Com QR Code', 'Soma Total (R$)', 'Média de Valor (R$)'],
                    'Valor': [
                        len(dados_para_export),
                        sum(1 for d in dados_para_export if d.get('Status') == 'Completo'),
                        sum(1 for d in dados_para_export if d.get('Status') == 'Parcial'),
                        sum(1 for d in dados_para_export if d.get('QR Code') != 'Não encontrado'),
                        sum(valores_encontrados),
                        sum(valores_encontrados) / max(1, len(valores_encontrados))
                    ]
                }
                df_stats = pd.DataFrame(stats_data)
                df_stats.to_excel(writer, sheet_name='Estatísticas', index=False)
            
            messagebox.showinfo("Sucesso", f"Relatório detalhado salvo em:\n{caminho_relatorio}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar relatório:\n{str(e)}")
    
    def copiar_dado_janela(self, texto, tipo):
        if texto and texto != 'Não encontrado':
            self.root.clipboard_clear()
            self.root.clipboard_append(texto)
            self.root.update()
            preview = texto[:70] + '...' if len(texto) > 70 else texto
            messagebox.showinfo("✅ Copiado!", f"{tipo} copiado:\n\n{preview}")
    
    def mostrar_tooltip(self, event):
        item = self.tree.identify_row(event.y)
        if item and item in self.dados_completos:
            if hasattr(self, '_tooltip_job'):
                self.root.after_cancel(self._tooltip_job)
            self._tooltip_job = self.root.after(500, lambda: self._criar_tooltip(event, item))
        else:
            self.ocultar_tooltip(event)

    def _criar_tooltip(self, event, item):
        dados = self.dados_completos[item]
        if self.tooltip:
            self.tooltip.destroy()
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.configure(bg='#ffffe0', relief='solid', bd=1)
        valor_formatado = f"R$ {dados.get('Valor'):.2f}" if isinstance(dados.get('Valor'), (int, float)) else dados.get('Valor', 'N/A')
        texto_tooltip = f"📄 {dados.get('Arquivo', '')}\n💰 Valor: {valor_formatado}\n📅 Vencimento: {dados.get('Vencimento', 'N/A')}\n"
        texto_tooltip += f"📱 QR Code: {'Sim' if dados.get('QR Code') != 'Não encontrado' else 'Não'}\n"
        texto_tooltip += f"📊 Status: {dados.get('Status', 'N/A')}\n📑 Páginas: {dados.get('Total_Paginas', 'N/A')}\n"
        texto_tooltip += f"\n💡 Duplo clique = Copiar linha | Clique direito = Menu"
        label_tooltip = tk.Label(self.tooltip, text=texto_tooltip, bg='#ffffe0', font=('Arial', 9), justify=tk.LEFT, padx=5, pady=5)
        label_tooltip.pack()
        x = event.x_root + 20
        y = event.y_root + 10
        self.tooltip.geometry(f"+{x}+{y}")
    
    def ocultar_tooltip(self, event=None):
        if hasattr(self, '_tooltip_job'):
            self.root.after_cancel(self._tooltip_job)
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
    
    def abrir_pasta_resultado(self):
        pasta = self.pasta_selecionada.get()
        if pasta and os.path.isdir(pasta):
            os.startfile(pasta)
        else:
            messagebox.showwarning("Aviso", "Nenhuma pasta válida selecionada!")

    # === MÉTODOS DE EXTRAÇÃO APRIMORADOS ===
    
    def extrair_qrcode_do_pdf(self, doc: fitz.Document) -> Optional[str]:
        if not QR_CODE_DISPONIVEL:
            return "Dependências não instaladas"
        
        todos_qrcodes_encontrados = []
        for pagina in doc:
            for img in pagina.get_images(full=True):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                try:
                    pil_image = Image.open(io.BytesIO(image_bytes))
                    for qr in decode(pil_image):
                        todos_qrcodes_encontrados.append(qr.data.decode("utf-8"))
                except Exception:
                    continue
        if not todos_qrcodes_encontrados:
            return None
        for qr_text in todos_qrcodes_encontrados:
            if qr_text.startswith("000201"):
                return qr_text
        return todos_qrcodes_encontrados[0]

    def extrair_valor_inteligente(self, texto: str, qr_code: str = None) -> Tuple[Optional[float], str]:
        """Extração inteligente de valores com múltiplas estratégias"""
        # 1. Tenta extrair do QR Code primeiro (mais confiável)
        if qr_code and qr_code != 'Não encontrado':
            try:
                match_valor_qr = re.search(r'54\d{2}(\d+\.\d{2})', qr_code)
                if match_valor_qr:
                    valor = float(match_valor_qr.group(1))
                    if 0.01 <= valor <= 999999.99:
                        return valor, 'QR Code PIX'
            except (ValueError, AttributeError):
                pass
        
        # 2. Se não achou no QR, busca no texto
        padroes_valor = [
            (r'(?:VALOR\s+TOTAL|TOTAL\s+A\s+PAGAR|VALOR\s+DO\s+DOCUMENTO)\s*:?\s*R?\$?\s*([\d.,]+)', 'PDF - Campo Específico'),
            (r'(\d{1,3}(?:\.\d{3})*,\d{2})', 'PDF - Formato Monetário'),
        ]
        
        valores_encontrados = []
        for padrao, fonte in padroes_valor:
            matches = re.findall(padrao, texto, re.IGNORECASE)
            for match in matches:
                try:
                    valor_limpo = match.replace('.', '').replace(',', '.')
                    valor = float(valor_limpo)
                    if 0.01 <= valor <= 999999.99:
                        valores_encontrados.append(valor)
                except ValueError:
                    continue
        
        # 3. Heurística: se encontrou múltiplos valores, pega o maior (geralmente o total)
        if valores_encontrados:
            return max(valores_encontrados), 'PDF - Maior Valor'
        
        return None, 'Não encontrado'

    def extrair_data_vencimento_inteligente(self, texto: str, qr_code: str = None) -> Tuple[Optional[str], str]:
        """Extração inteligente de data de vencimento com múltiplas estratégias"""
        datas_candidatas = []
        padrao_data = re.finditer(r'(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{4})', texto)
        
        for match in padrao_data:
            data_str = match.group(1)
            posicao = match.start()
            try:
                data_normalizada = data_str.replace('.', '/').replace('-', '/')
                partes = data_normalizada.split('/')
                if len(partes) == 3:
                    dia, mes, ano = int(partes[0]), int(partes[1]), int(partes[2])
                    if 1 <= dia <= 31 and 1 <= mes <= 12 and 2000 <= ano <= 2050:
                        datas_candidatas.append({'data': data_normalizada, 'posicao': posicao})
            except (ValueError, TypeError):
                continue
        
        # 1. Tenta extrair do QR Code primeiro
        if qr_code and qr_code != 'Não encontrado':
            try:
                match_venc_qr = re.search(r'Venc[.:]\s*(\d{1,2}[./]\d{1,2}[./]\d{4})', qr_code, re.IGNORECASE)
                if match_venc_qr:
                    return match_venc_qr.group(1).replace('.', '/'), 'QR Code'
            except AttributeError:
                pass
        
        if not datas_candidatas:
            return None, 'Não encontrado'
        
        # 2. Procura por palavras-chave
        palavras_vencimento = [r'VENCIMENTO', 'VENC.', 'VENC:', 'PAGAR ATÉ', 'DATA LIMITE', 'DATA DE VENCIMENTO']
        for data_info in datas_candidatas:
            inicio_contexto = max(0, data_info['posicao'] - 50)
            contexto = texto[inicio_contexto:data_info['posicao']]
            for palavra in palavras_vencimento:
                if re.search(palavra, contexto, re.IGNORECASE):
                    return data_info['data'], f'PDF - Contexto ({palavra})'
        
        # 3. Pega a última data do documento
        return datas_candidatas[-1]['data'], 'PDF - Última Encontrada'

    def extrair_linha_digitavel_melhorada(self, texto: str) -> Tuple[Optional[str], str]:
        """
        Extração definitiva da linha digitável, tentando os padrões mais comuns em ordem.
        """
        # 1. Tenta o padrão de Boleto Bancário (47 dígitos, mais complexo e específico)
        #    Usa s* para aceitar "zero ou mais espaços", corrigindo a regressão.
        padrao_bancario = re.compile(r'(\d{5}\.?\d{5}\s*?\d{5}\.?\d{6}\s*?\d{5}\.?\d{6}\s*?\d\s*?\d{14})')
        match = padrao_bancario.search(texto)
        if match:
            return match.group(1).strip(), 'PDF - Boleto Bancário'

        # 2. Se não achou, tenta o padrão de Conta Convênio (48 dígitos)
        #    Este geralmente tem espaços, então usamos s+ (um ou mais espaços).
        padrao_convenio = re.compile(r'(\d{11,12}\s+\d{11,12}\s+\d{11,12}\s+\d{11,12})')
        match = padrao_convenio.search(texto)
        if match:
            return match.group(1).strip(), 'PDF - Conta Convênio'
            
        return None, 'Não encontrado'

    def extrair_dados_boleto_avancado(self, caminho_pdf: str) -> Optional[Dict[str, str]]:
        """Versão aprimorada da extração com análise inteligente"""
        try:
            doc = fitz.open(caminho_pdf)
            total_paginas = len(doc)
        except Exception as e:
            return {"Arquivo": os.path.basename(caminho_pdf), "Erro": str(e), "Status": "Erro", "Total_Paginas": 0}

        textos_por_pagina = [{'numero': i+1, 'texto': p.get_text("text")} for i, p in enumerate(doc)]
        texto_completo = "\n\n--- PÁGINA {} ---\n\n".format(textos_por_pagina[0]['numero']).join([p['texto'] for p in textos_por_pagina])

        dados_boleto = {
            "Arquivo": os.path.basename(caminho_pdf), "Total_Paginas": total_paginas,
            "Linha Digitável": "Não encontrado", "Valor": "Não encontrado",
            "Vencimento": "Não encontrado", "QR Code": "Não encontrado", "Status": "Erro"
        }
        if self.var_backup_dados.get():
            dados_boleto["Texto_Bruto"] = texto_completo

        # Extrações
        qr_code = self.extrair_qrcode_do_pdf(doc)
        if qr_code: dados_boleto["QR Code"] = qr_code

        linha, fonte_linha = self.extrair_linha_digitavel_melhorada(texto_completo)
        if linha: dados_boleto["Linha Digitável"], dados_boleto["Fonte_Linha"] = linha, fonte_linha

        valor, fonte_valor = self.extrair_valor_inteligente(texto_completo, qr_code)
        if valor: dados_boleto["Valor"], dados_boleto["Fonte_Valor"] = valor, fonte_valor

        vencimento, fonte_vencimento = self.extrair_data_vencimento_inteligente(texto_completo, qr_code)
        if vencimento: dados_boleto["Vencimento"], dados_boleto["Fonte_Vencimento"] = vencimento, fonte_vencimento

        # Determina status final
        campos_ok = sum(1 for k in ["Linha Digitável", "Valor", "Vencimento"] if dados_boleto[k] != "Não encontrado")
        if campos_ok >= 3:
            dados_boleto["Status"] = "Completo"
        elif campos_ok >= 1:
            dados_boleto["Status"] = "Parcial"
        else:
            dados_boleto["Status"] = "Não Encontrado"

        doc.close()
        
        if not self.var_backup_dados.get():
            for k in ["Fonte_Linha", "Fonte_Valor", "Fonte_Vencimento", "Texto_Bruto"]:
                dados_boleto.pop(k, None)
        
        return dados_boleto

def main():
    root = tk.Tk()
    app = ExtratorBoletosGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()