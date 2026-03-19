import os
import sys
import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
from PIL import Image, ImageTk

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class AppWDC:
    def __init__(self, root):
        self.root = root
        self.root.title("WDC Networks - Gerador de Notas")
        self.root.geometry("500x680")
        self.root.configure(bg="#F8FAFC")
        self.root.resizable(False, False)

        base_onedrive = os.environ.get('ONEDRIVECOMMERCIAL') or os.environ.get('ONEDRIVE')
        sub_pasta = r"Exemplo de nota NORMAL\NOTAS"
        self.caminho_onedrive = os.path.join(base_onedrive, sub_pasta) if base_onedrive else None
        
        self.pasta_local = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "Modelos de Notas")

        # --- HEADER UI ---
        self.header = tk.Frame(root, bg="#004a99", height=80)
        self.header.pack(fill="x")
        self.header.pack_propagate(False)
        header_content = tk.Frame(self.header, bg="#004a99")
        header_content.pack(expand=True)

        try:
            img = Image.open(resource_path("wdc.png"))
            img.thumbnail((100, 40), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(img)
            tk.Label(header_content, image=self.logo, bg="#004a99").pack(side="left", padx=10)
        except: pass

        tk.Label(header_content, text="GERADOR DE NOTAS", font=("Segoe UI", 14, "bold"), fg="white", bg="#004a99").pack(side="left")

        # --- FORMULÁRIO ---
        self.body = tk.Frame(root, bg="#F8FAFC", padx=30, pady=20)
        self.body.pack(fill="both", expand=True)
        self.entradas = {}
        campos = [("CÓDIGOS PRODUTOS", "cod"), ("DESCRIÇÕES", "desc"), ("VALORES UNITÁRIOS", "val"), ("QUANTIDADES", "qtd"), ("N° CRG", "crg"), ("NOTAS FISCAIS AQUISIÇÃO", "nfs")]

        for label_text, key in campos:
            f = tk.Frame(self.body, bg="#F8FAFC")
            f.pack(fill="x", pady=6)
            tk.Label(f, text=label_text, font=("Segoe UI", 9, "bold"), bg="#F8FAFC", fg="#64748B").pack(anchor="w")
            ent = tk.Entry(f, font=("Segoe UI", 11), relief="flat", bg="white", highlightthickness=1, highlightbackground="#E2E8F0")
            ent.pack(fill="x", ipady=6, pady=(2,0))
            self.entradas[key] = ent

        self.footer = tk.Frame(root, bg="#F8FAFC", pady=20)
        self.footer.pack(fill="x")
        self.btn_gerar = tk.Button(self.footer, text="GERAR DOCUMENTO", command=self.processar, bg="#004a99", fg="white", font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2", padx=40, pady=10)
        self.btn_gerar.pack()

    def limpar_campos(self):
        for ent in self.entradas.values(): ent.delete(0, tk.END)
        self.entradas['cod'].focus_set()

    def processar(self):
        crg = self.entradas['crg'].get().strip()
        if not crg:
            messagebox.showwarning("Campo Vazio", "Por favor, informe o N° CRG.")
            return

        try:
            wb = openpyxl.load_workbook(resource_path("Modelo de nota.xlsx"))
            ws = wb.active
            
            # Definição de Fontes
            f_preta = InlineFont(color='000000', b=True) 
            f_vermelha = InlineFont(color='FF0000', b=True)
            f_azul = InlineFont(color='0000FF', b=True)

            t_cod = "\n".join([x.strip() for x in self.entradas['cod'].get().split(';')])
            t_desc = "\n".join([f"-> {x.strip().upper()}" for x in self.entradas['desc'].get().split(';')])
            # Colocando o R$ dentro da fonte vermelha
            t_val = "\n".join([f"R$ {x.strip()}" for x in self.entradas['val'].get().split(';')])
            t_qtd = "\n".join([x.strip() for x in self.entradas['qtd'].get().split(';')])

            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        v = cell.value.upper().strip()
                        align_l = Alignment(wrapText=True, vertical='center', horizontal='left', indent=1)
                        align_c = Alignment(wrapText=True, vertical='center', horizontal='center')

                        # PROTEÇÃO DO RODAPÉ (Igual à imagem 15)
                        if "FALE CONOSCO ATRAVÉS DO NÚMERO" in v or "(73) 3222-5250" in str(cell.value):
                            if "TELEFONE" not in v: # Garante que não é a célula do destinatário
                                continue

                        # 1. NATUREZA DA OPERAÇÃO
                        if "NATUREZA DA OPERAÇÃO" in v:
                            cell.value = CellRichText([
                                TextBlock(f_preta, "NATUREZA DA OPERAÇÃO (DESCRIÇÃO SUGERIDA, ABAIXO)\n"),
                                TextBlock(f_vermelha, "REMESSA PARA CONSERTO")
                            ])
                            cell.alignment = align_l

                        # 2. DESTINATÁRIO (CABEÇALHO)
                        elif "LIVETECH DA BAHIA" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "NOME/ RAZÃO SOCIAL\n"), TextBlock(f_vermelha, "LIVETECH DA BAHIA INDUSTRIA E COMERCIO SA")])
                        elif "05.917.486" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "CNPJ / CPF\n"), TextBlock(f_vermelha, "05.917.486/0001-40")])
                        elif "ROD BA 262" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "ENDEREÇO\n"), TextBlock(f_vermelha, "ROD BA 262, RODOVIA ILHEUS X URUCUCA, S/N KM 2,8")])
                        elif "IGUAPE" in v and "BAIRRO" in v:
                             cell.value = CellRichText([TextBlock(f_preta, "BAIRRO\n"), TextBlock(f_vermelha, "IGUAPE")])
                        elif "45658-335" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "CEP\n"), TextBlock(f_vermelha, "45658-335")])
                        elif "ILHÉUS" in v and "MUNICÍPIO" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "MUNICÍPIO\n"), TextBlock(f_vermelha, "ILHÉUS")])
                        elif "3222-5250" in v and "TELEFONE" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "TELEFONE\n"), TextBlock(f_vermelha, "73 3222-5250")])
                        elif v == "BA" or (v.startswith("UF") and "BA" in v):
                            cell.value = CellRichText([TextBlock(f_preta, "UF\n"), TextBlock(f_vermelha, "BA")])
                        elif "63250303" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "INSCRIÇÃO ESTADUAL\n"), TextBlock(f_vermelha, "63250303")])

                        # 3. IMPOSTOS (AZUL)
                        elif "BASE DE CÁLCULO" in v and "ICMS" in v:
                            suffix = " SUBST." if "SUBST." in v else ""
                            cell.value = CellRichText([
                                TextBlock(f_preta, f"BASE DE CÁLCULO DO ICMS{suffix}\n"),
                                TextBlock(f_azul, "NÃO DESTACAR")
                            ])
                            cell.alignment = align_c

                        # 4. ITENS (Onde estava o erro do CFOP)
                        elif "CODI. PROD:" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "CODI. PROD:\n"), TextBlock(f_vermelha, t_cod)])
                        elif "DESCRIÇÃO DO PRODUTO:" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "DESCRIÇÃO DO PRODUTO:\n"), TextBlock(f_vermelha, t_desc)])
                        
                        # CORREÇÃO CFOP AQUI:
                        elif "CFOP:" in v:
                            cell.value = CellRichText([
                                TextBlock(f_preta, "CFOP:\n"), 
                                TextBlock(f_vermelha, "5915\nou\n6915")
                            ])
                            cell.alignment = align_c

                        elif "QTDE:" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "QTDE:\n"), TextBlock(f_vermelha, t_qtd)])
                            cell.alignment = align_c
                        elif "VL. UNIT:" in v:
                            cell.value = CellRichText([TextBlock(f_preta, "VL. UNIT:\n"), TextBlock(f_vermelha, t_val)])
                            cell.alignment = align_c
                        elif "SOLICITAÇÃO" in v:
                            cell.value = CellRichText([
                                TextBlock(f_preta, "SOLICITAÇÃO DE REPARO OU GARANTIA: "), TextBlock(f_vermelha, crg),
                                TextBlock(f_preta, "\nNOTAS FISCAIS DE AQUISIÇÃO: "), TextBlock(f_vermelha, self.entradas['nfs'].get()),
                                TextBlock(f_preta, "\nAOS CUIDADOS DE: SAC/RMA")
                            ])

            nome_arquivo = f"Modelo de nota CRG-{crg.replace('/', '-')}.xlsx"
            if self.caminho_onedrive and os.path.exists(self.caminho_onedrive):
                caminho_final = os.path.join(self.caminho_onedrive, nome_arquivo)
                wb.save(caminho_final)
                messagebox.showinfo("Sucesso", f"Salvo no OneDrive!\n\nArquivo: {nome_arquivo}")
            else:
                if not os.path.exists(self.pasta_local): os.makedirs(self.pasta_local)
                caminho_final = os.path.join(self.pasta_local, nome_arquivo)
                wb.save(caminho_final)
                messagebox.showinfo("Sucesso", f"Salvo localmente na pasta: Modelos de Notas")
            
            self.limpar_campos()

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppWDC(root)
    root.mainloop()