import xml.etree.ElementTree as ET
import os
import pandas as pd
from collections import Counter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import traceback

# Suprimir o aviso de depreciação do Tk
os.environ["TK_SILENCE_DEPRECATION"] = "1"

def validar_cpf(cpf):
    try:
        # Remove caracteres não numéricos
        cpf = ''.join(filter(str.isdigit, str(cpf)))
        
        # Verifica se tem 11 dígitos
        if len(cpf) != 11:
            return False
        
        # Verifica se todos os dígitos são iguais
        if cpf == cpf[0] * 11:
            return False
        
        # Calcula o primeiro dígito verificador
        soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
        resto = soma % 11
        digito1 = 0 if resto < 2 else 11 - resto
        
        if digito1 != int(cpf[9]):
            return False
        
        # Calcula o segundo dígito verificador
        soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
        resto = soma % 11
        digito2 = 0 if resto < 2 else 11 - resto
        
        return digito2 == int(cpf[10])
    except Exception as e:
        print(f"Erro ao validar CPF: {e}")
        return False

def contar_cpfs_em_xml(xml_file, output_dir):
    try:
        if not os.path.exists(xml_file):
            return f"Erro: O arquivo '{xml_file}' não foi encontrado.", 0, {}
        
        # Lê o arquivo XML
        tree = ET.parse(xml_file)
        root = tree.getroot()
        cpfs = []
        
        # Procura por tags <Cli> e verifica o atributo Cd
        for elem in root.findall('.//Cli'):
            cpf = elem.get('Cd')
            if cpf and validar_cpf(cpf):
                cpfs.append(cpf)
        
        # Conta CPFs válidos e identifica duplicados
        cpf_counts = Counter(cpfs)
        total_cpfs = len(cpf_counts)  # Total de CPFs únicos
        duplicados = {cpf: count for cpf, count in cpf_counts.items() if count > 1}
        total_duplicados = len(duplicados)
        
        # Prepara dados para o arquivo CSV
        max_rows = max(total_cpfs, total_duplicados) + 1  # +1 para o cabeçalho com totais
        col1 = [f"Total: {total_cpfs}"] + list(cpf_counts.keys())
        col2 = [f"Total: {total_duplicados}"] + [f"{cpf} (x{count})" for cpf, count in duplicados.items()]
        
        # Preenche listas com strings vazias para igualar tamanhos
        col1 += [""] * (max_rows - len(col1))
        col2 += [""] * (max_rows - len(col2))
        
        # Cria DataFrame e salva em CSV
        df = pd.DataFrame({
            "Quantidade de CPFs": col1,
            "Quantidade de CPFs Duplicados": col2
        })
        output_file = os.path.join(output_dir, "resultado_cpfs.csv")
        df.to_csv(output_file, index=False, encoding='utf-8')
        
        # Prepara mensagem de resultado
        mensagem = f"Total de CPFs válidos únicos: {total_cpfs}\n"
        if total_cpfs > 0:
            mensagem += f"Exemplos de CPFs: {', '.join(list(cpf_counts.keys())[:5])}\n"
        mensagem += f"Total de CPFs duplicados: {total_duplicados}\n"
        if total_duplicados > 0:
            mensagem += f"CPFs duplicados: {duplicados}\n"
        mensagem += f"Arquivo CSV gerado: {output_file}"
        
        return mensagem, total_cpfs, duplicados
    
    except FileNotFoundError:
        return f"Erro: O arquivo '{xml_file}' não foi encontrado.", 0, {}
    except ET.ParseError:
        return "Erro: Arquivo XML inválido ou mal formatado.", 0, {}
    except Exception as e:
        return f"Erro inesperado: {str(e)}", 0, {}

class App:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("Contador de CPFs em XML")
            self.root.geometry("600x400")
            
            # Frame principal
            self.frame = ttk.Frame(self.root, padding="10")
            self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Botão para selecionar arquivo XML
            self.btn_select = ttk.Button(self.frame, text="Selecionar Arquivo XML", command=self.select_file)
            self.btn_select.grid(row=0, column=0, columnspan=2, pady=10)
            
            # Label para mostrar o arquivo selecionado
            self.file_label = ttk.Label(self.frame, text="Nenhum arquivo selecionado")
            self.file_label.grid(row=1, column=0, columnspan=2, pady=5)
            
            # Botão para processar o arquivo
            self.btn_process = ttk.Button(self.frame, text="Processar", command=self.process_file, state="disabled")
            self.btn_process.grid(row=2, column=0, columnspan=2, pady=10)
            
            # Área de texto para resultados
            self.result_text = tk.Text(self.frame, height=15, width=60)
            self.result_text.grid(row=3, column=0, columnspan=2, pady=10)
            
            # Variável para armazenar o caminho do arquivo XML
            self.xml_file = ""
        except Exception as e:
            print(f"Erro ao inicializar a interface: {e}")
            traceback.print_exc()
            raise

    def select_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Arquivos XML", "*.xml")])
            if file_path:
                self.xml_file = file_path
                self.file_label.config(text=f"Arquivo: {os.path.basename(file_path)}")
                self.btn_process.config(state="normal")
        except Exception as e:
            print(f"Erro ao selecionar arquivo: {e}")
            traceback.print_exc()

    def process_file(self):
        try:
            if not self.xml_file:
                messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
                return
            
            # Pede ao usuário para selecionar o diretório de saída
            output_dir = filedialog.askdirectory(title="Selecionar Diretório para Salvar o CSV")
            if not output_dir:
                messagebox.showwarning("Aviso", "Nenhum diretório selecionado. O CSV não será gerado.")
                return
            
            # Processa o arquivo
            mensagem, total_cpfs, duplicados = contar_cpfs_em_xml(self.xml_file, output_dir)
            
            # Atualiza a área de texto com os resultados
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, mensagem)
            
            # Mostra mensagem de sucesso ou erro
            if total_cpfs > 0 or duplicados:
                messagebox.showinfo("Sucesso", "Processamento concluído. Verifique o arquivo CSV gerado.")
            else:
                messagebox.showerror("Erro", mensagem)
        except Exception as e:
            print(f"Erro ao processar arquivo: {e}")
            traceback.print_exc()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    except Exception as e:
        print(f"Erro ao iniciar o aplicativo: {e}")
        traceback.print_exc()
        input("Pressione Enter para sair...")