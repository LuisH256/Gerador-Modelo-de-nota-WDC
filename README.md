# Gerador de Notas WDC 📑

[Português](#português) | [English](#english)

---

## Português

### 📝 Descrição
O **Gerador de Notas WDC** é uma aplicação desktop desenvolvida em Python para automatizar a criação de modelos de Notas Fiscais (Excel) para o setor de **SAC/RMA** da WDC Networks. Ele elimina o erro de colagem do Excel, preenche os campos corretamente, garantindo que os dados de produtos, valores e informações fiscais (como CFOP e Natureza de Operação) estejam sempre padronizados e corretos.

### ✨ Funcionalidades Principais
* **Automação de Excel:** Preenchimento automático nos campos, usando `.xlsx` com formatação rica (cores, negrito e quebras de linha).
* **Integração com OneDrive:** Detecta automaticamente pastas do OneDrive para salvamento direto na nuvem ou em diretório local.
* **Interface Intuitiva:** Formulário limpo desenvolvido em Tkinter para inserção rápida de códigos, descrições e quantidades.
* **Padronização de RMA:** Insere automaticamente dados da Livetech da Bahia e regras de impostos (ICMS/CFOP).

### 🛠️ Tecnologias e Dependências
* **Python 3.x**
* **openpyxl:** Manipulação de arquivos Excel.
* **Pillow (PIL):** Renderização da interface visual e logos.
* **Tkinter:** Interface gráfica de usuário (GUI).

**Para instalar as dependências:**
```bash
pip install openpyxl pillow pyinstaller
```
**Como gerar o arquivo executavel .exe**
```bash
pyinstaller --noconsole --onefile --add-data "wdc.png;." --add-data "Modelo de nota.xlsx;." app.py
```

---

## English

### 📝 Description
The **WDC Invoice Generator** is a desktop application developed in Python to automate the creation of Invoice templates (Excel) for the **Customer Service (SAC/RMA)** department at WDC Networks. It eliminates repetitive manual entry, ensuring that product data, values, and fiscal information (such as CFOP and Operation Nature) are always standardized and accurate.

### ✨ Key Features
* **Excel Automation:** Automatically fills `.xlsx` templates with rich formatting (colors, bold text, and line breaks).
* **OneDrive Integration:** Automatically detects OneDrive folders for direct cloud saving or local directory fallback.
* **Intuitive Interface:** A clean Tkinter-based form for quick insertion of codes, descriptions, and quantities.
* **RMA Standardization:** Automatically inserts Livetech da Bahia data and tax rules (ICMS/CFOP).

### 🛠️ Technologies & Dependencies
* **Python 3.x**
* **openpyxl:** Excel file manipulation.
* **Pillow (PIL):** Visual interface rendering and logos.
* **Tkinter:** Graphical User Interface (GUI).

**To install dependencies:**
```bash
pip install openpyxl pillow pyinstaller
```
**How to Generate the Executable .exe**
```bash
pyinstaller --noconsole --onefile --add-data "wdc.png;." --add-data "Modelo de nota.xlsx;." app.py
```

**Desenvolvido por / Developed by: LuisH256**
