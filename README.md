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
