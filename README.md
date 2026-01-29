# Ferramentas de Produtividade

Este é um programa desenvolvido em Python que combina duas ferramentas úteis em uma interface com abas: **Comparador de Planilhas** e **Conversor de PDF**.

## 🚀 Funcionalidades

### 📊 Comparador de Planilhas
- Comparação de planilhas Excel (.xlsx, .xls)
- Seleção de múltiplas colunas para comparação
- Algoritmo de similaridade configurável
- Normalização de texto (remover acentos, espaços extras, etc.)
- Detecção automática de colunas CPF
- Pré-visualização dos resultados
- Suporte a drag & drop

### 📄 Conversor de PDF
- **Conversão de Imagens para PDF**: Converte PNG, JPG, JPEG, GIF, BMP, TIFF, WEBP
- **Monitoramento Automático**: Monitora uma pasta e converte imagens automaticamente
- **Conversão Manual**: Selecione ou arraste imagens para converter
- **Junção de PDFs**: Junta múltiplos PDFs em um único arquivo
- **Conversões Especiais**: Excel → PDF, PDF → Word, PDF → Imagem
- **Sistema de Logs**: Registra todas as atividades com timestamp
- **Exclusão Automática**: Remove arquivos originais após conversão

## 🎨 Interface

- Interface moderna com abas separadas
- Temas claro e escuro
- Drag & drop para imagens
- Log de atividades em tempo real
- Interface responsiva e intuitiva

## 📦 Instalação

### 1. Instalar dependências
```bash
pip install -r requirements.txt
```

### 2. Executar o programa
```bash
python comparador.py
```

## 🔧 Dependências

- **Python 3.7+**
- **pandas** - Manipulação de planilhas
- **PyQt5** - Interface gráfica
- **rapidfuzz** - Algoritmo de similaridade
- **Pillow** - Processamento de imagens
- **reportlab** - Geração de PDFs
- **PyPDF2** - Manipulação de PDFs
- **openpyxl** - Manipulação de arquivos Excel
- **python-docx** - Criação de documentos Word
- **pdf2image** - Conversão de PDF para imagens

## 📖 Como usar

### Comparador de Planilhas
1. Vá para a aba "📊 Comparador de Planilhas"
2. Selecione a primeira planilha
3. Selecione a segunda planilha
4. Marque as colunas que formam a chave de comparação
5. Ajuste a similaridade desejada (0-100%)
6. Escolha o local de saída
7. Clique em "Comparar"

### Conversor de PDF

#### Monitoramento Automático
1. Vá para a aba "📄 Conversor de PDF"
2. Clique em "📂 Selecionar Pasta" para escolher a pasta de monitoramento
3. (Opcional) Clique em "📁 Pasta de Saída" para escolher onde salvar os PDFs
4. Use os botões de controle:
   - **▶️ Iniciar**: Inicia o monitoramento da pasta
   - **⏸️ Pausar**: Pausa temporariamente o monitoramento
   - **⏹️ Cancelar**: Para completamente o monitoramento
5. Coloque imagens na pasta monitorada - elas serão convertidas automaticamente!

#### Conversão Manual
1. Clique em "📂 Selecionar Imagens" ou arraste imagens para a área
2. Escolha a pasta de saída
3. Clique em "🔄 Converter Imagens"

#### Junção de PDFs
1. Clique em "📂 Selecionar PDFs"
2. Escolha os PDFs que deseja juntar
3. Digite o nome do PDF final
4. Clique em "🔗 Juntar PDFs"

#### Conversões Especiais
1. **Escolha o tipo de conversão**: Use os dropdowns "De" e "Para" para selecionar o formato
2. **Selecione o arquivo**: Clique em "📂 Selecionar Arquivo" (filtros automáticos baseados no tipo)
3. **Configure opções**: Marque "Manter arquivo original" se desejar
4. **Converta**: Clique em "🔄 Converter Arquivo"

**Conversões disponíveis:**
- 📊 Excel → PDF, Word
- 📄 PDF → Word, Imagem
- 📝 Word → PDF
- 🖼️ Imagem → PDF

## 📋 Log de Atividades

O programa mantém um log detalhado de todas as atividades:
- Conversões realizadas
- Erros encontrados
- Status do monitoramento
- Timestamps de todas as operações

## 🎯 Recursos Avançados

- **Drag & Drop**: Arraste imagens diretamente para a interface
- **Monitoramento em Tempo Real**: Detecta automaticamente novos arquivos
- **Controle Total do Monitoramento**: Botões Iniciar, Pausar e Cancelar
- **Estados Visuais**: Status colorido (Verde=Ativo, Laranja=Pausado, Vermelho=Parado)
- **Sistema de Conversões Flexível**: Dropdowns "De" e "Para" para qualquer combinação
- **Filtros Automáticos**: Seleção de arquivo adapta-se ao tipo escolhido
- **Múltiplos Formatos**: Suporta Excel, PDF, Word, Imagens
- **Qualidade Alta**: Conversões em 300 DPI
- **Interface Responsiva**: Adapta-se ao tema selecionado

### 🎮 Controles do Monitoramento

- **▶️ Iniciar**: Fica habilitado quando uma pasta é selecionada
- **⏸️ Pausar**: Fica habilitado quando o monitoramento está ativo
- **⏹️ Cancelar**: Fica habilitado quando o monitoramento está ativo ou pausado
- **Status Visual**: Mostra o estado atual do monitoramento com cores

## 🐛 Solução de Problemas

Se você receber um erro sobre dependências não instaladas:
```bash
pip install -r requirements.txt
```

Ou instale individualmente:
```bash
pip install Pillow reportlab PyPDF2 openpyxl python-docx pdf2image
```

## 👨‍💻 Autor

Desenvolvido por **GABRIEL**

## 📝 Versão

Versão 1.2026.01.29