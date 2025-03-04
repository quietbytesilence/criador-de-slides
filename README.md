# Gerador Automático de Apresentações (PPTX e PDF)

Este é um script Python que converte arquivos de texto (`*.txt`) em apresentações de slides no formato **PPTX** (PowerPoint) e **PDF**, com personalização de estilos, fontes, cores e alinhamentos. O software é ideal para quem deseja automatizar a criação de apresentações a partir de conteúdo textual, como relatórios, resumos ou notas.

## Funcionalidades

- **Conversão de Texto para Slides**:
  - Divide o conteúdo do arquivo de texto em blocos (parágrafos) para criar slides individuais.
  - A primeira linha de cada bloco é usada como título do slide, e o restante como corpo.

- **Personalização de Estilos**:
  - **Tema Escuro**: Fundo preto com texto branco (personalizável).
  - **Fontes Personalizadas**: Escolha fontes para títulos, subtítulos e corpo do texto.
  - **Tamanho de Fonte**: Defina tamanhos de fonte para títulos e corpo.
  - **Alinhamento**: Centralizado, à esquerda ou à direita para títulos e texto.

- **Exportação para PPTX e PDF**:
  - Gera automaticamente um arquivo **PPTX** (PowerPoint).
  - Converte o PPTX para **PDF** usando o **LibreOffice** (disponível no sistema).

- **Processamento em Lote**:
  - Processa todos os arquivos `.txt` na pasta atual, gerando uma apresentação para cada um.

## Como Usar

1. **Preparação**:
   - Coloque os arquivos de texto (`*.txt`) na mesma pasta do script.
   - Certifique-se de que o **LibreOffice** está instalado no sistema para a conversão para PDF.

2. **Execução**:
   - Execute o script Python:
   - Eu prefiro renomear o arquivo, mas vá com a sua preferência.
     ```bash
     python3 script.py
     ```
   - O script criará um arquivo `.pptx` e um `.pdf` para cada arquivo `.txt` encontrado.

3. **Personalização**:
   - Edite as variáveis no início do script para ajustar fontes, tamanhos, cores e alinhamentos.

## Exemplo de Estrutura do Arquivo de Texto

Cada bloco de texto (separado por duas quebras de linha) será convertido em um slide. A primeira linha é o título, e as linhas seguintes são o corpo do slide.
Dê preferência para utilizar (nas configurações default) no máximo 12 linhas (70 caracteres por linha). Opcionalmente você pode realizar modificações.

Exemplo (`exemplo.txt`):
```
Título do Slide 1
Este é o conteúdo do primeiro slide.
Pode ter múltiplas linhas.

Título do Slide 2
Aqui está o conteúdo do segundo slide.
```

## Requisitos

- Python 3.x
- Biblioteca `python-pptx`:
  ```bash
  pip install python-pptx
  ```
- LibreOffice (para conversão para PDF).

## Personalização

No início do script, você pode personalizar:
- **Cores**: Fundo e texto.
- **Fontes**: Para títulos, subtítulos e corpo.
- **Tamanhos de Fonte**: Para títulos e corpo.
- **Alinhamentos**: Centralizado, à esquerda ou à direita.

Exemplo de personalização:
```python
COR_FUNDO = RGBColor(0, 0, 0)  # Preto
COR_TEXTO = RGBColor(255, 255, 255)  # Branco

FONTE_TITULO_PRINCIPAL = "Times New Roman"
TAMANHO_TITULO_PRINCIPAL = Pt(44)
ALINHAMENTO_TITULO_PRINCIPAL = PP_ALIGN.CENTER
```

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests para melhorias, correções de bugs ou novas funcionalidades.

## Licença

Este projeto está licenciado sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.


