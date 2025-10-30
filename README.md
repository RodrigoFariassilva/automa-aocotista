
Automação de Cotistas - Processamento de Movimentações Financeiras
Este projeto tem como objetivo automatizar o processamento de movimentações financeiras de cotistas em fundos de investimento, gerando um arquivo .prn com formatação específica para integração com sistemas legados.
📌 Funcionalidades

Carregamento de dados de cotistas a partir de arquivos CSV
Leitura e consolidação de transações financeiras de múltiplos arquivos Excel
Normalização e limpeza de dados (acentos, espaços, termos irrelevantes)
Identificação de titulares e associação com clientes
Processamento de transações (aplicações e resgates)
Mapeamento de fundos para IDs específicos
Formatação de dados (datas, valores monetários)
Geração de arquivo .prn com colunas posicionadas em locais fixos

🗂 Estrutura do Código
Funções utilitárias: Normalização de texto, remoção de acentos e termos
Carregamento de arquivos: CSV e Excel com tratamento de erros
Transformações: Seleção de colunas, identificação de titulares, merge com cotistas
Formatação final: Reordenação de colunas, formatação de datas e valores
Exportação: Geração de arquivo .prn com colunas nas posições:

📦 Dependências
pandas
openpyxl
numpy
unicodedata
os
json
