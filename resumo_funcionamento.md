# PDFler Bot - Suplementos

Este bot foi desenvolvido para extrair dados de PDFs de suplementos e compará-los com o banco de dados antes de atualizar o estoque.

## Funcionamento Principal

O bot agora possui um fluxo em duas etapas:

1. **Extração:** Lê o PDF e extrai os produtos, quantidades e preços.
2. **Comparação:** Compara os produtos extraídos com o banco de dados e gera um relatório.
3. **Atualização Manual:** Permite ao usuário decidir quando atualizar o estoque após revisar o relatório.

## Principais Melhorias

1. **Extração Especializada para Suplementos:**
   - Reconhece produtos com nomes compostos
   - Extrai corretamente códigos, quantidades e preços
   - Processa formatos específicos de PDFs de suplementos

2. **Correspondência Inteligente:**
   - Normaliza nomes de produtos para melhorar comparações
   - Utiliza palavras-chave específicas para suplementos
   - Atribui maior peso para correspondências de termos como "creatina", "whey", etc.
   - Ignora marcas e embalagens, focando no tipo de produto

3. **Interface Melhorada:**
   - Visualização dos produtos extraídos
   - Relatório de comparação com o banco de dados
   - Atualização manual de estoque após revisão

## Como Usar

1. Execute o arquivo `bot_suplementos.py`
2. Clique em "Adicionar PDF" para selecionar um PDF de suplementos
3. O bot processará o PDF e extrairá os produtos
4. Clique em "Ver Produtos" para verificar os itens extraídos
5. Clique em "Ver Relatório" para ver a comparação com o banco de dados
6. Se estiver tudo correto, clique em "Atualizar Estoque" para confirmar as mudanças

## Exemplo de Correspondências

O sistema compara produtos como:

PDF: "BLACKSKULL - CREATINE 300G - BLACK"
DB: "Creatina Monohidratada 300g - Black Skull"
- Alta correspondência por conter "creatina/creatine" e "300g"

PDF: "INTEGRAL MEDICA - CREATINA 300G HARDCORE"
DB: "Creatina Hardcore 300g - Integralmédica"
- Alta correspondência por conter "creatina", "hardcore" e "300g"

## Vantagens da Nova Abordagem

1. **Segurança:** Atualização manual permite revisão antes de modificar o estoque
2. **Precisão:** Melhor correspondência de produtos com o algoritmo especializado
3. **Rastreabilidade:** Geração de relatórios para referência futura
4. **Flexibilidade:** Funciona com diferentes formatos de PDFs de suplementos 