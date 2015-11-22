# Planejamento-do-Louvor
Planilha para geração automática de PowerPoint para louvor e planejamento de programas e escalas.

<h3>Instruções para gerar cifras e slides:</h3>
- Os arquivos "Todas as cifras" e "Todos os slides" devem estar na mesma pasta da planilha e devem ter uma música por página/slide, numeradas idênticas à planilha "Índice", abaixo.
- Exemplo: se há 150 músicas na planilha "Índice", o doc "Todas as cifras" deve ter 150 páginas, e o ppt "Todos os slides" deve ter 150 slides.

<h3>Instruções para incluir uma música:</h3>
Cada nova música deve ser adicionada em 3 lugares:
- O nome e o número da música na planilha "Índice"
- As cifras da música no arquivo "Todas as cifras", apontado na planilha "Programa"
- Os slides da música no arquivo "Todos os slides", apontado na planilha "Programa"

<h3>Mensagem "Não é possível executar a macro"</h3>
- Vá em Opções do Excel - Central de Confiabilidade - Configurações da Central de Confiabilidade - Configurações de Macro e selecione "Habilitar todas as macros" e "Confiar no acesso ao modelo de objeto do Projeto VBA"

<h3>Mensagem "Erro ao carregar DLL"</h3>
1. Feche a planilha
2. Desabilite as macros nas Opções do Excel (veja as instruções acima)
3. Abra a planilha
4. Digite Alt+F11, vá no menu Ferramentas - Referências e desmarque as referências que começam com "AUSENTE".
5. Salve e feche a planilha
6. Habilite as macros nas Opções do Excel (veja as instruções acima)

SOLUÇÃO DEFINITIVA: Instale o Service Pack mais recente do seu Office.
- Office 2007: 
- http://www.microsoft.com/pt-br/download/details.aspx?id=27838
- Office 2010:
- http://www.microsoft.com/pt-br/download/details.aspx?id=39667
- Office 2013:
- http://www.microsoft.com/pt-br/download/details.aspx?id=42017
