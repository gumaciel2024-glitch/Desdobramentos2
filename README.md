# Protótipo de triagem no SISBAJUD

## Preparar no Windows

Baixe o repositório em **Code → Download ZIP** e extraia todos os arquivos. Com o Python 3.11 ou superior e o Python Launcher instalados, dê dois cliques em **iniciar.bat**. No primeiro uso, mantenha a conexão à internet para a instalação do Chromium e das dependências.

Informe o caminho da planilha e o endereço HTTPS do SISBAJUD quando o programa pedir. Se ocorrer um erro, a janela permanece aberta para mostrar a mensagem.

Use os cabeçalhos **Número do Processo** e **Status** na primeira linha da aba ativa. Formate o número completo como texto. Por padrão, o protótipo pula as linhas que já têm um status preenchido.

Entre no SISBAJUD na janela aberta e acesse **Ordens Judiciais → Busca por filtros de pesquisa**. Volte ao terminal e pressione Enter. Informe um exemplo com uma única ordem e ícone de página apagado e outro com o ícone ativo.

| Resultado | Status |
| --- | --- |
| Nenhuma ordem encontrada | `Sem ordem` |
| Mais de uma ordem para o mesmo processo | `Ordem Exedente` |
| Uma ordem com o ícone de página apagado | `NEGATIVO` |
| Uma ordem com o ícone de página ativo | `Verificar operações` |

A saída é uma cópia com o sufixo `_triagem.xlsx`. Resultados inconclusivos deixam o status em branco e aparecem no terminal para revisão.

Para execução pelo terminal, use `.venv\Scripts\python.exe triagem.py "processos.xlsx" "ENDEREÇO_DO_SISBAJUD"`. Acrescente `--sem-ordem "mensagem exibida"` se o texto de busca vazia na sua tela for diferente. O ícone ativo sinaliza a necessidade de conferir os detalhes da ordem. A validação dos campos, ícones e paginação na interface real do SISBAJUD ainda está pendente.
