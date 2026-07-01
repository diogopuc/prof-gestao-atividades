# prof-gestao-atividades

Automação para extrair a lista de professores da **Agenda Docente PUCPR**, a partir da página **Aprovações > Gestão de Atividades**, e exportar os resultados para uma planilha Excel.

## O que a automação faz

1. Abre o navegador (Microsoft Edge) na página de login da Agenda Docente PUCPR e **pausa a execução**, aguardando o login manual do usuário no terminal (pressionar ENTER após autenticar).
2. Acessa a página de workflows.
3. Seleciona o módulo **"Gestão de Atividades"** no filtro.
4. Mantém o status **"Em andamento"** (já vem selecionado por padrão).
5. Seleciona o intervalo de datas configurado, navegando pelo calendário (mês/ano) e clicando nos dias de início e fim.
6. Clica em **Filtrar** para aplicar os critérios de busca.
7. Percorre todas as páginas de resultados, extraindo os nomes dos professores exibidos (via leitura do texto dos elementos `<span>` que estejam em maiúsculas).
8. Remove nomes duplicados, mantendo a ordem original.
9. Salva o resultado em um arquivo Excel (`.xlsx`), com colunas `#` e `Nome do Professor`.

## Requisitos

- Python 3.8+
- Microsoft Edge instalado no computador
- Dependências:
  ```bash
  pip install playwright openpyxl
  playwright install msedge
  ```

## Configuração

As variáveis de configuração ficam no topo do arquivo [main.py](main.py):

| Variável        | Descrição                                              | Exemplo        |
|-----------------|---------------------------------------------------------|----------------|
| `URL_LOGIN`     | URL base da Agenda Docente PUCPR                         | `https://agendadocente.pucpr.br` |
| `DATA_INICIO`   | Data inicial do filtro (formato `DD/MM/AAAA`)            | `01/07/2026`   |
| `DATA_FIM`      | Data final do filtro (formato `DD/MM/AAAA`)              | `07/07/2026`   |
| `ARQUIVO_SAIDA` | Nome do arquivo Excel gerado (montado automaticamente a partir das datas) | `professores_01_07_2026_07_07_2026.xlsx` |

Ajuste `DATA_INICIO` e `DATA_FIM` antes de cada execução, conforme o período desejado.

## Como executar

```bash
python main.py
```

O navegador abre em modo visível (`headless=False`) e com um pequeno atraso entre ações (`slow_mo=200`) para facilitar o acompanhamento visual da automação.

> **Login manual:** ao iniciar, o script abre a página de login e o terminal exibe a mensagem `Faça o login manualmente na janela do navegador e pressione ENTER aqui para continuar...`. Faça o login normalmente na janela do Edge e, em seguida, volte ao terminal e pressione **ENTER** para que a automação prossiga.

Ao final da execução, é gerado um arquivo `.xlsx` no diretório do projeto contendo a lista de professores únicos encontrados no período informado.

## Estrutura do arquivo de saída

| # | Nome do Professor |
|---|--------------------|
| 1 | NOME DO PROFESSOR  |
| 2 | ...                |

## Observações

- A extração depende da estrutura HTML atual do site (seletores CSS/texto específicos). Alterações no layout da Agenda Docente PUCPR podem exigir ajustes no script.
- O script imprime no console o progresso da extração (página atual, quantidade de nomes por página e total acumulado).
