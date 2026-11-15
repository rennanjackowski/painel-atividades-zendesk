# Painel de Atividades para Zendesk

Favorito do Chrome (bookmarklet) que acompanha os tickets que eu trato no Zendesk, registra cada tratativa em uma planilha Excel e, no fim da semana, gera um relatório em PDF para os gestores.

Criei esta ferramenta para o meu próprio dia a dia de trabalho: antes, o controle semanal de atividades era preenchido à mão.

> Esta é a versão pública do código. Nomes, logotipo e dados da empresa foram removidos, e nenhum dado de ticket acompanha o repositório.

## O que ele faz

- **Captura as tratativas sozinho.** Consulta as auditorias dos tickets pela API do próprio Zendesk, usando a sessão já aberta no navegador. Só faz leitura: nunca escreve no Zendesk.
- **Uma linha por demanda.** Cada ticket ocupa uma linha na planilha da semana; novas tratativas entram na mesma linha, com a data na frente.
- **Limpa o texto.** Remove saudação, despedida, assinatura e marcação Markdown antes de gravar.
- **Atividades sem ticket.** Podem ser lançadas direto pelo painel.
- **Relatório semanal em PDF.** Indicadores da semana, distribuição por unidade, visão geral e a ficha de cada demanda com o histórico.
- **Resumo opcional com IA local.** Usa os recursos de IA embutidos no Chrome, quando disponíveis; nada é enviado para fora da máquina.

## Como funciona

Não há instalação, servidor nem extensão. O arquivo `instalar.html` junta os scripts em um único favorito; ao clicar nele dentro do Zendesk, o painel aparece na página.

| Arquivo | Função |
| :-- | :-- |
| `instalar.html` | Monta o favorito e roda um diagnóstico da planilha, sem alterar o arquivo em disco |
| `atividades_zendesk.js` | Painel, captura das tratativas, regras de consolidação e gravação |
| `pdf_relatorio.js` | Gera o PDF sem biblioteca externa |
| `docs/guia-de-uso.md` | Guia de uso detalhado |

A leitura e a gravação da planilha usam a File System Access API do Chrome: o usuário escolhe a pasta e autoriza o acesso.

## Tecnologias

JavaScript puro (sem bibliotecas), API REST do Zendesk (somente leitura), File System Access API, HTML e CSS.

## Limitações

- Funciona apenas no Chrome ou Edge, por causa da File System Access API.
- O módulo que grava no arquivo `.xlsx` (`xlsx_cirurgico.js`) ainda não está neste repositório; sem ele o favorito não é montado.
- A busca do Zendesk leva alguns minutos para enxergar uma atualização recém-enviada.
- Foi feito para o formato de planilha que eu uso; outro formato pode exigir ajustes.

## Autor

Rennan Jackowski · [github.com/rennanjackowski](https://github.com/rennanjackowski)

Desenvolvido com apoio de assistente de IA.
