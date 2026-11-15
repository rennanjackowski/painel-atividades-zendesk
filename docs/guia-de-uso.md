# Automação de Registro de Atividades — v2

Mesma ideia da versão anterior (nada é instalado, é um favorito no Chrome), com
três mudanças grandes:

1. **Uma linha por demanda.** O ticket ocupa uma única linha na semana e cada
   nova tratativa completa essa linha.
2. **Uma aba armada cobre todas.** Você clica no favorito em uma aba só; o que
   for atualizado em qualquer outra aba é capturado do mesmo jeito.
3. **PDF detalhado.** Indicadores, distribuição por unidade, visão geral e a
   ficha de cada demanda com o histórico das tratativas.

A parte do Google Chat foi removida: atividade sem ticket se lança pelo campo de
atividade avulsa do próprio painel.

---

## Montar a pasta (uma vez)

Esta pasta é independente da antiga. Nada da estrutura antiga é alterado.

```
Documentos\Automacao Atividades v2\
   Controle de atividades semanais.xlsx   <- copie da pasta antiga
   logo.png                                      <- copie da pasta antiga
   navegador\
      instalar.html
      atividades_zendesk.js
      pdf_relatorio.js
      xlsx_cirurgico.js                          <- copie de "navegador" antigo
      README.md
```

`backups\` e `arquivos_semanais\` são criadas sozinhas na primeira gravação.

Depois:

1. Abra `navegador\instalar.html`, selecione os **3** arquivos `.js` e arraste o
   botão **📋 Painel V2** para a barra de favoritos.
2. Rode o diagnóstico do passo 3 com a planilha da pasta nova.
3. No Zendesk, dê F5, clique no favorito novo e em **Conectar planilha**,
   escolhendo a pasta **nova**.

**Melhor momento para trocar:** segunda-feira, depois que a versão antiga fechou
a semana — a planilha já estará limpa. Para trocar no meio da semana, copie a
planilha como está e clique em **Organizar** (veja abaixo).

A v2 guarda o estado dela em chaves próprias (`aa2_*`), separadas da antiga. Na
primeira execução ela **copia** da versão antiga a lista de tickets já lidos,
para não lançar nada em dobro; a antiga não é modificada. Só não use as duas ao
mesmo tempo: se o favorito antigo estiver ativo na aba, o novo avisa e não liga.

---

## Uma linha por demanda

| Situação | O que acontece na planilha |
|---|---|
| Primeira tratativa do ticket na semana | cria a linha, com o texto simples |
| Nova tratativa no mesmo ticket | entra na **mesma** linha, com a data na frente |
| Mudança de status sem comentário | só o Status da linha muda |
| Resposta que é só cortesia ("Cientes, agradecemos a gentileza") | não entra no histórico; o Status é atualizado |
| Ticket renomeado no Zendesk | a coluna Problema passa a mostrar o nome atual |
| Mesma tratativa capturada duas vezes | ignorada |

A célula **O Que Foi Feito** de uma demanda com várias tratativas fica assim:

```
09/11 - Fornecedor esteve na unidade; a chuva impediu o serviço. Reagendado para amanhã, 08h.
10/11 - Manutenção realizada com sucesso, segue validação: (imagem em anexo)
12/11 - Colaborador confirmou o funcionamento dos refletores.
```

Com uma tratativa só, fica apenas o texto (a data já está na coluna Data). A
altura da linha é liberada para o Excel ajustar sozinho ao abrir.

**Limpeza do texto**, além da que já existia (saudação, despedida, assinatura):

- `![](https://.../image.png)` vira `(imagem em anexo)`;
- `**Fulano**` no início vira `Para Fulano:`; negrito no meio perde os asteriscos;
- `Olá,` agora é removido de verdade (a versão anterior deixava passar);
- `Cientes, muito obrigado` não sobra mais como `Cientes, muito`.

**Atividade avulsa também se completa.** Lance de novo com o mesmo assunto:

```
Inventário da unidade | Iniciei a conferência da planilha de patrimônio
Inventário da unidade | Conferência finalizada e enviada ao gestor | Resolvido
```

A segunda entra na linha da primeira e muda o Status.

**Corrigir** ganhou o campo **nova tratativa**: o que você escrever ali entra no
histórico daquela linha com a data de hoje, sem montar o formato na mão.

**Organizar** converte a planilha atual para o formato novo: as várias linhas do
mesmo ticket viram uma, com o histórico datado, cortesias saem e a marcação
Markdown é limpa. Mostra antes quantas linhas viram quantas demandas, faz backup
e confere o arquivo gerado antes de gravar.

Para voltar ao formato antigo (uma linha por atualização), mude
`consolidarTicketRepetido` para `false` no topo de `atividades_zendesk.js`.

---

## Uma aba armada cobre todas

Sem extensão o Chrome não deixa nenhum script ligar sozinho ao abrir uma página
— isso não mudou e não tem como mudar por aqui. O que mudou é que **não é mais
preciso ligar em cada aba**.

Antes, o painel só acompanhava os tickets abertos na aba em que ele estava. Agora
a cada varredura ele pergunta ao próprio Zendesk (busca, só leitura) em quais
tickets você comentou ou é o responsável nesta semana, e acompanha todos.

Na prática:

- clique no favorito **uma vez**, em uma aba — o ideal é uma aba do Zendesk
  fixada (botão direito na aba → *Fixar*) que você deixa aberta o dia todo;
- abra tickets em quantas abas quiser, sem clicar em nada nelas;
- o painel mostra `alcance ...... todas as abas (busca do Zendesk)`.

Detalhes honestos:

- A busca do Zendesk leva alguns minutos para enxergar uma atualização recém
  enviada. A tratativa feita em outra aba aparece na planilha com esse atraso.
- Se a aba armada for fechada ou recarregada (F5), clique no favorito de novo.
  Nada se perde: na volta ele recupera tudo o que foi feito na semana.
- Se você armar em duas abas, só uma faz a captura; a outra fica de reserva e
  assume se a primeira fechar.
- Se a busca falhar, o painel avisa e volta a acompanhar só os tickets abertos
  na aba armada.
- O Chrome pode "adormecer" abas paradas para economizar memória. Se a captura
  parar depois de muito tempo sem você olhar a aba armada, adicione o endereço
  do Zendesk em `chrome://settings/performance` → *Manter estes sites sempre
  ativos*.

---

## Resumo semanal em PDF

| Bloco | O que traz |
|---|---|
| **Panorama da semana** | demandas, tratativas, concluídas, em andamento |
| **A semana em resumo** | parágrafo da IA local (só com o resumo ativado) |
| **Distribuição por unidade** | demandas por UF, lida do início do assunto (`GO - ...`, `PE | ...`) |
| **Visão geral das demandas** | ticket, assunto, período, nº de tratativas, situação |
| **Demandas em andamento** | ficha de cada uma, com o histórico datado |
| **Demandas concluídas na semana** | idem |
| **Atividades sem ticket** | as avulsas |

Com a IA local ativada (**Ativar resumo**), cada demanda com mais de uma
tratativa ganha também uma **síntese** de uma ou duas frases acima do histórico.
A síntese nunca substitui o histórico, que continua logo abaixo. Com muitas
demandas isso leva um ou dois minutos; o painel avisa quando começa.

O botão **Resumo** gera o PDF sem fechar a semana, para conferir antes.

Para o formato curto (sem as fichas), `detalharDemandas: false`. Para desligar
só as sínteses, `sinteseIA: false`.

---

## Configurações novas (topo de `atividades_zendesk.js`)

| Opção | Padrão | Para que serve |
|---|---|---|
| `consolidarTicketRepetido` | `true` | uma linha por demanda |
| `ignorarCortesia` | `true` | cortesia não entra no histórico |
| `dataDaDemanda` | `'primeira'` | `'ultima'` faz a coluna Data mostrar a última tratativa |
| `descobrirPelaBusca` | `true` | uma aba armada cobre todas |
| `detalharDemandas` | `true` | fichas por demanda no PDF |
| `sinteseIA` | `true` | síntese por demanda no PDF |

Depois de alterar qualquer `.js`: reabra o `instalar.html`, selecione os 3
arquivos de novo, arraste o favorito e apague o antigo de mesmo nome.

Proteções, fechamento semanal, mover o painel, logo e selo: tudo igual à versão
anterior.
