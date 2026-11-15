/* ==========================================================================
 * atividades_zendesk.js
 *
 * Roda dentro da aba do Zendesk (via favorito/bookmarklet). Nao instala nada.
 *
 * COMO CAPTURA
 *   Nao espiona cliques nem seletores CSS. Ele anota quais tickets voce abriu
 *   e consulta as AUDITORIAS desses tickets pela API do proprio Zendesk, usando
 *   a sessao que ja esta aberta no seu navegador. Cada auditoria e uma
 *   atualizacao realmente enviada, com id unico dado pelo servidor.
 *   Consequencias:
 *     - abrir/visualizar/filtrar nunca gera registro (nao cria auditoria);
 *     - o texto vem do que VOCE enviou (author_id == seu usuario);
 *     - duplicidade e impossivel (audit.id e unico);
 *     - se voce fechar o navegador, nada se perde: na proxima vez ele recupera.
 *
 * O QUE ELE NUNCA FAZ
 *   Nenhuma escrita no Zendesk. Apenas requisicoes GET de leitura.
 * ========================================================================== */

(function () {
  'use strict';

  if (window.__automacaoAtividades2) {
    window.__automacaoAtividades2.mostrarPainel();
    return;
  }
  /* A versao anterior (favorito antigo) nao pode rodar junto na mesma aba:
     as duas capturariam as mesmas atualizacoes, cada uma na sua planilha. */
  if (window.__automacaoAtividades) {
    alert('A versao ANTIGA do painel de atividades esta ativa nesta aba.\n\n' +
          'Recarregue a pagina (F5) e clique somente no favorito novo.');
    return;
  }

  /* ====================================================== CONFIGURACAO ==== */
  var CFG = {
    /* Nome preferido da planilha. Se ele NAO existir na pasta escolhida, a
       automacao procura o unico .xlsx que estiver la e usa esse. E o que
       permite entregar o painel a um colega sem editar uma linha de codigo:
       a planilha dele pode ter qualquer nome. Havendo mais de um .xlsx solto
       na pasta, ela para e diz quais achou, em vez de chutar. */
    nomeArquivo: 'Controle de atividades semanais.xlsx',
    pastaBackups: 'backups',
    pastaSemanais: 'arquivos_semanais',
    /* {nome} = primeiro nome lido do proprio Zendesk. */
    padraoNomeSemanal: 'Controle_{nome}_{dd}_{mm}_{aaaa}.xlsx',

    intervaloVarreduraSeg: 120,
    diasDeTicketsRastreados: 7,
    maxTicketsPorLote: 300,

    /* UMA ABA ARMADA COBRE TODAS. Em vez de depender dos tickets abertos NESTA
       aba, a varredura pergunta ao Zendesk (busca) em quais tickets voce
       comentou ou e o responsavel nesta semana. Entao tanto faz em qual aba, ou
       ate em qual computador, a atualizacao foi enviada: basta UMA aba com o
       painel armado. Se a busca falhar, vale o modo antigo (tickets abertos
       nesta aba). A busca do Zendesk leva alguns minutos para enxergar uma
       atualizacao recem-enviada; o ticket aberto na aba armada continua sendo
       visto na hora. */
    descobrirPelaBusca: true,
    consultasDeBusca: ['type:ticket commenter:me updated>={inicio}',
                       'type:ticket assignee:me updated>={inicio}'],
    maxPaginasBusca: 3,
    maxPaginasAuditoria: 10,

    /* Ticket que a API nega (403) para de ser tentado por este tempo - assim
       um so ticket sem acesso (forms privados, marca restrita etc.) nao fica
       gerando aviso a cada varredura. Reabrir o ticket no Zendesk cancela a
       espera e tenta de novo na proxima rodada. */
    horasSuspensaoTicketNegado: 12,

    /* UMA LINHA POR DEMANDA. true = o ticket ocupa uma unica linha na semana e
       cada nova tratativa completa essa linha: o "O Que Foi Feito" vira um
       historico datado (uma tratativa por linha da celula) e o Status acompanha
       o ticket. false = formato antigo, uma linha da planilha por atualizacao. */
    consolidarTicketRepetido: true,

    /* Resposta que e so cortesia ("Cientes, agradecemos a gentileza") nao entra
       no historico de uma demanda que ja tem linha: so o Status e atualizado.
       Se for a PRIMEIRA coisa registrada no ticket, entra normalmente - nenhuma
       demanda deixa de aparecer na planilha por causa disto. */
    ignorarCortesia: true,

    /* Data da linha da demanda: 'primeira' mantem o dia em que ela entrou na
       semana; 'ultima' passa a mostrar o dia da tratativa mais recente. */
    dataDaDemanda: 'primeira',

    /* Teto do historico dentro de uma celula (o Excel aceita 32 mil). */
    maxCaracteresCelula: 8000,
    registrarSemComentario: false,
    maxCaracteresComentario: 2000,
    manterBackups: 40,

    /* A planilha usa apenas dois status: Resolvido quando o ticket foi encerrado
       no Zendesk, Em Andamento em qualquer outra situacao. */
    statusResolvido: 'Resolvido',
    statusAndamento: 'Em Andamento',
    statusBaseResolvido: ['solved', 'closed'],
    usarStatusPersonalizado: false,

    /* Resumo do comentario. A limpeza por regra roda sempre (tira saudacao,
       despedida, assinatura e historico citado). A IA local do Chrome, quando
       ativada no painel, resume os textos longos - e roda 100% na sua maquina,
       sem chave de API e sem enviar nada para fora. */
    resumoIA: true,
    minCaracteresParaIA: 220,
    maxCaracteresResumo: 400,
    naoIdentificado: 'Nao identificado',
    ticketAvulso: '-',

    /* Anotacao avulsa curta JA E o titulo da atividade. Abaixo deste tamanho a
       IA nao e chamada para estruturar - so as regras. Sem isso o modelo
       encolhe demais: "Continuacao do curso disponibilizado pelo RH Pensamento
       Critico" virava so "Pensamento Critico RH" e o resto se perdia. */
    minCaracteresAvulsaIA: 170,

    /* "O que foi feito" nunca sai em branco numa avulsa: quando o texto nao
       descreve uma acao, vale a sua propria anotacao, com as suas palavras.
       Nada e inventado - so reaproveitado. */
    completarCampoFeito: true,

    /* Quantas linhas o botao "Corrigir" mostra para edicao direta. */
    linhasParaCorrigir: 8,

    fecharNaSexta: true,
    horaFechamento: 17,

    /* ------------------------------ resumo semanal em PDF ------------------
       Sai junto com o arquivo semanal, no padrao visual da empresa, para enviar
       aos gestores. O logo e o selo sao lidos da PASTA - basta largar os dois
       arquivos de imagem ao lado da planilha e eles aparecem no cabecalho. Sem
       eles o documento sai igual, so sem as imagens. */
    resumoPdf: true,
    padraoNomeResumo: 'Resumo_{nome}_{dd}_{mm}_{aaaa}.pdf',
    arquivoLogo: 'logo.png',
    arquivoSelo: 'selo_iso.png',

    /* O PDF traz o panorama, o resumo da semana, a distribuicao por unidade,
       a visao geral e o detalhamento de cada demanda com o historico das
       tratativas. Mude para false se quiser so as tres primeiras partes. */
    detalharDemandas: true,
    maxDemandasDetalhadas: 80,

    /* Sintese de 1-2 frases, pela IA local, para cada demanda com mais de uma
       tratativa. O historico completo aparece logo abaixo dela, entao a
       sintese nunca substitui informacao. Sem IA, o PDF sai igual, sem ela. */
    sinteseIA: true,
    maxSintesesIA: 30,
    maxCaracteresContextoIA: 4000
  };

  var CH = {
    tickets: 'aa2_tickets',
    atividades: 'aa2_atividades',
    estado: 'aa2_estado'
  };

  /* ============================================================ UTIL ===== */
  function agora() { return new Date(); }

  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function diaISO(d) {
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }

  function carimbo(d) {
    return diaISO(d).replace(/-/g, '') + '_' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  }

  function normalizar(t) {
    if (t === null || t === undefined) { return ''; }
    var s = String(t).normalize('NFKD').replace(/[̀-ͯ]/g, '');
    s = s.replace(/[°º]/g, 'o').toLowerCase();
    s = s.replace(/[^a-z0-9]+/g, ' ');
    return s.replace(/\s+/g, ' ').trim();
  }

  function soDigitos(v) { return String(v === null || v === undefined ? '' : v).replace(/\D/g, ''); }

  function ler(chave, padrao) {
    try {
      var v = localStorage.getItem(chave);
      return v ? JSON.parse(v) : padrao;
    } catch (e) { return padrao; }
  }

  function gravar(chave, valor) {
    try { localStorage.setItem(chave, JSON.stringify(valor)); } catch (e) { registrar('Falha ao salvar estado local: ' + e); }
  }

  /* Primeira execucao da versao nova: COPIA (so leitura) o que a versao antiga
     ja sabia - quais tickets voce abriu e ate qual auditoria cada um foi lido.
     Sem isso a v2 releria a semana inteira e lancaria de novo o que ja esta na
     planilha. As chaves antigas (aa_*) nao sao alteradas nem apagadas. */
  function herdarEstadoAntigo() {
    try {
      if (localStorage.getItem(CH.tickets) !== null ||
          localStorage.getItem(CH.atividades) !== null) { return false; }
      var t = localStorage.getItem('aa_tickets');
      var a = localStorage.getItem('aa_atividades');
      if (t === null && a === null) { return false; }
      if (t !== null) { localStorage.setItem(CH.tickets, t); }
      if (a !== null) {
        /* O que a versao antiga ainda nao tinha gravado fica com ela. */
        var mapa = JSON.parse(a) || {};
        Object.keys(mapa).forEach(function (id) { mapa[id].gravado = true; });
        localStorage.setItem(CH.atividades, JSON.stringify(mapa));
      }
      return true;
    } catch (e) { return false; }
  }

  var LOG = [];
  function registrar(msg) {
    var linha = pad(agora().getHours()) + ':' + pad(agora().getMinutes()) + ':' + pad(agora().getSeconds()) + '  ' + msg;
    LOG.push(linha);
    if (LOG.length > 300) { LOG.shift(); }
    try { console.log('[atividades] ' + msg); } catch (e) {}
    atualizarPainel();
  }

  function semanaDe(d) {
    var base = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    var offset = (base.getDay() + 6) % 7;            /* segunda = 0 */
    var inicio = new Date(base); inicio.setDate(base.getDate() - offset);
    var fim = new Date(inicio); fim.setDate(inicio.getDate() + 6);
    return { inicio: diaISO(inicio), fim: diaISO(fim) };
  }

  function dataDaCelula(v) {
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      var d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
      return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    }
    if (typeof v === 'string') {
      var m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/.exec(v.trim());
      if (m) {
        var ano = m[3].length === 2 ? ('20' + m[3]) : m[3];
        return ano + '-' + pad(parseInt(m[2], 10)) + '-' + pad(parseInt(m[1], 10));
      }
      m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v.trim());
      if (m) { return m[1] + '-' + m[2] + '-' + m[3]; }
    }
    return null;
  }

  /* ============================================ TEXTO E RESUMO =========== */

  var RE_SAUDACAO = new RegExp('^[\\s,;.!-]*\\b(bom dia|boa tarde|boa noite|ol[aá]|' +
    'prezad[oa]s?|car[oa]s?|senhor(?:es|a|as)?|pessoal|time|equipe)(?![A-Za-zÀ-ÿ])[\\s,;.!:-]*', 'i');

  var RE_DESPEDIDA = new RegExp('[\\s,;.!-]*\\b(?:muito\\s+)?(atenciosamente|att\\.?|abra[cç]os?|' +
    'cordialmente|grat[oa]|obrigad[oa]|agrade[cç]o|agradecemos|' +
    'fico [aà] disposi[cç][aã]o|permanecemos [aà] disposi[cç][aã]o)' +
    '\\b[\\s,;.!]*$', 'i');

  /* Despedida formal seguida do NOME de quem assina. So corta quando o que vem
     depois parece nome proprio: "Atenciosamente, Fulano" sai, mas
     "Grato pela ajuda, seguimos com o processo" fica. */
  var RE_ASSINATURA = new RegExp('[\\s,;.!-]*\\b(atenciosamente|att\\.?|abra[cç]os?|' +
    'cordialmente|grat[oa]|cordiais sauda[cç][oõ]es)\\b[\\s,;.!-]*' +
    '(?:[A-ZÀ-Ý][a-zà-ÿ\'.-]*\\s*){0,4}$', 'i');

  /* O comentario do Zendesk chega em Markdown. Na planilha e no PDF isso vira
     sujeira: "**Fulano**", "![](https://.../image.png)". Aqui a marcacao sai e
     o conteudo fica:
       imagem colada        -> "(imagem em anexo)"
       [texto](link)        -> "texto"
       **Fulano** no inicio -> "Para Fulano: ..."   (e uma mencao)
       **negrito** no meio  -> "negrito"                                       */
  function limparMarkdown(bruto) {
    var t = String(bruto || '');
    t = t.replace(/!\[[^\]]*\]\([^)]*\)/g, ' (imagem em anexo) ');
    t = t.replace(/\[([^\]]+)\]\((?:https?:|mailto:)[^)]*\)/g, '$1');
    t = t.replace(/^(\s*(?:(?:ol[aá]|bom dia|boa tarde|boa noite)[\s,!.]*)?)\*\*([^*\n]{2,60})\*\*[\s,:!-]*/i,
      function (tudo, antes, nome) { return antes + 'Para ' + nome.trim() + ': '; });
    t = t.replace(/\*\*([^*\n]+)\*\*/g, '$1').replace(/\*\*/g, '');
    t = t.replace(/(\(imagem em anexo\)\s*){2,}/g, '(imagens em anexo) ');
    return t;
  }

  /* Tira saudacao, despedida, assinatura e historico citado. Nenhuma palavra do
     conteudo e alterada - so o que nao e conteudo sai. */
  function limparComentario(bruto) {
    bruto = limparMarkdown(bruto);
    var original = String(bruto || '').replace(/\s+/g, ' ').trim();
    var t = String(bruto || '').replace(/\r\n?/g, '\n');

    t = t.split(/\n\s*(?:Em .{0,80}escreveu:|De:\s|-{3,}\s*Mensagem)/i)[0];
    t = t.split(/\n\s*(?:--|__)\s*\n/)[0];

    t = t.split('\n').filter(function (l) {
      var s = l.trim();
      if (/^>/.test(s)) { return false; }
      if (/^(tel|fone|celular|ramal|e-?mail|www\.|https?:)\b/i.test(s)) { return false; }
      return true;
    }).join('\n');

    t = t.replace(/[ \t]*\n[ \t]*/g, ' ').replace(/ {2,}/g, ' ').trim();

    var antes;
    do { antes = t; t = t.replace(RE_SAUDACAO, ''); } while (t !== antes && t);
    do { antes = t; t = t.replace(RE_DESPEDIDA, ''); } while (t !== antes && t);
    do { antes = t; t = t.replace(RE_ASSINATURA, ''); } while (t !== antes && t);

    t = t.replace(/^[\s,;.!-]+/, '').replace(/[\s,;-]+$/, '').trim();
    if (!t) { return original; }          /* so havia cortesia: devolve o original */
    return t.charAt(0).toUpperCase() + t.slice(1);
  }

  /* ----------------------------- so cortesia? ----------------------------
     "Cientes, agradecemos pela gentileza Arthur." nao e tratativa: e educacao.
     Num historico de demanda isso so ocupa espaco. A regra e conservadora:
     tira as palavras de cortesia e ve o que sobra. Sobrando ate duas palavras
     de conteudo, era cortesia. "Agradecemos, com isso encerramos a demanda"
     sobra bastante coisa e portanto fica. */
  var PALAVRAS_CORTESIA = ('ciente cientes agradecemos agradeco agradecimento obrigado ' +
    'obrigada muito pela pelo pelas pelos gentileza atencao retorno informacoes ' +
    'informacao disposicao permanecemos permaneceremos permaneco ficamos fico ' +
    'aguardo aguardamos aguardando sempre desde dispensada dispensadas envio ' +
    'para qualquer duvida duvidas esclarecimentos ajuda apoio colaboracao tambem ' +
    'prezado prezada prezados bom boa dia tarde noite ola tudo certo perfeito ' +
    'combinado entendido anotado recebido').split(' ');

  function ehSoCortesia(texto) {
    var palavras = normalizar(texto).split(' ').filter(function (p) {
      return p.length >= 4 && PALAVRAS_CORTESIA.indexOf(p) < 0;
    });
    return palavras.length <= 2;
  }

  /* Rede de seguranca de idioma: qualquer saida de IA passa por aqui. Se vier em
     ingles, e descartada e vale o texto original em portugues. */
  var PALAVRAS_EN = [' the ', ' has ', ' have ', ' been ', ' not ', ' yet ', ' because ',
    ' and ', ' of ', ' to ', ' is ', ' was ', ' were ', ' by ', ' from ', ' with ',
    ' that ', ' this ', ' its ', ' on ', ' for ', ' between ', ' received ', ' device '];
  var PALAVRAS_PT = [' o ', ' a ', ' os ', ' as ', ' de ', ' do ', ' da ', ' dos ', ' das ',
    ' em ', ' no ', ' na ', ' para ', ' com ', ' que ', ' foi ', ' nao ', ' ainda ',
    ' porque ', ' um ', ' uma ', ' pelo ', ' pela ', ' sobre ', ' ao '];

  function pareceIngles(texto) {
    var s = ' ' + String(texto || '').toLowerCase()
      .normalize('NFKD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z\s]/g, ' ').replace(/\s+/g, ' ') + ' ';
    var en = 0, pt = 0;
    PALAVRAS_EN.forEach(function (p) { if (s.indexOf(p) >= 0) { en++; } });
    PALAVRAS_PT.forEach(function (p) { if (s.indexOf(p) >= 0) { pt++; } });
    return en > pt;
  }

  function cortar(texto) {
    var t = String(texto || '');
    var lim = CFG.maxCaracteresResumo;
    if (lim > 0 && t.length > lim) { t = t.slice(0, lim).replace(/\s+\S*$/, '') + '...'; }
    return t;
  }

  /* ---------------------------- tirar o narrador -------------------------
     A IA gosta de narrar em terceira pessoa: "O analista Fulano informou que
     o equipamento chegou". Numa planilha de atividade isso e ruido - quem fez
     ja e voce, em toda linha. Aqui sai APENAS o sujeito; o verbo e todo o
     conteudo permanecem, entao nenhuma informacao se perde:

       "O analista Fulano informou que o equipamento chegou"
         -> "Informou que o equipamento chegou"

     Nome de terceiro nunca e removido: "Fulano reclamou que o ar
     condicionado parou" continua inteiro. So caem papeis genericos e o SEU
     proprio nome, lido de /users/me. */
  var PAPEIS = 'analista|agente|atendente|t[eé]cnic[oa]|respons[aá]vel|suporte|' +
               'operador|colaborador|autor';

  var VERBOS_NARRA = 'informou|informa|relatou|relata|explicou|explica|comunicou|' +
    'comunica|escreveu|escreve|respondeu|responde|disse|diz|afirmou|afirma|' +
    'mencionou|menciona|descreveu|descreve|registrou|registra|confirmou|confirma|' +
    'solicitou|solicita|pediu|pede|orientou|orienta|encaminhou|encaminha|' +
    'verificou|verifica|realizou|realiza|abriu|abre|acionou|aciona|enviou|envia|' +
    'atualizou|atualiza|resolveu|resolve|finalizou|finaliza|concluiu|conclui|' +
    'agendou|agenda|retornou|retorna|ajustou|ajusta|substituiu|substitui';

  var meuNome = '';        /* preenchido por descobrirUsuario() */

  function regexDoMeuNome() {
    var partes = String(meuNome || '')
      .normalize('NFKD').replace(/[̀-ͯ]/g, '')
      .split(/\s+/)
      .filter(function (p) { return p.length >= 3 && /^[A-Za-z][A-Za-z'.-]*$/.test(p); })
      .map(function (p) { return p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); });
    return partes.length ? '(?:' + partes.join('|') + ')' : null;
  }

  function prefixosDeNarrador() {
    var lista = [
      /* "Neste comentario, ..." / "No ticket, ..." */
      new RegExp('^n(?:o|a|este|esta)\\s+(?:coment[aá]rio|mensagem|texto|' +
                 'atualiza[cç][aã]o|ticket|chamado)[^,.:;]{0,30}[,.:;]\\s*', 'i'),
      /* "Resumo: ..." */
      /^(?:resumo|s[íi]ntese|em resumo|no geral)\s*[:\-–—]\s*/i,
      /* "Foi informado pelo analista que ..." */
      new RegExp('^foi\\s+(?:informad[oa]|relatad[oa]|comunicad[oa]|registrad[oa])\\s+' +
                 '(?:pel[oa]\\s+(?:' + PAPEIS + ')\\s+)?que\\s+', 'i'),
      /* "O analista Fulano informou que ..." -> "Informou que ..." */
      new RegExp('^(?:[oa]\\s+)?(?:' + PAPEIS + ')\\b' +
                 '(?:\\s+[A-ZÀ-Ý][\\wÀ-ÿ.\'-]*){0,3}\\s+(?=(?:' + VERBOS_NARRA + ')\\b)', 'i')
    ];
    var nome = regexDoMeuNome();
    if (nome) {
      /* "Fulano informou que ..." -> "Informou que ..." */
      lista.push(new RegExp('^(?:[oa]\\s+)?' + nome +
        '(?:\\s+[A-ZÀ-Ý][\\wÀ-ÿ.\'-]*){0,2}\\s+(?=(?:' + VERBOS_NARRA + ')\\b)', 'i'));
    }
    return lista;
  }

  function tirarNarrador(texto) {
    var t = String(texto || '').trim();
    if (!t) { return t; }
    var res = prefixosDeNarrador();
    var antes;
    do {
      antes = t;
      res.forEach(function (re) { t = t.replace(re, '').trim(); });
    } while (t !== antes && t);

    var original = String(texto || '').trim();
    if (!t) { return original; }                     /* sobrou nada: devolve tudo */
    /* Sobrou so um fragmento? Entao o que parecia narrador era o conteudo.
       Melhor a frase inteira do que gravar "Informou" sozinho na planilha. */
    if (t.split(/\s+/).length < 3 && original.split(/\s+/).length >= 3) { return original; }
    return t.charAt(0).toUpperCase() + t.slice(1);
  }

  /* Porta unica de saida de QUALQUER texto vindo da IA: tira aspas, marcadores
     de lista e o narrador. */
  function limparSaidaIA(saida) {
    var s = String(saida || '').replace(/\s+/g, ' ')
      .replace(/^[-*•]\s*/, '')
      .replace(/^["'“]+|["'”]+$/g, '')
      .trim();
    return tirarNarrador(s);
  }

  /* --------------------------- IA local do Chrome ------------------------ */
  /* Dois motores, nesta ordem de preferencia:
       1. Summarizer com outputLanguage 'pt' - especializado, mas so serve se o
          Chrome tiver o pacote de portugues;
       2. LanguageModel - generico, e aceita a instrucao de idioma no proprio
          prompt, entao resume em portugues mesmo sem pacote dedicado.
     Se nenhum servir, vale a limpeza por regra. Em qualquer caminho, toda saida
     passa por pareceIngles(): a planilha nunca recebe texto em ingles. */
  var resumidor = null;
  var sessaoLM = null;
  var modoResumo = 'regras';
  var estadoIA = 'nao verificada';

  var ESQUEMA_RESUMO = {
    type: 'object',
    properties: { resumo: { type: 'string' } },
    required: ['resumo'],
    additionalProperties: false
  };

  function monitorDownload(m) {
    m.addEventListener('downloadprogress', function (e) {
      estadoIA = 'baixando modelo ' + Math.round((e.loaded || 0) * 100) + '%';
      atualizarPainel();
    });
  }

  function prepararSummarizer(podeBaixar) {
    if (typeof Summarizer === 'undefined') { return Promise.resolve(null); }
    var IDIOMA = { expectedInputLanguages: ['pt'], outputLanguage: 'pt' };
    return Summarizer.availability(IDIOMA)['catch'](function () { return 'unavailable'; })
      .then(function (d) {
        if (d === 'unavailable' || (d === 'downloadable' && !podeBaixar)) { return null; }
        return Summarizer.create({
          type: 'tldr', format: 'plain-text', length: 'short',
          expectedInputLanguages: ['pt'], outputLanguage: 'pt',
          sharedContext: 'Comentarios de atendimento de TI em portugues do Brasil. ' +
                         'Resuma sempre em portugues do Brasil, comecando pela acao ' +
                         'ou pelo fato. Nao narre em terceira pessoa e nao use ' +
                         '"o analista", "o agente" nem nomes de quem atendeu.',
          monitor: monitorDownload
        })['catch'](function () { return null; });
      })['catch'](function () { return null; });
  }

  function prepararLM(podeBaixar) {
    if (sessaoLM) { return Promise.resolve(sessaoLM); }
    if (typeof LanguageModel === 'undefined') { return Promise.resolve(null); }
    return LanguageModel.availability()['catch'](function () { return 'unavailable'; })
      .then(function (d) {
        if (d === 'unavailable' || (d === 'downloadable' && !podeBaixar)) { return null; }
        if (d !== 'available') { estadoIA = 'baixando modelo'; atualizarPainel(); }
        return LanguageModel.create({
          monitor: monitorDownload,
          initialPrompts: [{
            role: 'system',
            content: 'Voce escreve EXCLUSIVAMENTE em portugues do Brasil, nunca em ' +
                     'ingles. Trabalha com anotacoes de atendimento de TI e nunca ' +
                     'inventa informacao que nao esteja no texto recebido.'
          }]
        }).then(function (s) { sessaoLM = s; return s; })['catch'](function () { return null; });
      })['catch'](function () { return null; });
  }

  /* Cada resumo comeca do zero: sem isso o historico cresce ate estourar. */
  function sessaoLimpa(podeBaixar) {
    return prepararLM(podeBaixar).then(function (s) {
      if (!s || typeof s.clone !== 'function') { return s; }
      return s.clone()['catch'](function () { return s; });
    })['catch'](function () { return null; });
  }

  function prepararIA(podeBaixar) {
    if (!CFG.resumoIA) {
      estadoIA = 'desligado na configuracao';
      return Promise.resolve(null);
    }
    if (modoResumo !== 'regras') { return Promise.resolve(true); }

    return prepararSummarizer(podeBaixar).then(function (s) {
      if (s) {
        resumidor = s;
        modoResumo = 'summarizer';
        estadoIA = 'ativa (pt-BR)';
        registrar('Resumo ativo em portugues (Summarizer).');
        atualizarPainel();
        return true;
      }
      return prepararLM(podeBaixar).then(function (lm) {
        if (lm) {
          modoResumo = 'prompt';
          estadoIA = 'ativa (pt-BR)';
          registrar('Resumo ativo em portugues (modelo local, idioma fixado no prompt).');
          atualizarPainel();
          return true;
        }
        estadoIA = (typeof Summarizer === 'undefined' && typeof LanguageModel === 'undefined')
          ? 'indisponivel neste Chrome' : 'clique em "Ativar resumo"';
        atualizarPainel();
        return null;
      });
    })['catch'](function (e) {
      estadoIA = 'erro';
      registrar('AVISO: IA local indisponivel (' +
                String((e && e.message) || e).slice(0, 70) +
                '). Seguindo com a limpeza por regra.');
      return null;
    });
  }

  /* Porta unica de saida da IA: nada passa daqui em ingles nem narrado. */
  function validarPortugues(saida, original) {
    var s = limparSaidaIA(saida);
    if (!s) { return null; }
    if (pareceIngles(s) && !pareceIngles(original)) {
      registrar('AVISO: a IA respondeu em ingles. Resumo descartado; gravado o ' +
                'texto original em portugues.');
      return null;
    }
    return s;
  }

  /* As mesmas regras ditas ao modelo e cobradas depois por tirarNarrador():
     instruir reduz o retrabalho, mas quem garante e o pos-processamento. */
  var REGRAS_ESCRITA =
    '- Escreva em portugues do Brasil. Nunca em ingles.\n' +
    '- Comece pela acao ou pelo fato. NUNCA comece com "O analista", "O agente",\n' +
    '  "O tecnico", "O atendente", "O usuario" nem com o nome de quem atendeu.\n' +
    '- Nao narre em terceira pessoa. Isto e um registro de atividade, nao um relato.\n' +
    '- Use somente o que esta no texto. Nao invente fato, nome, data, numero nem causa.\n' +
    '- Sem aspas e sem marcadores de lista.\n';

  function resumirComPrompt(limpo) {
    return sessaoLimpa(false).then(function (s) {
      if (!s) { return null; }
      function encerrar() {
        try { if (s !== sessaoLM) { s.destroy(); } } catch (e) {}
      }
      return s.prompt(
        'Reescreva o texto abaixo como uma anotacao curta de registro de atendimento,\n' +
        'em no maximo duas frases.\n' + REGRAS_ESCRITA + '\n' +
        'TEXTO:\n' + limpo,
        { responseConstraint: ESQUEMA_RESUMO }
      ).then(function (r) {
        var o = JSON.parse(r);
        var fim = validarPortugues(o && o.resumo, limpo);
        encerrar();
        return fim;
      })['catch'](function () { encerrar(); return null; });
    })['catch'](function () { return null; });
  }

  /* Limpeza sempre; IA so quando o texto e longo o bastante para valer. */
  function prepararTexto(bruto) {
    var limpo = limparComentario(bruto);
    if (!limpo || limpo.length < CFG.minCaracteresParaIA) {
      return Promise.resolve(cortar(limpo));
    }

    if (modoResumo === 'summarizer' && resumidor) {
      return resumidor.summarize(limpo, {
        context: 'Responda obrigatoriamente em portugues do Brasil.'
      }).then(function (r) {
        return cortar(validarPortugues(r, limpo) || limpo);
      })['catch'](function () { return cortar(limpo); });
    }

    if (modoResumo === 'prompt') {
      return resumirComPrompt(limpo).then(function (s) { return cortar(s || limpo); });
    }

    return Promise.resolve(cortar(limpo));
  }

  /* ------------------------- atividade avulsa ---------------------------- */
  var UFS = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG',
             'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'];

  var VERBOS_RELATO = 'reclamou|informou|solicitou|pediu|relatou|avisou|comunicou|' +
                      'questionou|reportou|sinalizou|abriu chamado';
  var VERBOS_ACAO = 'acionei|entrei em contato|solicitei|abri|encaminhei|verifiquei|' +
                    'orientei|repassei|cobrei|agendei|resolvi|conclu[ií]|finalizei|' +
                    'analisei|retornei|enviei|ajustei|configurei|instalei|substitu[ií]';

  /* Formas impessoais e futuras: "..., vai ser acionado um tecnico". Estas sao
     frases longas e especificas, entao da para separar tambem apos virgula sem
     risco de cortar a frase no meio. */
  var VERBOS_ACAO_IMPESSOAL =
    '(?:vai|ir[aá]|ser[aá]|foi|sera|est[aá] sendo|ficou de ser)\\s+' +
    '(?:ser\\s+)?(?:acionad[oa]|verificad[oa]|encaminhad[oa]|solicitad[oa]|abert[oa]|' +
    'agendad[oa]|substitu[ií]d[oa]|analisad[oa]|providenciad[oa]|tratad[oa])';

  /* "Fulano reclamou que o ar condicionado na BA nao esta funcionando"
        -> problema: "BA - O ar condicionado nao esta funcionando"
        -> feito:    ""  (nada foi descrito)
        -> status:   "Em Andamento"                                          */
  /* Data escrita no inicio: "16/09/2026 - ..." ou "16/09 - ...". Exige ano OU um
     separador explicito, para nao confundir com "10/10 servidores fora do ar". */
  function extrairData(texto) {
    var m = /^\s*(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-](\d{2,4}))?\s*([-–—:]?)\s*/.exec(texto);
    if (!m || (!m[3] && !m[4])) { return null; }
    var dia = parseInt(m[1], 10), mes = parseInt(m[2], 10);
    if (dia < 1 || dia > 31 || mes < 1 || mes > 12) { return null; }
    var hoje = new Date();
    var ano = m[3] ? (m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10))
                   : hoje.getFullYear();
    var d = new Date(ano, mes - 1, dia, hoje.getHours(), hoje.getMinutes());
    if (isNaN(d.getTime())) { return null; }
    return { data: d, resto: texto.slice(m[0].length).trim() };
  }

  var RE_STATUS_FINAL = new RegExp('[\\s,;.–—-]*\\b(?:status[\\s:]*)?' +
    '(resolvid[oa]|conclu[ií]d[oa]|finalizad[oa]|encerrad[oa]|' +
    'em andamento|pendente)\\b[\\s.!]*$', 'i');

  function extrairStatusFinal(texto) {
    var m = RE_STATUS_FINAL.exec(texto);
    if (!m) { return null; }
    var resolvido = /resolvid|conclu|finaliz|encerrad/i.test(m[1]);
    return {
      status: resolvido ? CFG.statusResolvido : CFG.statusAndamento,
      resto: texto.slice(0, m.index).trim()
    };
  }

  function interpretarAvulsa(bruto) {
    var texto = String(bruto || '').replace(/\s+/g, ' ').trim();
    if (!texto) { return null; }

    /* Data e status que VOCE escreveu ganham de qualquer deducao. */
    var quando = null;
    var dt = extrairData(texto);
    if (dt) { quando = dt.data; texto = dt.resto; }

    var statusExplicito = null;
    var st = extrairStatusFinal(texto);
    if (st) { statusExplicito = st.status; texto = st.resto; }

    var limpo = texto;

    if (texto.indexOf('|') >= 0) {           /* separador explicito tem prioridade */
      var p = texto.split('|');
      return {
        assunto: p[0].trim(),
        feito: (p[1] || '').trim(),
        status: (p.slice(2).join('|') || '').trim() || statusExplicito || CFG.statusAndamento,
        quando: quando, statusExplicito: statusExplicito, limpo: limpo
      };
    }

    var feito = '';
    var reAcao = new RegExp('(?:^|[.;]\\s+)((?:' + VERBOS_ACAO + ')\\b[\\s\\S]*)$', 'i');
    var mAcao = reAcao.exec(texto);
    if (!mAcao) {
      mAcao = new RegExp('[,.;]\\s*((?:' + VERBOS_ACAO_IMPESSOAL + ')\\b[\\s\\S]*)$', 'i')
        .exec(texto);
    }
    if (mAcao) {
      feito = mAcao[1].trim();
      texto = texto.slice(0, mAcao.index).trim();
    }

    var reRelato = new RegExp('^\\s*[A-ZÀ-Ý][\\wA-Za-zÀ-ÿ\'.-]*' +
      '(?:\\s+[A-ZÀ-Ý][\\wA-Za-zÀ-ÿ\'.-]*)?\\s+(?:' + VERBOS_RELATO +
      ')\\s+(?:que\\s+|sobre\\s+|d[eoa]\\s+)?', 'i');
    var problema = texto.replace(reRelato, '');

    var uf = '';
    var mUf = /\b([A-Z]{2})\b/.exec(problema);
    if (mUf && UFS.indexOf(mUf[1]) >= 0) {
      uf = mUf[1];
      problema = problema
        .replace(new RegExp('\\s*\\b(?:n[ao]|em|d[ao])\\s+' + uf + '\\b', 'g'), '')
        .replace(new RegExp('\\s*\\b' + uf + '\\b', 'g'), '');
    }

    problema = problema.replace(/\s{2,}/g, ' ')
                       .replace(/^[\s,;.-]+/, '').replace(/[\s,;.]+$/, '').trim();
    if (problema) { problema = problema.charAt(0).toUpperCase() + problema.slice(1); }
    if (uf && problema) { problema = uf + ' - ' + problema; }
    if (feito) { feito = feito.charAt(0).toUpperCase() + feito.slice(1); }

    return {
      assunto: problema || texto,
      feito: feito,
      status: statusExplicito || CFG.statusAndamento,
      quando: quando, statusExplicito: statusExplicito, limpo: limpo
    };
  }

  /* ---- estruturacao por IA local (opcional, com queda para as regras) ----
     Usa o mesmo motor do resumo, definido la em cima: prepararLM/sessaoLimpa. */
  var ESQUEMA_AVULSA = {
    type: 'object',
    properties: {
      problema: { type: 'string' },
      feito: { type: 'string' },
      resolvido: { type: 'boolean' }
    },
    required: ['problema', 'feito', 'resolvido'],
    additionalProperties: false
  };

  /* Conta palavras de conteudo (>= 4 letras). Serve para detectar quando a IA
     resumiu ALEM da conta e jogou fora metade da anotacao. */
  function palavrasDeConteudo(texto) {
    return normalizar(texto).split(' ').filter(function (p) { return p.length >= 4; });
  }

  function perdeuConteudo(original, problema, feito) {
    var orig = palavrasDeConteudo(original);
    if (orig.length < 4) { return false; }
    var saida = ' ' + normalizar(problema + ' ' + feito) + ' ';
    var mantidas = orig.filter(function (p) { return saida.indexOf(' ' + p + ' ') >= 0; });
    return (mantidas.length / orig.length) < 0.5;
  }

  function interpretarAvulsaIA(texto) {
    /* Anotacao curta ja e o titulo: nao ha o que estruturar, e o modelo so
       teria a chance de encurtar demais. As regras resolvem melhor. */
    if (!texto || texto.length < CFG.minCaracteresAvulsaIA) { return Promise.resolve(null); }
    return sessaoLimpa(false).then(function (s) {
      if (!s) { return null; }
      function encerrar() {
        try { if (s !== sessaoLM) { s.destroy(); } } catch (e) {}
      }
      return s.prompt(
        'Organize a anotacao abaixo em tres campos, para um registro de atividade.\n' +
        REGRAS_ESCRITA +
        '- problema: titulo curto da demanda, ate 60 caracteres.\n' +
        '- feito: frase completa com a acao ou o andamento, usando as palavras da\n' +
        '  propria anotacao. NUNCA deixe vazio: se nada foi concluido, descreva o\n' +
        '  que esta em curso, sem acrescentar fato novo.\n' +
        '- resolvido: true somente se a anotacao indicar conclusao.\n' +
        '- Juntos, os dois campos de texto precisam conter tudo o que a anotacao\n' +
        '  diz. Nada de informacao pode ficar de fora.\n\n' +
        'ANOTACAO:\n' + texto,
        { responseConstraint: ESQUEMA_AVULSA }
      ).then(function (r) {
        var o = JSON.parse(r);
        encerrar();
        if (!o || typeof o.problema !== 'string' || !o.problema.trim()) { return null; }
        /* Traduziu sem pedir? Descarta e deixa as regras cuidarem. */
        if (!pareceIngles(texto) &&
            (pareceIngles(o.problema) || pareceIngles(o.feito || ''))) {
          registrar('AVISO: a IA respondeu em ingles na avulsa. Usando as regras.');
          return null;
        }
        var problema = limparSaidaIA(o.problema);
        var feito = limparSaidaIA(o.feito || '');
        if (perdeuConteudo(texto, problema, feito)) {
          registrar('AVISO: a IA resumiu demais a avulsa e perdeu parte do texto. ' +
                    'Usando a sua anotacao como ela veio.');
          return null;
        }
        return {
          assunto: problema.slice(0, 150),
          feito: feito,
          status: o.resolvido ? CFG.statusResolvido : CFG.statusAndamento
        };
      })['catch'](function () { encerrar(); return null; });
    })['catch'](function () { return null; });
  }

  var ESQUEMA_TITULO = {
    type: 'object',
    properties: { titulo: { type: 'string' } },
    required: ['titulo'],
    additionalProperties: false
  };

  /* So o titulo. O texto integral ja esta guardado em "O que foi feito", entao
     aqui o modelo nao tem como perder informacao: no pior caso o titulo sai
     ruim, e o conteudo continua inteiro na linha. */
  function tituloCurtoIA(texto) {
    return sessaoLimpa(false).then(function (s) {
      if (!s) { return null; }
      function encerrar() {
        try { if (s !== sessaoLM) { s.destroy(); } } catch (e) {}
      }
      return s.prompt(
        'Escreva um titulo curto, de ate 60 caracteres, para a anotacao abaixo.\n' +
        REGRAS_ESCRITA +
        '- E um titulo, nao uma frase: sem ponto final.\n' +
        '- Use as palavras da propria anotacao.\n\n' +
        'ANOTACAO:\n' + texto,
        { responseConstraint: ESQUEMA_TITULO }
      ).then(function (r) {
        var o = JSON.parse(r);
        encerrar();
        var t = limparSaidaIA(o && o.titulo);
        if (!t) { return null; }
        if (!pareceIngles(texto) && pareceIngles(t)) {
          registrar('AVISO: a IA titulou em ingles. Mantida a sua anotacao.');
          return null;
        }
        return t.replace(/[.;,]+$/, '').slice(0, 90);
      })['catch'](function () { encerrar(); return null; });
    })['catch'](function () { return null; });
  }

  /* ================================================= API DO ZENDESK ====== */
  function api(caminho) {
    return fetch(caminho, {
      credentials: 'same-origin',
      headers: { 'Accept': 'application/json' }
    }).then(function (r) {
      if (r.status === 401 || r.status === 403) {
        return r.text().then(function (corpo) {
          var detalhe = '';
          try {
            var j = JSON.parse(corpo);
            var d = j.description || j.error || j.errors || j.message || '';
            detalhe = (typeof d === 'string') ? d : JSON.stringify(d);
          } catch (e) { detalhe = corpo.slice(0, 200); }
          throw new Error('HTTP ' + r.status + ' em ' + caminho +
            (detalhe ? ' — ' + detalhe : '') + '. Confira se voce esta logado como agente.');
        });
      }
      if (!r.ok) { throw new Error('HTTP ' + r.status + ' em ' + caminho); }
      return r.json();
    });
  }

  var meuId = null;
  var statusPersonalizados = null;

  /* Se o lote (show_many) for negado - por exemplo, um unico ticket do lote
     saiu do seu acesso - busca um a um em vez de travar a rodada inteira.
     O(s) ticket(s) negado(s) fica(m) de fora desta rodada (e sao avisados no
     painel), mas os demais continuam sendo capturados normalmente. */
  /* Busca o lote e, se ele falhar por QUALQUER motivo, parte no meio e tenta
     as duas metades. Repete ate sobrar um ticket so - e ai o problema esta
     identificado. Cobre os dois casos reais: um ticket fora do seu acesso
     (403) e um lote grande demais para o servidor (500). Antes, um 500 unico
     derrubava a rodada inteira e nenhum ticket era capturado. */
  function buscarTickets(ids) {
    return api('/api/v2/tickets/show_many.json?ids=' + ids.join(','))
      .then(function (d) { return { tickets: (d && d.tickets) || [], negados: [] }; })
      ['catch'](function (e) {
        if (ids.length === 1) {
          marcarTicketNegado(ids[0], e.message);
          return { tickets: [], negados: ids };
        }
        var meio = Math.ceil(ids.length / 2);
        return buscarTickets(ids.slice(0, meio)).then(function (a) {
          return buscarTickets(ids.slice(meio)).then(function (b) {
            return { tickets: a.tickets.concat(b.tickets),
                     negados: a.negados.concat(b.negados) };
          });
        });
      });
  }

  /* Registra que este ticket negou acesso e poe ele "em espera": nas proximas
     rodadas ele nem entra mais no lote (ticketsRelevantes filtra), entao para
     de gerar aviso repetido. So avisa uma vez, na primeira negativa. Reabrir o
     ticket no Zendesk (anotarTicketVisitado) cancela a espera. */
  function marcarTicketNegado(id, motivo) {
    var mapa = ler(CH.tickets, {});
    var m = mapa[id];
    if (!m) { return; }
    var jaEmEspera = m.negadoAte && m.negadoAte > Date.now();
    m.negadoAte = Date.now() + CFG.horasSuspensaoTicketNegado * 3600000;
    gravar(CH.tickets, mapa);
    if (!jaEmEspera) {
      registrar('AVISO: ticket ' + id + ' sem acesso (' + motivo + '); vou parar de tentar por ' +
                CFG.horasSuspensaoTicketNegado + 'h ou ate voce abrir esse ticket de novo.');
    }
  }

  /* Mapa  custom_status_id -> rotulo que o agente ve na tela. */
  function carregarStatusPersonalizados() {
    if (statusPersonalizados) { return Promise.resolve(statusPersonalizados); }
    return api('/api/v2/custom_statuses.json').then(function (d) {
      var mapa = {};
      ((d && d.custom_statuses) || []).forEach(function (s) {
        if (s && s.id) { mapa[String(s.id)] = s.agent_label || s.end_user_label || ''; }
      });
      statusPersonalizados = mapa;
      registrar('Status da interface carregados: ' + Object.keys(mapa).length + '.');
      return mapa;
    })['catch'](function (e) {
      statusPersonalizados = {};
      registrar('AVISO: nao consegui ler os status personalizados (' + e.message +
                '). Usando o status padrao do Zendesk.');
      return statusPersonalizados;
    });
  }

  function descobrirUsuario() {
    if (meuId) { return Promise.resolve(meuId); }
    return api('/api/v2/users/me.json').then(function (d) {
      meuId = d && d.user && d.user.id;
      if (!meuId) { throw new Error('Nao foi possivel identificar seu usuario no Zendesk.'); }
      /* O nome serve so para uma coisa: reconhecer "Fulano informou que..." na
         saida da IA e tirar esse sujeito do texto. Fica na memoria da aba. */
      meuNome = (d.user && d.user.name) || '';
      registrar('Conectado como usuario ' + meuId + '.');
      return meuId;
    });
  }

  /* ------------------------------------------- tickets que voce visitou */
  function anotarTicketVisitado(id) {
    if (!id) { return; }
    var mapa = ler(CH.tickets, {});
    if (!mapa[id]) { mapa[id] = { visto: Date.now(), updated_at: null, ultimoAudit: 0 }; }
    else { mapa[id].visto = Date.now(); mapa[id].negadoAte = 0; }
    gravar(CH.tickets, mapa);
  }

  function ticketDaUrl() {
    var m = /\/agent\/tickets\/(\d+)/.exec(location.pathname);
    return m ? m[1] : null;
  }

  function vigiarNavegacao() {
    var ultima = '';
    setInterval(function () {
      if (location.href === ultima) { return; }
      ultima = location.href;
      var id = ticketDaUrl();
      if (id) { anotarTicketVisitado(id); }
    }, 1000);
    var id0 = ticketDaUrl();
    if (id0) { anotarTicketVisitado(id0); }
  }

  /* Pergunta ao Zendesk quais tickets tiveram a sua mao nesta semana e poe
     todos no mapa de rastreados. So leitura (GET), com a sua propria sessao. */
  var buscaFalhou = false;

  function descobrirTicketsDaSemana() {
    if (!CFG.descobrirPelaBusca) { return Promise.resolve(0); }
    var inicio = semanaDe(agora()).inicio;
    var achados = {};

    function pagina(url, restantes) {
      return api(url).then(function (d) {
        ((d && d.results) || []).forEach(function (t) {
          if (t && t.id && (!t.result_type || t.result_type === 'ticket')) {
            achados[String(t.id)] = t.updated_at || null;
          }
        });
        if (d && d.next_page && restantes > 1) {
          return pagina(d.next_page.replace(/^https?:\/\/[^\/]+/, ''), restantes - 1);
        }
        return null;
      });
    }

    return CFG.consultasDeBusca.reduce(function (p, consulta) {
      return p.then(function () {
        return pagina('/api/v2/search.json?per_page=100&sort_by=updated_at&sort_order=desc&query=' +
                      encodeURIComponent(consulta.replace('{inicio}', inicio)), CFG.maxPaginasBusca);
      });
    }, Promise.resolve()).then(function () {
      var mapa = ler(CH.tickets, {});
      var novos = 0;
      Object.keys(achados).forEach(function (id) {
        var quando = achados[id] ? new Date(achados[id]).getTime() : Date.now();
        if (!mapa[id]) {
          mapa[id] = { visto: quando || Date.now(), updated_at: null, ultimoAudit: 0 };
          novos++;
        } else if (quando && quando > (mapa[id].visto || 0)) {
          mapa[id].visto = quando;
        }
      });
      gravar(CH.tickets, mapa);
      if (buscaFalhou) { registrar('Busca do Zendesk voltou a responder.'); }
      buscaFalhou = false;
      if (novos) { registrar('Busca: ' + novos + ' ticket(s) seu(s) da semana passaram a ser acompanhados.'); }
      return novos;
    })['catch'](function (e) {
      if (!buscaFalhou) {
        registrar('AVISO: a busca do Zendesk nao respondeu (' + e.message + '). Seguindo so ' +
                  'com os tickets abertos nesta aba.');
      }
      buscaFalhou = true;
      return 0;
    });
  }

  /* O show_many aceita no maximo 100 ids por chamada. */
  function buscarTicketsEmLotes(ids) {
    var lotes = [];
    for (var i = 0; i < ids.length; i += 100) { lotes.push(ids.slice(i, i + 100)); }
    return lotes.reduce(function (p, lote) {
      return p.then(function (acc) {
        return buscarTickets(lote).then(function (r) {
          return { tickets: acc.tickets.concat(r.tickets), negados: acc.negados.concat(r.negados) };
        });
      });
    }, Promise.resolve({ tickets: [], negados: [] }));
  }

  function ticketsRelevantes() {
    var mapa = ler(CH.tickets, {});
    var limite = Date.now() - CFG.diasDeTicketsRastreados * 86400000;
    var agoraMs = Date.now();
    var ids = Object.keys(mapa).filter(function (id) {
      var m = mapa[id];
      if (m.visto < limite) { return false; }
      if (m.negadoAte && m.negadoAte > agoraMs) { return false; }   /* em espera */
      return true;
    });
    ids.sort(function (a, b) { return mapa[b].visto - mapa[a].visto; });
    return ids.slice(0, CFG.maxTicketsPorLote);
  }

  /* --------------------------------------------------- varredura -------- */
  var varrendo = false;

  function varrer() {
    if (varrendo) { return Promise.resolve(0); }
    varrendo = true;
    atualizarPainel();

    return descobrirUsuario().then(function () {
      return CFG.usarStatusPersonalizado ? carregarStatusPersonalizados() : null;
    }).then(function () {
      return descobrirTicketsDaSemana();
    }).then(function () {
      var ids = ticketsRelevantes();
      if (!ids.length) { return []; }
      return buscarTicketsEmLotes(ids).then(function (r) {
        if (r.negados.length) {
          registrar('AVISO: ' + r.negados.length + ' de ' + ids.length +
                    ' ticket(s) nao puderam ser lidos; os outros ' + r.tickets.length +
                    ' foram capturados normalmente.');
        }
        return r.tickets;
      });
    }).then(function (tickets) {
      var mapa = ler(CH.tickets, {});
      var mudaram = tickets.filter(function (t) {
        var m = mapa[String(t.id)];
        return !m || m.updated_at !== t.updated_at;
      });
      if (!mudaram.length) {
        registrar('Varredura: nenhum ticket mudou.');
        return 0;
      }
      registrar('Varredura: ' + mudaram.length + ' ticket(s) com mudanca.');

      var novas = 0;
      var fila = mudaram.reduce(function (p, t) {
        return p.then(function () {
          return coletarAuditorias(t).then(function (n) { novas += n; });
        });
      }, Promise.resolve());

      return fila.then(function () {
        var m2 = ler(CH.tickets, {});
        mudaram.forEach(function (t) {
          if (m2[String(t.id)]) { m2[String(t.id)].updated_at = t.updated_at; }
        });
        gravar(CH.tickets, m2);
        return novas;
      });
    }).then(function (novas) {
      var est = ler(CH.estado, {});
      est.ultimaVarredura = Date.now();
      gravar(CH.estado, est);
      varrendo = false;
      atualizarPainel();
      return novas;
    })['catch'](function (e) {
      varrendo = false;
      registrar('ERRO na varredura: ' + e.message);
      atualizarPainel();
      return 0;
    });
  }

  function coletarAuditorias(ticket) {
    var mapa = ler(CH.tickets, {});
    var info = mapa[String(ticket.id)] || { ultimoAudit: 0 };
    var limiteSemana = semanaDe(agora()).inicio;
    var coletadas = [];

    function pagina(url, restantes) {
      return api(url).then(function (d) {
        var audits = (d && d.audits) || [];
        audits.forEach(function (a) { coletadas.push(a); });
        if (d && d.next_page && restantes > 1) {
          return pagina(d.next_page.replace(/^https?:\/\/[^\/]+/, ''), restantes - 1);
        }
        return null;
      });
    }

    return pagina('/api/v2/tickets/' + ticket.id + '/audits.json', CFG.maxPaginasAuditoria)
      .then(function () {
        var novas = 0;
        var maiorId = info.ultimoAudit || 0;

        var fila = Promise.resolve();
        coletadas.forEach(function (a) {
          if (a.id > maiorId) { maiorId = a.id; }
          if (a.id <= (info.ultimoAudit || 0)) { return; }
          if (a.author_id !== meuId) { return; }
          if (diaISO(new Date(a.created_at)) < limiteSemana) { return; }
          fila = fila.then(function () {
            return registrarAtividadeDeAudit(a, ticket).then(function (ok) {
              if (ok) { novas++; }
            });
          });
        });

        return fila.then(function () {
          var m = ler(CH.tickets, {});
          if (!m[String(ticket.id)]) { m[String(ticket.id)] = { visto: Date.now() }; }
          m[String(ticket.id)].ultimoAudit = maiorId;
          gravar(CH.tickets, m);
          return novas;
        });
      });
  }

  function registrarAtividadeDeAudit(audit, ticket) {
    var comentario = '', statusNovo = '';
    (audit.events || []).forEach(function (ev) {
      if (!ev) { return; }
      if (!comentario && (ev.type === 'Comment' || ev.type === 'VoiceComment') &&
          typeof ev.body === 'string') {
        comentario = ev.body;
      }
      if (ev.type === 'Change' && ev.value && ev.field_name === 'status') {
        statusNovo = String(ev.value);
      }
    });

    if (!comentario && !statusNovo) { return Promise.resolve(false); }

    var id = 'zd-' + audit.id;
    if (ler(CH.atividades, {})[id]) { return Promise.resolve(false); }

    /* A planilha so tem dois estados. Encerrado no Zendesk = Resolvido. */
    var base = String(statusNovo || ticket.status || '').toLowerCase();
    var rotuloStatus = (CFG.statusBaseResolvido.indexOf(base) >= 0)
      ? CFG.statusResolvido : CFG.statusAndamento;

    return prepararTexto(comentario).then(function (texto) {
      var todas = ler(CH.atividades, {});
      if (todas[id]) { return false; }      /* outra rodada chegou antes */
      todas[id] = {
        quando: audit.created_at,
        ticket: String(ticket.id),
        assunto: ticket.subject || ticket.raw_subject || '',
        feito: texto,
        status: rotuloStatus,
        origem: 'zendesk',
        gravado: false
      };
      gravar(CH.atividades, todas);
      registrar('Capturado: ticket ' + ticket.id + ' -> ' + rotuloStatus +
                (texto ? ' | "' + texto.slice(0, 38) +
                         (texto.length > 38 ? '...' : '') + '"' : ' (sem texto)'));
      return true;
    });
  }

  /* ============================================== ATIVIDADE AVULSA ======= */
  function adicionarManual(texto) {
    var base = interpretarAvulsa(texto);
    if (!base || !base.assunto) { return Promise.resolve(false); }

    return interpretarAvulsaIA(base.limpo).then(function (ia) {
      var r = (ia && ia.assunto)
        ? {
            assunto: ia.assunto,
            feito: ia.feito,
            /* o que voce escreveu explicitamente ganha do que a IA deduziu */
            status: base.statusExplicito || ia.status,
            quando: base.quando
          }
        : { assunto: base.assunto, feito: base.feito, status: base.status,
            quando: base.quando };

      var como = ia ? 'IA' : 'regra';
      if (!CFG.completarCampoFeito || String(r.feito || '').trim()) {
        return { r: r, como: como };
      }

      /* Sem acao descrita, a anotacao inteira E a atividade: ela vai para
         "O que foi feito" com as suas palavras - nada e inventado. Se o titulo
         ficaria identico ao texto, a IA escreve so um titulo curto. Nesse
         pedido ela nao carrega conteudo nenhum, entao nao tem como perder
         informacao: o texto integral ja esta guardado na outra coluna. */
      var nota = String(base.limpo || '').trim();
      r.feito = nota.charAt(0).toUpperCase() + nota.slice(1);

      if (normalizar(r.assunto) !== normalizar(nota)) { return { r: r, como: como }; }
      return tituloCurtoIA(nota).then(function (titulo) {
        if (titulo) { r.assunto = titulo; }
        return { r: r, como: titulo ? 'IA titulo' : 'regra' };
      });
    }).then(function (res) {
      var r = res.r;
      var atividades = ler(CH.atividades, {});
      var id = 'man-' + Date.now().toString(36);
      atividades[id] = {
        quando: (r.quando || new Date()).toISOString(),
        ticket: CFG.ticketAvulso,
        assunto: r.assunto,
        feito: r.feito,
        status: r.status,
        origem: 'manual',
        gravado: false
      };
      gravar(CH.atividades, atividades);
      registrar('Avulsa (' + res.como + '): "' + r.assunto.slice(0, 38) + '" -> ' + r.status +
                (base.quando ? ' | data ' + diaISO(base.quando) : ''));
      return true;
    })['catch'](function (e) {
      registrar('ERRO ao interpretar a avulsa: ' + String((e && e.message) || e));
      return false;
    });
  }

  /* ================================================= ACESSO AO ARQUIVO === */
  var dirHandle = null;

  function bdHandles() {
    return new Promise(function (ok, erro) {
      var req = indexedDB.open('aa2_fs', 1);
      req.onupgradeneeded = function () { req.result.createObjectStore('h'); };
      req.onsuccess = function () { ok(req.result); };
      req.onerror = function () { erro(req.error); };
    });
  }

  function guardarHandle(h) {
    return bdHandles().then(function (db) {
      return new Promise(function (ok, erro) {
        var tx = db.transaction('h', 'readwrite');
        tx.objectStore('h').put(h, 'pasta');
        tx.oncomplete = function () { ok(true); };
        tx.onerror = function () { erro(tx.error); };
      });
    });
  }

  function recuperarHandle() {
    return bdHandles().then(function (db) {
      return new Promise(function (ok) {
        var tx = db.transaction('h', 'readonly');
        var r = tx.objectStore('h').get('pasta');
        r.onsuccess = function () { ok(r.result || null); };
        r.onerror = function () { ok(null); };
      });
    })['catch'](function () { return null; });
  }

  function temPermissao(h, pedir) {
    var opt = { mode: 'readwrite' };
    return h.queryPermission(opt).then(function (p) {
      if (p === 'granted') { return true; }
      if (!pedir) { return false; }
      return h.requestPermission(opt).then(function (p2) { return p2 === 'granted'; });
    });
  }

  /* Nome REAL da planilha nesta pasta. Pode nao ser o CFG.nomeArquivo: ver a
     nota na configuracao. Resolvido uma vez por conexao. */
  var nomeReal = null;

  function planilhaDaPasta() { return nomeReal || CFG.nomeArquivo; }

  /* Acha a planilha: o nome preferido primeiro; nao existindo, o unico .xlsx
     solto na pasta. Ignora os temporarios do Excel (~$) e as subpastas, entao
     backups/ e arquivos_semanais/ nunca entram na conta. */
  function localizarPlanilha(h) {
    return h.getFileHandle(CFG.nomeArquivo)
      .then(function () { return CFG.nomeArquivo; })
      ['catch'](function () {
        var achados = [];
        var it = h.entries();
        function proximo() {
          return it.next().then(function (r) {
            if (r.done) { return null; }
            var nome = r.value[0], item = r.value[1];
            if (item && item.kind === 'file' && /\.xlsx$/i.test(nome) &&
                nome.indexOf('~$') !== 0) {
              achados.push(nome);
            }
            return proximo();
          });
        }
        return proximo().then(function () {
          if (achados.length === 1) { return achados[0]; }
          if (achados.length === 0) {
            throw new Error('a pasta escolhida nao tem nenhum arquivo .xlsx');
          }
          throw new Error('a pasta tem ' + achados.length + ' planilhas (' +
            achados.slice(0, 4).join(', ') + '). Deixe so a do controle na pasta, ' +
            'ou ajuste "nomeArquivo" no topo do atividades_zendesk.js');
        });
      });
  }

  function escolherPasta() {
    if (!window.showDirectoryPicker) {
      registrar('ERRO: este Chrome nao oferece acesso a pastas (File System Access API).');
      return Promise.resolve(false);
    }
    /* startIn/id: abre direto em Documentos e lembra o ultimo lugar escolhido.
       Lembrete: e um seletor de PASTA - arquivos .xlsx nao aparecem nele. */
    return window.showDirectoryPicker({
      mode: 'readwrite', id: 'atividades2', startIn: 'documents'
    }).then(function (h) {
      return localizarPlanilha(h).then(function (nome) {
        dirHandle = h;
        nomeReal = nome;
        return guardarHandle(h).then(function () {
          registrar('Pasta conectada. Planilha: "' + nome + '".');
          atualizarPainel();
          return true;
        });
      })['catch'](function (e) {
        registrar('ERRO ao conectar: ' + String((e && e.message) || e) + '.');
        return false;
      });
    })['catch'](function (e) {
      registrar('Selecao de pasta cancelada.' + (e && e.message ? ' (' + e.message + ')' : ''));
      return false;
    });
  }

  /* Permissao concedida NAO significa pasta utilizavel: se a pasta foi renomeada
     ou movida, o Chrome continua dizendo "granted" para um lugar que nao existe
     mais. Por isso aqui a gente confirma que a planilha esta mesmo la. */
  function conectarPasta(pedirPermissao) {
    function validar(h) {
      return temPermissao(h, pedirPermissao).then(function (ok) {
        if (!ok) { dirHandle = null; return false; }
        return localizarPlanilha(h).then(function (nome) {
          dirHandle = h;
          nomeReal = nome;
          return true;
        })['catch'](function () {
          dirHandle = null;
          nomeReal = null;
          return false;
        });
      });
    }
    if (dirHandle) { return validar(dirHandle); }
    return recuperarHandle().then(function (h) {
      return h ? validar(h) : false;
    });
  }

  /* ============================================ MAPEAMENTO DA PLANILHA === */
  var APELIDOS = {
    data: ['data', 'dia', 'data atendimento', 'data da atividade'],
    ticket: ['n ticket', 'no ticket', 'no do ticket', 'numero do ticket', 'numero ticket',
             'ticket', 'chamado', 'n chamado', 'no chamado', 'id do ticket'],
    problema: ['problema solicitacao', 'problema ou solicitacao', 'problema', 'solicitacao',
               'assunto', 'descricao', 'demanda'],
    feito: ['o que foi feito', 'oque foi feito', 'que foi feito', 'o que foi realizado',
            'atividade realizada', 'acao realizada', 'tratativa', 'acao'],
    status: ['status andamento', 'status e andamento', 'status', 'andamento', 'situacao',
             'status do ticket']
  };
  var CAMPOS = ['data', 'ticket', 'problema', 'feito', 'status'];

  function classificar(texto) {
    var alvo = normalizar(texto);
    if (!alvo) { return []; }
    var achados = [];
    Object.keys(APELIDOS).forEach(function (campo) {
      var melhor = 0;
      APELIDOS[campo].forEach(function (ap) {
        if (alvo === ap) { melhor = Math.max(melhor, 3); }
        else if (ap.length >= 5 && new RegExp('(^| )' + ap + '( |$)').test(alvo)) {
          melhor = Math.max(melhor, 1);
        }
      });
      if (melhor) { achados.push({ campo: campo, pontos: melhor }); }
    });
    return achados;
  }

  function mapaDaAbaAberta(pasta) {
    var melhorLinha = 0, melhorCols = null, melhorPts = -1, melhorTextos = null;
    for (var linha = 1; linha <= 30; linha++) {
      var cols = {}, pts = 0, textos = {}, achou = 0;
      for (var col = 1; col <= 40; col++) {
        var v = pasta.ler(linha, col);
        if (v === null || String(v).trim() === '') { continue; }
        textos[col] = String(v).trim();
        classificar(v).forEach(function (a) {
          if (!cols[a.campo] || a.pontos > cols[a.campo].pontos) {
            cols[a.campo] = { col: col, pontos: a.pontos };
          }
        });
      }
      achou = Object.keys(cols).length;
      pts = Object.keys(cols).reduce(function (s, k) { return s + cols[k].pontos; }, 0) + achou;
      if (!melhorCols || achou > Object.keys(melhorCols).length ||
          (achou === Object.keys(melhorCols).length && pts > melhorPts)) {
        melhorLinha = linha; melhorCols = cols; melhorPts = pts; melhorTextos = textos;
      }
    }
    if (!melhorCols) { return null; }
    var faltando = CAMPOS.filter(function (c) { return !melhorCols[c]; });
    var simples = {};
    Object.keys(melhorCols).forEach(function (k) { simples[k] = melhorCols[k].col; });
    return { linha: melhorLinha, colunas: simples, faltando: faltando, textos: melhorTextos };
  }

  function escolherAbaEMapa(pasta) {
    var visiveis = pasta.abas.filter(function (a) { return a.estado === 'visible'; });
    if (!visiveis.length) { visiveis = pasta.abas; }
    var melhor = null;

    return visiveis.reduce(function (p, aba) {
      return p.then(function () {
        return pasta.abrirAba(aba.nome).then(function () {
          var m = mapaDaAbaAberta(pasta);
          if (!m) { return; }
          var completo = m.faltando.length === 0;
          if (!melhor || (completo && melhor.faltando.length) ||
              (completo === (melhor.faltando.length === 0) &&
               Object.keys(m.colunas).length > Object.keys(melhor.colunas).length)) {
            melhor = { aba: aba.nome, linha: m.linha, colunas: m.colunas,
                       faltando: m.faltando, textos: m.textos };
          }
        });
      });
    }, Promise.resolve()).then(function () {
      if (!melhor) {
        throw new Error('Nenhum cabecalho reconhecido na planilha.');
      }
      if (melhor.faltando.length) {
        throw new Error('Nao consegui identificar a(s) coluna(s): ' + melhor.faltando.join(', ') +
          '. Cabecalhos lidos na aba "' + melhor.aba + '" (linha ' + melhor.linha + '): ' +
          Object.keys(melhor.textos).map(function (c) {
            return XLSX.letraColuna(parseInt(c, 10)) + '=' + melhor.textos[c];
          }).join(' | '));
      }
      return pasta.abrirAba(melhor.aba).then(function () { return melhor; });
    });
  }

  function ultimaLinhaPreenchida(pasta, mapa) {
    var maior = mapa.linha;
    var rows = pasta.linhas();
    var limite = 0, i;
    for (i = 0; i < rows.length; i++) {
      var n = parseInt(rows[i].getAttribute('r'), 10);
      if (n > limite) { limite = n; }
    }
    for (var linha = mapa.linha + 1; linha <= limite; linha++) {
      for (var k = 0; k < CAMPOS.length; k++) {
        var v = pasta.ler(linha, mapa.colunas[CAMPOS[k]]);
        if (v !== null && String(v).trim() !== '') { maior = linha; break; }
      }
    }
    return maior;
  }

  /* ================================================ GRAVACAO NO EXCEL ==== */
  function pendentes() {
    var a = ler(CH.atividades, {});
    return Object.keys(a).filter(function (id) { return !a[id].gravado; })
      .sort(function (x, y) { return new Date(a[x].quando) - new Date(a[y].quando); })
      .map(function (id) { return { id: id, d: a[id] }; });
  }

  /* ------------------------- historico da demanda ------------------------
     A celula "O Que Foi Feito" de uma demanda com varias tratativas tem este
     formato, uma tratativa por linha:

        10/11 - Fornecedor esteve na unidade; chuva impediu o servico.
        11/11 - Manutencao realizada com sucesso, validacao enviada.

     Demanda com UMA tratativa fica so com o texto, sem data na frente (a data
     ja esta na coluna Data). A data so aparece quando chega a segunda. */
  var RE_ENTRADA = /^\s*(\d{1,2})\/(\d{1,2})(?:\/\d{2,4})?(?:\s+\d{1,2}:\d{2})?\s*[-–—]\s+(.*)$/;

  function diaMes(iso) {
    return iso ? (iso.slice(8, 10) + '/' + iso.slice(5, 7)) : '';
  }

  function lerHistorico(texto, diaPadraoISO) {
    var entradas = [];
    String(texto === null || texto === undefined ? '' : texto)
      .replace(/\r\n?/g, '\n').split('\n').forEach(function (l) {
        var s = l.trim();
        if (!s) { return; }
        var m = RE_ENTRADA.exec(s);
        if (m) {
          entradas.push({ dia: pad(parseInt(m[1], 10)) + '/' + pad(parseInt(m[2], 10)),
                          texto: m[3].trim() });
        } else if (entradas.length) {
          entradas[entradas.length - 1].texto += ' ' + s;      /* continuacao */
        } else {
          entradas.push({ dia: diaMes(diaPadraoISO), texto: s });
        }
      });
    return entradas;
  }

  function escreverHistorico(entradas) {
    if (entradas.length === 1) { return entradas[0].texto; }
    var t = entradas.map(function (e) {
      return (e.dia ? e.dia + ' - ' : '') + e.texto;
    }).join('\n');
    if (CFG.maxCaracteresCelula > 0 && t.length > CFG.maxCaracteresCelula) {
      t = t.slice(0, CFG.maxCaracteresCelula).replace(/\s+\S*$/, '') + '...';
    }
    return t;
  }

  /* Ja esta no historico? Compara sem acento, caixa e pontuacao. */
  function jaConsta(entradas, texto) {
    var alvo = normalizar(texto);
    if (!alvo) { return true; }
    return entradas.some(function (e) {
      var n = normalizar(e.texto);
      return n === alvo || (alvo.length >= 25 && n.indexOf(alvo) >= 0);
    });
  }

  /* Deixa o Excel recalcular a altura da linha ao abrir: sem isto a linha
     continua com a altura de duas linhas de texto e o historico fica cortado. */
  function alturaAutomatica(pasta, linha) {
    try {
      var row = pasta.acharLinha(linha);
      if (row && row.getAttribute('customHeight') !== '1') { row.removeAttribute('ht'); }
    } catch (e) { /* cosmetico: nunca impede a gravacao */ }
  }

  /* Trava entre abas: se o Zendesk estiver aberto em duas abas e as duas
     estiverem armadas, elas nao podem ler-modificar-gravar o mesmo arquivo ao
     mesmo tempo (a segunda gravacao apagaria a linha da primeira). */
  var ID_ABA = Math.random().toString(36).slice(2);

  /* So UMA aba faz a varredura periodica. Se voce armar o painel em duas, a
     segunda fica de reserva e assume sozinha se a primeira for fechada. */
  function souLider() {
    var l = ler('aa2_lider', null);
    if (l && l.dono !== ID_ABA && (Date.now() - l.quando) < 90000) { return false; }
    gravar('aa2_lider', { dono: ID_ABA, quando: Date.now() });
    return true;
  }

  function travar() {
    var t = ler('aa2_trava', null);
    if (t && t.dono !== ID_ABA && (Date.now() - t.quando) < 60000) { return false; }
    gravar('aa2_trava', { dono: ID_ABA, quando: Date.now() });
    return true;
  }

  function destravar() {
    var t = ler('aa2_trava', null);
    if (!t || t.dono === ID_ABA) { gravar('aa2_trava', null); }
  }

  function gravarNoExcel() {
    var lista = pendentes();
    if (!lista.length) { registrar('Nada pendente para gravar.'); return Promise.resolve(0); }

    if (!travar()) {
      registrar('Outra aba do Zendesk esta gravando agora. Fica para a proxima rodada.');
      return Promise.resolve(0);
    }

    return conectarPasta(true).then(function (ok) {
      if (!ok) {
        registrar('Planilha nao conectada. Clique em "Conectar planilha".');
        return 0;
      }
      return dirHandle.getFileHandle(planilhaDaPasta()).then(function (fh) {
        return fh.getFile().then(function (f) { return f.arrayBuffer(); })
          .then(function (buf) { return aplicar(buf, lista, fh); });
      });
    }).then(function (n) {
      destravar();
      return n;
    })['catch'](function (e) {
      destravar();
      var msg = String((e && e.message) || e);
      if ((e && e.name === 'NotFoundError') || /could not be found|not be found/i.test(msg)) {
        dirHandle = null;
        registrar('A pasta ou a planilha mudou de lugar. Clique em "Conectar planilha" ' +
                  'e escolha a pasta de novo. Nada foi perdido: ' + pendentes().length +
                  ' atividade(s) continuam guardadas.');
      } else {
        registrar('ERRO ao gravar: ' + msg);
      }
      atualizarPainel();
      return 0;
    });
  }

  function aplicar(bufferOriginal, lista, fileHandle) {
    var bytesOriginais = new Uint8Array(bufferOriginal.slice(0));
    var aplicadas = [];

    return XLSX.abrir(bufferOriginal).then(function (pasta) {
      return escolherAbaEMapa(pasta).then(function (mapa) {
        var semana = semanaDe(agora());
        var ultima = ultimaLinhaPreenchida(pasta, mapa);
        var colunas = CAMPOS.map(function (c) { return mapa.colunas[c]; });
        var maxColuna = Math.max.apply(null, colunas);

        lista.forEach(function (item) {
          var d = item.d;
          var quando = new Date(d.quando);
          var linhaDoTicket = localizarLinhaTicket(pasta, mapa, d, semana, ultima) ||
                              localizarLinhaAvulsa(pasta, mapa, d, semana, ultima);

          /* Uma linha por demanda: a tratativa completa a linha que ja existe. */
          if (CFG.consolidarTicketRepetido && linhaDoTicket) {
            var oQue = consolidar(pasta, mapa, linhaDoTicket, d, quando);
            aplicadas.push({ id: item.id, linha: linhaDoTicket, acao: oQue });
            return;
          }

          /* Ticket do Zendesk sem texto = voce so mudou o status. Isso nao
             merece linha nova, mas nao pode sumir: atualiza o status da ultima
             linha daquele ticket na semana.
             ATENCAO: so vale para o Zendesk. Atividade avulsa sem acao descrita
             DEVE virar linha, com o campo em branco. */
          if (!d.feito && d.origem === 'zendesk') {
            if (linhaDoTicket && d.status) {
              pasta.escreverTexto(linhaDoTicket, mapa.colunas.status, d.status);
              aplicadas.push({ id: item.id, linha: linhaDoTicket, acao: 'status' });
            } else {
              aplicadas.push({ id: item.id, linha: null, acao: 'ignorado' });
            }
            return;
          }

          {
            var nova = ultima + 1;
            if (ultima > mapa.linha) {
              pasta.copiarEstiloDaLinha(ultima, nova, range(1, maxColuna));
            }
            pasta.escreverNumero(nova, mapa.colunas.data, pasta.serialData(quando));
            var digitos = soDigitos(d.ticket);
            if (digitos && digitos === String(d.ticket).trim()) {
              pasta.escreverNumero(nova, mapa.colunas.ticket, digitos);
            } else {
              pasta.escreverTexto(nova, mapa.colunas.ticket, d.ticket);
            }
            /* Para o Zendesk, assunto e status sempre existem: se vierem vazios
               ha algo errado e isso precisa ficar visivel. Para atividade avulsa,
               vazio e so um campo que voce nao preencheu - deixar em branco e
               mais limpo do que poluir o relatorio com "Nao identificado". */
            var reserva = (d.origem === 'zendesk') ? CFG.naoIdentificado : '';
            pasta.escreverTexto(nova, mapa.colunas.problema, d.assunto || reserva);
            /* Sem carimbo de hora: cada linha ja tem a sua propria data na
               coluna Data. O carimbo so existe no modo consolidado, onde serve
               para separar uma atualizacao da outra dentro da mesma celula. */
            var cel = pasta.escreverTexto(nova, mapa.colunas.feito, d.feito || '');
            aplicarQuebra(pasta, cel);
            alturaAutomatica(pasta, nova);
            pasta.escreverTexto(nova, mapa.colunas.status, d.status || reserva);
            pasta.estenderValidacoes(ultima, nova);
            ultima = nova;
            aplicadas.push({ id: item.id, linha: nova, acao: 'criado' });
          }
        });

        pasta.ajustarDimensao(ultima, maxColuna);
        return XLSX.gerarBytes(pasta).then(function (bytes) {
          return conferir(bytes, mapa, aplicadas).then(function () {
            return backup(bytesOriginais).then(function () {
              return escrever(fileHandle, bytes);
            });
          });
        });
      });
    }).then(function () {
      var todas = ler(CH.atividades, {});
      var n = 0;
      aplicadas.forEach(function (a) {
        if (todas[a.id]) { todas[a.id].gravado = true; todas[a.id].linha = a.linha; todas[a.id].acao = a.acao; }
        if (a.acao !== 'ignorado') { n++; }
      });
      gravar(CH.atividades, todas);
      var criadas = aplicadas.filter(function (x) { return x.acao === 'criado'; }).length;
      var completadas = aplicadas.filter(function (x) { return x.acao === 'atualizado'; }).length;
      registrar('Planilha atualizada: ' + criadas + ' demanda(s) nova(s), ' +
                completadas + ' completada(s)' +
                (n - criadas - completadas ? ', ' + (n - criadas - completadas) +
                 ' so com status' : '') + '.');
      atualizarPainel();
      return n;
    });
  }

  function range(a, b) {
    var r = [];
    for (var i = a; i <= b; i++) { r.push(i); }
    return r;
  }

  /* Ultima linha daquele ticket dentro da semana atual, ou null. */
  function localizarLinhaTicket(pasta, mapa, d, semana, ultima) {
    if (!soDigitos(d.ticket)) { return null; }   /* '-' nunca casa com nada */
    var achada = null;
    for (var l = mapa.linha + 1; l <= ultima; l++) {
      if (soDigitos(pasta.ler(l, mapa.colunas.ticket)) !== soDigitos(d.ticket)) { continue; }
      var dia = dataDaCelula(pasta.ler(l, mapa.colunas.data));
      if (dia === null || (dia >= semana.inicio && dia <= semana.fim)) { achada = l; }
    }
    return achada;
  }

  /* Avulsa com o MESMO assunto de outra da semana completa aquela linha. E o
     que permite acompanhar uma demanda sem ticket: lance de novo usando
     "Mesmo assunto | nova tratativa" e ela entra no historico. */
  function localizarLinhaAvulsa(pasta, mapa, d, semana, ultima) {
    if (d.origem !== 'manual' || soDigitos(d.ticket)) { return null; }
    var alvo = normalizar(d.assunto);
    if (!alvo) { return null; }
    for (var l = ultima; l > mapa.linha; l--) {
      if (soDigitos(pasta.ler(l, mapa.colunas.ticket))) { continue; }
      if (normalizar(pasta.ler(l, mapa.colunas.problema)) !== alvo) { continue; }
      var dia = dataDaCelula(pasta.ler(l, mapa.colunas.data));
      if (dia === null || (dia >= semana.inicio && dia <= semana.fim)) { return l; }
    }
    return null;
  }

  /* Completa a linha da demanda. Devolve o que fez: 'atualizado' quando a
     tratativa entrou no historico, 'status' quando so o Status mudou. */
  function consolidar(pasta, mapa, linha, d, quando) {
    var acao = 'status';
    var texto = String(d.feito || '').replace(/\s*\n\s*/g, ' ').trim();
    if (texto) {
      var diaLinha = dataDaCelula(pasta.ler(linha, mapa.colunas.data));
      var entradas = lerHistorico(pasta.ler(linha, mapa.colunas.feito), diaLinha);
      var cortesia = CFG.ignorarCortesia && d.origem === 'zendesk' &&
                     entradas.length > 0 && ehSoCortesia(texto);
      if (!cortesia && !jaConsta(entradas, texto)) {
        /* A linha so tinha cortesia? Entao a tratativa de verdade toma o lugar. */
        if (CFG.ignorarCortesia && entradas.length === 1 && ehSoCortesia(entradas[0].texto)) {
          entradas = [];
        }
        entradas.push({ dia: pad(quando.getDate()) + '/' + pad(quando.getMonth() + 1),
                        texto: texto });
        var cel = pasta.escreverTexto(linha, mapa.colunas.feito, escreverHistorico(entradas));
        aplicarQuebra(pasta, cel);
        alturaAutomatica(pasta, linha);
        if (CFG.dataDaDemanda === 'ultima') {
          pasta.escreverNumero(linha, mapa.colunas.data, pasta.serialData(quando));
        }
        acao = 'atualizado';
      }
    }
    if (d.status) { pasta.escreverTexto(linha, mapa.colunas.status, d.status); }
    /* O assunto acompanha o ticket: se ele foi renomeado no Zendesk (ganhou a
       UF na frente, por exemplo), a linha passa a mostrar o nome atual. */
    var prob = pasta.ler(linha, mapa.colunas.problema);
    if (d.assunto && (d.origem === 'zendesk' || prob === null || !String(prob).trim())) {
      pasta.escreverTexto(linha, mapa.colunas.problema, d.assunto);
    }
    return acao;
  }

  function aplicarQuebra(pasta, celula) {
    try {
      var atual = celula.getAttribute('s');
      var novo = pasta.estiloComQuebra(atual === null ? 0 : parseInt(atual, 10));
      if (novo !== null && novo !== undefined) { celula.setAttribute('s', String(novo)); }
    } catch (e) { /* estilo e cosmetico: nunca impede a gravacao */ }
  }

  /* Reabre os bytes gerados e confere ANTES de tocar no arquivo em disco. */
  function conferir(bytes, mapa, aplicadas) {
    var criadas = aplicadas.filter(function (a) { return a.linha; });
    return XLSX.abrir(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
      .then(function (p2) {
        return p2.abrirAba(mapa.aba).then(function () {
          for (var i = 0; i < criadas.length; i++) {
            var v = p2.ler(criadas[i].linha, mapa.colunas.ticket);
            if (v === null || String(v).trim() === '') {
              throw new Error('conferencia falhou: linha ' + criadas[i].linha + ' saiu vazia');
            }
          }
          var cab = p2.ler(mapa.linha, mapa.colunas.feito);
          if (cab === null || String(cab).trim() === '') {
            throw new Error('conferencia falhou: cabecalho sumiu');
          }
          return true;
        });
      })['catch'](function (e) {
        throw new Error('Arquivo gerado nao passou na conferencia (' + e.message +
                        '). NADA foi gravado; sua planilha continua intacta.');
      });
  }

  function backup(bytes) {
    return dirHandle.getDirectoryHandle(CFG.pastaBackups, { create: true })
      .then(function (dir) {
        var nome = planilhaDaPasta().replace(/\.xlsx$/i, '') + '__' + carimbo(agora()) + '.xlsx';
        return dir.getFileHandle(nome, { create: true }).then(function (fh) {
          return escrever(fh, bytes);
        }).then(function () { return limparBackups(dir); });
      })['catch'](function (e) {
        registrar('AVISO: nao foi possivel criar backup (' + e.message + ').');
      });
  }

  function limparBackups(dir) {
    var nomes = [];
    var it = dir.keys();
    function proximo() {
      return it.next().then(function (r) {
        if (r.done) { return null; }
        var nome = String(r.value);
        if (nome.indexOf('__') > 0 && /\.xlsx$/i.test(nome)) { nomes.push(nome); }
        return proximo();
      });
    }
    return proximo().then(function () {
      nomes.sort();   /* o carimbo AAAAMMDD_HHMMSS ordena cronologicamente */
      var excedente = nomes.slice(0, Math.max(0, nomes.length - CFG.manterBackups));
      return excedente.reduce(function (p, n) {
        return p.then(function () { return dir.removeEntry(n)['catch'](function () {}); });
      }, Promise.resolve());
    })['catch'](function () { /* limpeza e opcional, nunca bloqueia a gravacao */ });
  }

  function escrever(fileHandle, bytes) {
    return fileHandle.createWritable().then(function (w) {
      return w.write(bytes).then(function () { return w.close(); });
    });
  }

  /* ============================================ RESUMO SEMANAL EM PDF ==== */
  /* O que o gestor recebe junto com a planilha. Duas camadas, de proposito:
     o PANORAMA e curto e resume; o DETALHAMENTO lista tudo, demanda por
     demanda. Assim resumir nunca custa informacao - o texto inteiro continua
     no documento, logo abaixo do resumo. A IA escreve so o paragrafo de
     abertura, e sempre a partir das linhas da planilha: nada e inventado. */

  function nomeParaArquivo(txt) {
    return String(txt || '')
      .normalize('NFKD').replace(/[̀-ͯ]/g, '')
      .replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'Analista';
  }

  function primeiroNome() {
    return nomeParaArquivo(String(meuNome || '').trim().split(/\s+/)[0]);
  }

  /* {nome} sai do proprio Zendesk. E o que permite um colega usar o painel sem
     editar uma linha de codigo: o nome no arquivo se ajusta sozinho. */
  function aplicarPadrao(padrao, d) {
    return String(padrao)
      .replace('{dd}', pad(d.getDate()))
      .replace('{mm}', pad(d.getMonth() + 1))
      .replace('{aaaa}', d.getFullYear())
      .replace('{nome}', primeiroNome());
  }

  function dataBR(iso) {
    if (!iso) { return ''; }
    return iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4);
  }

  function linhasDaPlanilha(pasta, mapa, ultima) {
    var c = mapa.colunas;
    var itens = [];
    for (var l = mapa.linha + 1; l <= ultima; l++) {
      var it = {
        linha: l,
        dataISO: dataDaCelula(pasta.ler(l, c.data)),
        ticket: textoDaCelula(pasta.ler(l, c.ticket)),
        problema: textoDaCelula(pasta.ler(l, c.problema)),
        feito: textoDaCelula(pasta.ler(l, c.feito)),
        status: textoDaCelula(pasta.ler(l, c.status))
      };
      if (!it.ticket && !it.problema && !it.feito) { continue; }
      itens.push(it);
    }
    return itens;
  }

  /* Varias atualizacoes do mesmo ticket viram UMA demanda com varias
     ocorrencias - que e como o gestor pensa. Atividade avulsa nunca agrupa:
     cada uma e uma demanda propria. */
  function agruparDemandas(itens) {
    var porChave = {}, ordem = [];
    itens.forEach(function (it) {
      var digitos = soDigitos(it.ticket);
      var chave = digitos ? ('t' + digitos) : ('avulsa' + it.linha);
      if (!porChave[chave]) {
        porChave[chave] = {
          ticket: digitos ? it.ticket : CFG.ticketAvulso,
          temTicket: !!digitos,
          problema: it.problema,
          status: it.status,
          primeira: it.dataISO,
          ultima: it.dataISO,
          ocorrencias: [],
          entradas: []
        };
        ordem.push(chave);
      }
      var g = porChave[chave];
      g.ocorrencias.push(it);
      lerHistorico(it.feito, it.dataISO).forEach(function (e) {
        if (!jaConsta(g.entradas, e.texto)) { g.entradas.push(e); }
      });
      if (it.problema) { g.problema = it.problema; }      /* vale o nome atual */
      if (it.status) { g.status = it.status; }        /* vale o ultimo status */
      if (it.dataISO) {
        if (!g.primeira || it.dataISO < g.primeira) { g.primeira = it.dataISO; }
        if (!g.ultima || it.dataISO > g.ultima) { g.ultima = it.dataISO; }
      }
    });
    return ordem.map(function (k) {
      var g = porChave[k];
      /* Na linha unica, a coluna Data guarda o dia em que a demanda entrou. O
         dia da ultima tratativa esta no historico. */
      g.entradas.forEach(function (e) {
        if (!e.dia || !g.primeira) { return; }
        var iso = g.primeira.slice(0, 4) + '-' + e.dia.slice(3, 5) + '-' + e.dia.slice(0, 2);
        if (iso < g.primeira) { return; }           /* virada de ano: ignora */
        if (!g.ultima || iso > g.ultima) { g.ultima = iso; }
      });
      g.uf = ufDoAssunto(g.problema);
      g.resolvida = normalizar(g.status) === normalizar(CFG.statusResolvido);
      /* Com varias tratativas, as que sao so cortesia saem do documento. */
      if (CFG.ignorarCortesia && g.entradas.length > 1) {
        var uteis = g.entradas.filter(function (e) { return !ehSoCortesia(e.texto); });
        if (uteis.length) { g.entradas = uteis; }
      }
      return g;
    });
  }

  /* "GO | Setor - ...", "SP/Capital - ...", "PE - ..." -> a UF que abre o assunto. */
  function ufDoAssunto(assunto) {
    var m = /^\s*([A-Za-z]{2})(?=\s*[-|\/–—:]|\s+[-|–—])/.exec(String(assunto || ''));
    var uf = m ? m[1].toUpperCase() : '';
    return UFS.indexOf(uf) >= 0 ? uf : '';
  }

  function distribuicaoPorUF(grupos) {
    var mapa = {}, ordem = [];
    grupos.forEach(function (g) {
      var k = g.uf || 'Outras';
      if (!mapa[k]) { mapa[k] = { uf: k, total: 0, resolvidas: 0, andamento: 0 }; ordem.push(k); }
      mapa[k].total++;
      if (g.resolvida) { mapa[k].resolvidas++; } else { mapa[k].andamento++; }
    });
    return ordem.map(function (k) { return mapa[k]; }).sort(function (a, b) {
      if (a.uf === 'Outras') { return 1; }
      if (b.uf === 'Outras') { return -1; }
      return b.total - a.total || (a.uf < b.uf ? -1 : 1);
    });
  }

  function contarPanorama(itens, grupos) {
    var resolvidas = 0, andamento = 0, avulsas = 0;
    grupos.forEach(function (g) {
      if (normalizar(g.status) === normalizar(CFG.statusResolvido)) { resolvidas++; }
      else { andamento++; }
      if (!g.temTicket) { avulsas++; }
    });
    var tratativas = grupos.reduce(function (t, g) { return t + (g.entradas.length || 1); }, 0);
    return {
      atualizacoes: tratativas,
      demandas: grupos.length,
      resolvidas: resolvidas,
      andamento: andamento,
      avulsas: avulsas,
      tickets: grupos.length - avulsas
    };
  }

  function periodoDeItens(itens) {
    var datas = itens.map(function (i) { return i.dataISO; })
                     .filter(Boolean).sort();
    if (!datas.length) {
      var s = semanaDe(agora());
      return { inicio: s.inicio, fim: s.fim,
               texto: dataBR(s.inicio) + ' a ' + dataBR(s.fim) };
    }
    var ini = datas[0], fim = datas[datas.length - 1];
    return { inicio: ini, fim: fim,
             texto: ini === fim ? dataBR(ini) : (dataBR(ini) + ' a ' + dataBR(fim)) };
  }

  /* Logo e selo vem da pasta, nao do codigo. Assim o favorito nao carrega
     imagem embutida, e quem receber o painel so precisa largar os arquivos ao
     lado da planilha. Faltando, o documento sai sem eles. */
  function carregarImagemDaPasta(nome) {
    if (!dirHandle || !nome) { return Promise.resolve(null); }
    return dirHandle.getFileHandle(nome)
      .then(function (fh) { return fh.getFile(); })
      .then(function (f) { return PDFS.imagemParaJpeg(f); })
      ['catch'](function () { return null; });
  }

  var ESQUEMA_PANORAMA = {
    type: 'object',
    properties: { texto: { type: 'string' } },
    required: ['texto'],
    additionalProperties: false
  };

  /* O paragrafo de abertura. Recebe a lista real de demandas e so pode falar
     do que esta nela. Passa pela mesma porta de saida do resto: portugues
     conferido e sem narrador. */
  function panoramaIA(grupos) {
    if (modoResumo === 'regras') { return Promise.resolve(null); }

    var lista = '';
    for (var i = 0; i < grupos.length; i++) {
      var g = grupos[i];
      var acoes = g.entradas.map(function (e) { return e.texto; })
                            .filter(Boolean).join('; ').slice(0, 320);
      var bloco = (i + 1) + '. ' + (g.temTicket ? 'Ticket ' + g.ticket : 'Sem ticket') +
                  ' - ' + (g.problema || 'sem assunto') + ' [' + (g.status || '-') + ']' +
                  (acoes ? ': ' + acoes : '') + '\n';
      if (lista.length + bloco.length > CFG.maxCaracteresContextoIA) { break; }
      lista += bloco;
    }
    if (!lista) { return Promise.resolve(null); }

    return sessaoLimpa(false).then(function (s) {
      if (!s) { return null; }
      function encerrar() {
        try { if (s !== sessaoLM) { s.destroy(); } } catch (e) {}
      }
      return s.prompt(
        'Escreva UM paragrafo corrido, de 4 a 6 frases, resumindo a semana de\n' +
        'trabalho para um gestor que nao acompanha os chamados no dia a dia.\n' +
        REGRAS_ESCRITA +
        '- Seja objetivo. AGRUPE as demandas por tipo de assunto; nao percorra\n' +
        '  uma a uma nem descreva cada caso separadamente.\n' +
        '- Nao cite numero de ticket, nome de pessoa nem lista de topicos.\n' +
        '- Diga os tipos de demanda tratados, o que foi concluido, o que segue\n' +
        '  em andamento e do que depende (fornecedor, aprovacao, entrega etc.).\n' +
        '- Use apenas a lista abaixo. Nao acrescente nada que nao esteja nela.\n\n' +
        'ATIVIDADES DA SEMANA:\n' + lista,
        { responseConstraint: ESQUEMA_PANORAMA }
      ).then(function (r) {
        var o = JSON.parse(r);
        var fim = validarPortugues(o && o.texto, lista);
        encerrar();
        return fim;
      })['catch'](function () { encerrar(); return null; });
    })['catch'](function () { return null; });
  }

  var ESQUEMA_SINTESE = {
    type: 'object',
    properties: { sintese: { type: 'string' } },
    required: ['sintese'],
    additionalProperties: false
  };

  /* Uma ou duas frases dizendo onde a demanda esta. So para demandas com mais
     de uma tratativa - com uma so, o proprio texto ja e a sintese. Roda uma de
     cada vez (o modelo local nao atende em paralelo) e qualquer falha apenas
     deixa aquela demanda sem sintese: o historico completo esta logo abaixo. */
  function sintesesIA(grupos) {
    if (!CFG.sinteseIA || modoResumo === 'regras') { return Promise.resolve(0); }
    var alvos = grupos.filter(function (g) { return g.entradas.length >= 2; })
                      .slice(0, CFG.maxSintesesIA);
    if (!alvos.length) { return Promise.resolve(0); }
    registrar('Escrevendo a sintese de ' + alvos.length + ' demanda(s) com a IA local...');
    var feitas = 0;

    function uma(g) {
      var historico = g.entradas.map(function (e) {
        return (e.dia ? e.dia + ': ' : '') + e.texto;
      }).join('\n').slice(0, CFG.maxCaracteresContextoIA);
      return sessaoLimpa(false).then(function (s) {
        if (!s) { return null; }
        function encerrar() {
          try { if (s !== sessaoLM) { s.destroy(); } } catch (e) {}
        }
        return s.prompt(
          'Abaixo esta o historico de UMA demanda de atendimento, em ordem de data.\n' +
          'Escreva a sintese em 1 ou 2 frases: o que foi tratado e qual e a\n' +
          'situacao agora (' + (g.resolvida ? 'a demanda foi concluida' :
                                              'a demanda segue em andamento') + ').\n' +
          REGRAS_ESCRITA +
          '- Nao repita as datas uma a uma e nao copie o historico: assimile.\n\n' +
          'ASSUNTO: ' + (g.problema || 'sem assunto') + '\n' +
          'HISTORICO:\n' + historico,
          { responseConstraint: ESQUEMA_SINTESE }
        ).then(function (r) {
          var o = JSON.parse(r);
          var fim = validarPortugues(o && o.sintese, historico);
          encerrar();
          if (fim) { g.sintese = fim.slice(0, 420); feitas++; }
          return null;
        })['catch'](function () { encerrar(); return null; });
      })['catch'](function () { return null; });
    }

    return alvos.reduce(function (p, g) {
      return p.then(function () { return uma(g); });
    }, Promise.resolve()).then(function () { return feitas; });
  }

  /* Frase factual montada por regra. Vale sozinha quando a IA esta desligada,
     e junto com o paragrafo quando ela esta ativa. */
  /* Texto do documento vai ACENTUADO. O codigo evita acento por conveniencia,
     mas isto aqui e o que o gestor le. */
  function panoramaPorRegra(n, periodo) {
    var partes = ['No período de ' + periodo.texto + ' foram registradas ' +
                  n.atualizacoes + (n.atualizacoes === 1 ? ' tratativa' : ' tratativas') +
                  ' em ' + n.demandas + (n.demandas === 1 ? ' demanda' : ' demandas') + '.'];
    partes.push(n.resolvidas + (n.resolvidas === 1 ? ' foi concluída' : ' foram concluídas') +
                ' e ' + n.andamento +
                (n.andamento === 1 ? ' segue em andamento.' : ' seguem em andamento.'));
    if (n.avulsas) {
      partes.push(n.avulsas + (n.avulsas === 1 ? ' atividade não teve' : ' atividades não tiveram') +
                  ' ticket associado.');
    }
    return partes.join(' ');
  }

  function desenharResumo(grupos, n, periodo, logo, selo, paragrafo) {
    var rel = new PDFS.Relatorio({
      logo: logo ? { apelido: 'Logo', largura: logo.largura, altura: logo.altura } : null,
      selo: selo ? { apelido: 'Selo', largura: selo.largura, altura: selo.altura } : null,
      rodape: 'Gerado automaticamente a partir de ' + planilhaDaPasta()
    });
    if (logo) { rel.doc.registrarImagem('Logo', logo.bytes, logo.largura, logo.altura); }
    if (selo) { rel.doc.registrarImagem('Selo', selo.bytes, selo.largura, selo.altura); }
    rel.novaPagina();

    var hoje = agora();
    rel.rotulo('Resumo semanal de atividades');
    rel.titulo('Semana de ' + periodo.texto);
    rel.apoio((meuNome || 'Analista') + '  —  emitido em ' + pad(hoje.getDate()) + '/' +
              pad(hoje.getMonth() + 1) + '/' + hoje.getFullYear() + ' às ' +
              pad(hoje.getHours()) + ':' + pad(hoje.getMinutes()) +
              '. Documento gerado a partir do controle semanal de atividades.');

    /* ---------------------------------------------------------- panorama */
    rel.secao('Panorama da semana');
    rel.indicadores([
      { valor: n.demandas, rotulo: 'Demandas' },
      { valor: n.atualizacoes, rotulo: 'Tratativas' },
      { valor: n.resolvidas, rotulo: 'Concluídas', cor: PDFS.COR.verde },
      { valor: n.andamento, rotulo: 'Em andamento', cor: PDFS.COR.azulAco }
    ]);
    rel.paragrafo(panoramaPorRegra(n, periodo), { cor: PDFS.COR.cinza, tamanho: 8.5 });

    if (paragrafo) {
      rel.secao('A semana em resumo');
      rel.paragrafo(semMarca(paragrafo));
    }

    /* -------------------------------------------------- por unidade (UF) */
    var porUF = distribuicaoPorUF(grupos);
    if (porUF.length >= 2) {
      rel.secao('Distribuição por unidade');
      rel.tabela(
        ['Unidade', 'Demandas', 'Concluídas', 'Em andamento'],
        [0.31, 0.23, 0.23, 0.23],
        porUF.map(function (u) {
          return [u.uf, String(u.total), String(u.resolvidas), String(u.andamento)];
        }),
        { colunasNegrito: [0], zebra: true }
      );
    }

    /* ------------------------------------------------------ visao geral */
    rel.secao('Visão geral das demandas');
    rel.tabela(
      ['Ticket', 'Problema / Solicitação', 'Período', 'Trat.', 'Situação'],
      [0.12, 0.47, 0.17, 0.07, 0.17],
      grupos.map(function (g) {
        return [g.temTicket ? '#' + g.ticket : '—',
                semMarca(g.problema) || '(sem assunto)', periodoCurto(g),
                String(g.entradas.length || 1), g.status || '—'];
      }),
      { colunaDestaque: 0, zebra: true, tamanho: 8 }
    );

    if (!CFG.detalharDemandas) { return rel.gerar(); }

    /* ----------------------------------------------------- detalhamento
       Em andamento primeiro: e o que o gestor precisa acompanhar. Cada demanda
       traz a sintese (quando ha IA) e o historico datado das tratativas. */
    var detalhar = grupos.slice(0, CFG.maxDemandasDetalhadas);
    var blocos = [
      { titulo: 'Demandas em andamento',
        nota: 'Seguem abertas para a próxima semana. A última tratativa de cada uma ' +
              'mostra a situação atual.',
        itens: detalhar.filter(function (g) { return g.temTicket && !g.resolvida; }) },
      { titulo: 'Demandas concluídas na semana',
        nota: '',
        itens: detalhar.filter(function (g) { return g.temTicket && g.resolvida; }) },
      { titulo: 'Atividades sem ticket',
        nota: 'Reuniões, treinamentos, direcionamentos e demais atividades lançadas ' +
              'diretamente no painel.',
        itens: detalhar.filter(function (g) { return !g.temTicket; }) }
    ];

    blocos.forEach(function (b) {
      if (!b.itens.length) { return; }
      rel.secao(b.titulo + ' (' + b.itens.length + ')');
      if (b.nota) { rel.paragrafo(b.nota, { cor: PDFS.COR.cinza, tamanho: 8.5 }); }
      b.itens.forEach(function (g) {
        var qtd = g.entradas.length || 1;
        rel.demanda({
          ticket: g.temTicket ? '#' + g.ticket : '',
          titulo: semMarca(g.problema) || '(sem assunto)',
          meta: periodoCurto(g) + '   ·   ' + qtd + (qtd === 1 ? ' tratativa' : ' tratativas'),
          status: g.status || 'não informado',
          corStatus: g.resolvida ? PDFS.COR.verde : PDFS.COR.azulAco,
          sintese: g.sintese ? semMarca(g.sintese) : '',
          entradas: g.entradas.map(function (e) {
            return { dia: e.dia, texto: semMarca(e.texto) };
          })
        });
      });
    });

    if (grupos.length > detalhar.length) {
      rel.paragrafo('Outras ' + (grupos.length - detalhar.length) + ' demanda(s) constam na ' +
                    'visão geral desta semana e na planilha anexa.',
                    { cor: PDFS.COR.cinza, tamanho: 8.5 });
    }

    return rel.gerar();
  }

  /* O motor de PDF usa *asterisco* como marca de negrito. Texto vindo da
     planilha nao pode acionar isso por acidente. */
  function semMarca(t) { return String(t || '').replace(/\*/g, ''); }

  function periodoCurto(g) {
    if (!g.primeira) { return 'sem data'; }
    return g.primeira === g.ultima ? diaMes(g.primeira)
                                   : diaMes(g.primeira) + ' a ' + diaMes(g.ultima);
  }

  function montarResumoPdf(pasta, mapa, ultima) {
    var itens = linhasDaPlanilha(pasta, mapa, ultima);
    if (!itens.length) { return Promise.resolve(null); }

    var grupos = agruparDemandas(itens);
    var n = contarPanorama(itens, grupos);
    var periodo = periodoDeItens(itens.concat(grupos.map(function (g) {
      return { dataISO: g.ultima };
    })));

    return Promise.all([
      carregarImagemDaPasta(CFG.arquivoLogo),
      carregarImagemDaPasta(CFG.arquivoSelo)
    ]).then(function (imgs) {
      if (!imgs[0]) {
        registrar('Resumo sem logo: coloque "' + CFG.arquivoLogo + '" na pasta da planilha ' +
                  'para ele aparecer no cabecalho.');
      }
      return panoramaIA(grupos).then(function (paragrafo) {
        return sintesesIA(grupos).then(function () {
          var bytes = desenharResumo(grupos, n, periodo, imgs[0], imgs[1], paragrafo);
          PDFS.conferir(bytes);        /* estourou aqui, nada e gravado */
          return { bytes: bytes, periodo: periodo, numeros: n, comIA: !!paragrafo };
        });
      });
    });
  }

  /* Grava o PDF em arquivos_semanais/, ao lado da copia da planilha. Falhar
     aqui NUNCA pode derrubar o fechamento: a planilha e o registro oficial, o
     PDF e o acompanhamento. Por isso o catch devolve null em vez de propagar. */
  function gravarResumoPdf(pasta, mapa, ultima) {
    if (!CFG.resumoPdf) { return Promise.resolve(null); }
    return montarResumoPdf(pasta, mapa, ultima).then(function (r) {
      if (!r) {
        registrar('Resumo nao gerado: a planilha nao tem linhas preenchidas.');
        return null;
      }
      var nome = aplicarPadrao(CFG.padraoNomeResumo, agora());
      return dirHandle.getDirectoryHandle(CFG.pastaSemanais, { create: true })
        .then(function (dir) {
          return nomeLivre(dir, nome).then(function (final) {
            return dir.getFileHandle(final, { create: true }).then(function (fh) {
              return escrever(fh, r.bytes).then(function () {
                registrar('Resumo em PDF criado: ' + final + ' (' +
                          r.numeros.demandas + ' demandas, ' +
                          (r.bytes.length / 1024).toFixed(0) + ' KB' +
                          (r.comIA ? ', com paragrafo da IA' : ', sem IA') + ').');
                return final;
              });
            });
          });
        });
    })['catch'](function (e) {
      registrar('AVISO: nao foi possivel gerar o resumo em PDF (' +
                String((e && e.message) || e) + '). A planilha nao foi afetada.');
      return null;
    });
  }

  /* Botao "Resumo": gera o PDF da semana atual SEM fechar nada. Serve para
     conferir o texto antes do fechamento - depois de fechar, a planilha ja
     esta limpa e nao da mais para revisar. */
  function resumoAvulso() {
    registrar('Gerando o resumo da semana...');
    return abrirPlanilha().then(function (ctx) {
      if (ctx.ultima <= ctx.mapa.linha) {
        registrar('A planilha esta vazia: nao ha o que resumir.');
        return null;
      }
      return gravarResumoPdf(ctx.pasta, ctx.mapa, ctx.ultima);
    })['catch'](function (e) {
      registrar('ERRO ao gerar o resumo: ' + String((e && e.message) || e));
      return null;
    });
  }

  /* ================================================ FECHAMENTO SEMANAL === */
  function precisaFechar(pasta, mapa) {
    var ultima = ultimaLinhaPreenchida(pasta, mapa);
    if (ultima <= mapa.linha) { return false; }

    var semana = semanaDe(agora());
    var temDeOutraSemana = false, temDaSemana = false;
    for (var l = mapa.linha + 1; l <= ultima; l++) {
      var dia = dataDaCelula(pasta.ler(l, mapa.colunas.data));
      if (dia === null) { continue; }
      if (dia < semana.inicio) { temDeOutraSemana = true; }
      if (dia >= semana.inicio && dia <= semana.fim) { temDaSemana = true; }
    }
    if (temDeOutraSemana && !temDaSemana) { return true; }   /* semana virou */

    var d = agora();
    return CFG.fecharNaSexta && d.getDay() === 5 && d.getHours() >= CFG.horaFechamento;
  }

  function fecharSemana(forcar) {
    return conectarPasta(true).then(function (ok) {
      if (!ok) { registrar('Planilha nao conectada.'); return false; }
      return dirHandle.getFileHandle(planilhaDaPasta()).then(function (fh) {
        return fh.getFile().then(function (f) { return f.arrayBuffer(); })
          .then(function (buf) {
            var originais = new Uint8Array(buf.slice(0));
            return XLSX.abrir(buf).then(function (pasta) {
              return escolherAbaEMapa(pasta).then(function (mapa) {
                var ultima = ultimaLinhaPreenchida(pasta, mapa);
                if (ultima <= mapa.linha) {
                  registrar('Fechamento: planilha ja esta vazia, nada a arquivar.');
                  return false;
                }
                if (!forcar && !precisaFechar(pasta, mapa)) {
                  registrar('Fechamento: ainda nao e hora.');
                  return false;
                }
                return arquivarEntaoLimpar(originais, pasta, mapa, ultima, fh);
              });
            });
          });
      });
    })['catch'](function (e) {
      registrar('ERRO no fechamento: ' + e.message);
      return false;
    });
  }

  function arquivarEntaoLimpar(bytesOriginais, pasta, mapa, ultima, fileHandle) {
    /* aplicarPadrao entende tambem {nome}, preenchido com o primeiro nome lido
       do Zendesk. O padrao atual nao usa - fica disponivel para quem quiser o
       arquivo nomeado sozinho, sem editar o codigo. */
    var nome = aplicarPadrao(CFG.padraoNomeSemanal, agora());

    /* 1) copia semanal ------------------------------------------------- */
    return dirHandle.getDirectoryHandle(CFG.pastaSemanais, { create: true })
      .then(function (dir) {
        return nomeLivre(dir, nome).then(function (final) {
          return dir.getFileHandle(final, { create: true }).then(function (fh) {
            return escrever(fh, bytesOriginais).then(function () {
              /* 2) conferir que a copia abre e tem as linhas ---------- */
              return fh.getFile().then(function (f) { return f.arrayBuffer(); })
                .then(function (buf) { return XLSX.abrir(buf); })
                .then(function (p2) {
                  return p2.abrirAba(mapa.aba).then(function () {
                    var n = ultimaLinhaPreenchida(p2, mapa);
                    if (n !== ultima) {
                      throw new Error('copia com ' + (n - mapa.linha) + ' linhas, esperado ' +
                                      (ultima - mapa.linha));
                    }
                    registrar('Arquivo semanal criado e conferido: ' + final + '.');
                    return final;
                  });
                });
            });
          });
        });
      })
      /* 3) o resumo em PDF, com os dados AINDA em memoria ----------------
         Tem que ser aqui: um passo adiante a planilha ja foi limpa e nao
         haveria mais o que resumir. */
      .then(function (final) {
        return gravarResumoPdf(pasta, mapa, ultima).then(function (pdf) {
          return { xlsx: final, pdf: pdf };
        });
      })
      /* 4) so agora limpar o original ---------------------------------- */
      .then(function (saidas) {
        return backup(bytesOriginais).then(function () {
          var apagadas = limparLinhas(pasta, mapa, ultima);
          return XLSX.gerarBytes(pasta).then(function (bytes) {
            return escrever(fileHandle, bytes).then(function () {
              registrar('Fechamento concluido: ' + apagadas + ' linha(s) limpas. ' +
                        'Planilha pronta para a nova semana. Arquivos em ' +
                        CFG.pastaSemanais + ': ' + saidas.xlsx +
                        (saidas.pdf ? ' e ' + saidas.pdf : ' (sem PDF)') + '.');
              var est = ler(CH.estado, {});
              est.ultimoFechamento = Date.now();
              est.ultimoArquivoSemanal = saidas.xlsx;
              est.ultimoResumoPdf = saidas.pdf || '';
              gravar(CH.estado, est);
              return true;
            });
          });
        });
      })['catch'](function (e) {
        registrar('ERRO: fechamento abortado (' + e.message +
                  '). O arquivo original NAO foi alterado.');
        return false;
      });
  }

  /* Preserva a extensao original: serve tanto para .xlsx quanto para .pdf. */
  function nomeLivre(dir, nome) {
    return dir.getFileHandle(nome).then(function () {
      var m = /^(.*?)(\.[A-Za-z0-9]+)$/.exec(nome);
      var base = m ? m[1] : nome;
      var ext = m ? m[2] : '';
      var tentar = function (n) {
        if (n > 20) { throw new Error('nao achei nome livre'); }
        var alt = base + '_' + n + ext;
        return dir.getFileHandle(alt).then(function () { return tentar(n + 1); })
          ['catch'](function () { return alt; });
      };
      return tentar(2);
    })['catch'](function () { return nome; });
  }

  function limparLinhas(pasta, mapa, ultima) {
    var apagadas = 0;
    for (var linha = mapa.linha + 1; linha <= ultima; linha++) {
      var row = pasta.acharLinha(linha);
      if (!row) { continue; }
      var cs = row.getElementsByTagName('c');
      var tocou = false;
      for (var i = 0; i < cs.length; i++) {
        if (cs[i].getElementsByTagName('f').length) { continue; }   /* formula fica */
        if (!cs[i].firstChild) { continue; }
        while (cs[i].firstChild) { cs[i].removeChild(cs[i].firstChild); }
        cs[i].removeAttribute('t');
        tocou = true;
      }
      if (tocou) { apagadas++; }
    }
    return apagadas;
  }

  /* ============================================ CORRECAO DE LINHAS ======= */
  /* Objetivo: nunca mais abrir o Excel para consertar um texto errado. O painel
     le as ultimas linhas do arquivo, deixa editar campo a campo e grava de
     volta com o mesmo cuidado da gravacao normal - backup antes, conferencia do
     arquivo gerado depois, disco so no fim. */

  function apagarCelula(pasta, linha, col) {
    var c = pasta.acharCelula(linha, col);
    if (!c) { return; }
    if (c.getElementsByTagName('f').length) { return; }      /* formula fica */
    while (c.firstChild) { c.removeChild(c.firstChild); }
    c.removeAttribute('t');
  }

  function textoDaCelula(v) {
    if (v === null || v === undefined) { return ''; }
    return String(v).trim();
  }

  function dataDigitada(txt) {
    var m = /^\s*(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-](\d{2,4}))?\s*$/.exec(String(txt || ''));
    if (!m) { return null; }
    var dia = parseInt(m[1], 10), mes = parseInt(m[2], 10);
    if (dia < 1 || dia > 31 || mes < 1 || mes > 12) { return null; }
    var ano = m[3] ? (m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10))
                   : agora().getFullYear();
    var d = new Date(ano, mes - 1, dia);
    return isNaN(d.getTime()) ? null : d;
  }

  /* Abre o arquivo e devolve tudo de que uma edicao precisa. */
  function abrirPlanilha() {
    return conectarPasta(true).then(function (ok) {
      if (!ok) { throw new Error('planilha nao conectada; clique em "Conectar planilha"'); }
      return dirHandle.getFileHandle(planilhaDaPasta());
    }).then(function (fh) {
      return fh.getFile().then(function (f) { return f.arrayBuffer(); }).then(function (buf) {
        var originais = new Uint8Array(buf.slice(0));
        return XLSX.abrir(buf).then(function (pasta) {
          return escolherAbaEMapa(pasta).then(function (mapa) {
            return { fh: fh, pasta: pasta, mapa: mapa, originais: originais,
                     ultima: ultimaLinhaPreenchida(pasta, mapa) };
          });
        });
      });
    });
  }

  function listarUltimasLinhas() {
    return abrirPlanilha().then(function (ctx) {
      var c = ctx.mapa.colunas, itens = [];
      for (var l = ctx.ultima; l > ctx.mapa.linha && itens.length < CFG.linhasParaCorrigir; l--) {
        var iso = dataDaCelula(ctx.pasta.ler(l, c.data));
        itens.push({
          linha: l,
          data: iso ? (iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4)) : '',
          ticket: textoDaCelula(ctx.pasta.ler(l, c.ticket)),
          problema: textoDaCelula(ctx.pasta.ler(l, c.problema)),
          feito: textoDaCelula(ctx.pasta.ler(l, c.feito)),
          status: textoDaCelula(ctx.pasta.ler(l, c.status))
        });
      }
      return itens;
    });
  }

  /* Reabre os bytes gerados e confere celula a celula ANTES de tocar no disco. */
  function conferirEdicao(bytes, mapa, checagens, linhasEsperadas) {
    return XLSX.abrir(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
      .then(function (p2) {
        return p2.abrirAba(mapa.aba).then(function () {
          var cab = p2.ler(mapa.linha, mapa.colunas.feito);
          if (cab === null || String(cab).trim() === '') { throw new Error('o cabecalho sumiu'); }
          (checagens || []).forEach(function (ck) {
            if (textoDaCelula(p2.ler(ck.linha, ck.col)) !== String(ck.valor)) {
              throw new Error('a celula ' + XLSX.letraColuna(ck.col) + ck.linha +
                              ' nao gravou como esperado');
            }
          });
          if (typeof linhasEsperadas === 'number') {
            var n = ultimaLinhaPreenchida(p2, mapa);
            if (n !== linhasEsperadas) {
              throw new Error('sobraram ' + (n - mapa.linha) + ' linha(s), esperado ' +
                              (linhasEsperadas - mapa.linha));
            }
          }
          return true;
        });
      })['catch'](function (e) {
        throw new Error('o arquivo gerado nao passou na conferencia (' + e.message +
                        '); NADA foi gravado, sua planilha continua intacta');
      });
  }

  /* Envolve qualquer edicao na mesma disciplina da gravacao: trava entre abas,
     backup do arquivo como estava, conferencia e so entao o disco. */
  function editar(descricao, mudar) {
    if (!travar()) {
      registrar('Outra aba do Zendesk esta gravando agora. Tente de novo em instantes.');
      return Promise.resolve(false);
    }
    return abrirPlanilha().then(function (ctx) {
      var plano = mudar(ctx);
      return XLSX.gerarBytes(ctx.pasta).then(function (bytes) {
        return conferirEdicao(bytes, ctx.mapa, plano.checagens, plano.linhasEsperadas)
          .then(function () { return backup(ctx.originais); })
          .then(function () { return escrever(ctx.fh, bytes); });
      });
    }).then(function () {
      destravar();
      registrar(descricao);
      return true;
    })['catch'](function (e) {
      destravar();
      var msg = String((e && e.message) || e);
      if ((e && e.name === 'NotFoundError') || /could not be found|not be found/i.test(msg)) {
        dirHandle = null;
        msg = 'a pasta ou a planilha mudou de lugar; clique em "Conectar planilha"';
      }
      registrar('ERRO na correcao: ' + msg);
      return false;
    });
  }

  function salvarCorrecao(linha, campos) {
    return editar('Linha ' + linha + ' corrigida na planilha.', function (ctx) {
      var c = ctx.mapa.colunas;
      if (linha <= ctx.mapa.linha || linha > ctx.ultima) {
        throw new Error('a linha ' + linha + ' saiu da area de dados; reabra a lista');
      }
      var texto = function (v) { return String(v === undefined || v === null ? '' : v).trim(); };
      var checagens = [];

      var dataTxt = texto(campos.data);
      var d = null;
      if (dataTxt) {
        d = dataDigitada(dataTxt);
        if (!d) { throw new Error('data "' + dataTxt + '" nao entendida; use dd/mm/aaaa'); }
        ctx.pasta.escreverNumero(linha, c.data, ctx.pasta.serialData(d));
      } else {
        apagarCelula(ctx.pasta, linha, c.data);
      }

      var tk = texto(campos.ticket);
      if (tk && soDigitos(tk) === tk) { ctx.pasta.escreverNumero(linha, c.ticket, tk); }
      else { ctx.pasta.escreverTexto(linha, c.ticket, tk); }
      checagens.push({ linha: linha, col: c.ticket, valor: tk });

      ctx.pasta.escreverTexto(linha, c.problema, texto(campos.problema));
      checagens.push({ linha: linha, col: c.problema, valor: texto(campos.problema) });

      /* "Nova tratativa": entra no historico com a data de hoje, sem voce
         precisar montar o formato na mao. */
      var feitoFinal = texto(campos.feito);
      var nova = texto(campos.novaTratativa).replace(/\s*\n\s*/g, ' ');
      if (nova) {
        var entradas = lerHistorico(feitoFinal, d ? diaISO(d) : null);
        entradas.push({ dia: pad(agora().getDate()) + '/' + pad(agora().getMonth() + 1),
                        texto: nova.charAt(0).toUpperCase() + nova.slice(1) });
        feitoFinal = escreverHistorico(entradas);
      }
      aplicarQuebra(ctx.pasta, ctx.pasta.escreverTexto(linha, c.feito, feitoFinal));
      alturaAutomatica(ctx.pasta, linha);
      checagens.push({ linha: linha, col: c.feito, valor: feitoFinal });

      ctx.pasta.escreverTexto(linha, c.status, texto(campos.status));
      checagens.push({ linha: linha, col: c.status, valor: texto(campos.status) });

      return { checagens: checagens };
    });
  }

  /* Apaga a linha e puxa as de baixo, para nao deixar buraco no meio da
     planilha. Move so as colunas mapeadas - as que a automacao escreve. Estilo,
     formato de data e validacao pertencem a celula de destino e ficam onde
     estao, entao o visual da planilha nao muda. */
  function excluirLinha(alvo) {
    return editar('Linha ' + alvo + ' excluida; as de baixo subiram uma posicao.',
      function (ctx) {
        var cols = CAMPOS.map(function (k) { return ctx.mapa.colunas[k]; });
        if (alvo <= ctx.mapa.linha || alvo > ctx.ultima) {
          throw new Error('a linha ' + alvo + ' saiu da area de dados; reabra a lista');
        }
        for (var l = alvo; l < ctx.ultima; l++) {
          /* IIFE: sem ela o `l` do forEach seria sempre o ultimo do laco. */
          (function (destino) {
            cols.forEach(function (col) {
              var v = ctx.pasta.ler(destino + 1, col);
              if (v === null || v === undefined || String(v).trim() === '') {
                apagarCelula(ctx.pasta, destino, col);
              } else if (typeof v === 'number') {
                ctx.pasta.escreverNumero(destino, col, v);
              } else {
                var cel = ctx.pasta.escreverTexto(destino, col, String(v));
                if (col === ctx.mapa.colunas.feito) { aplicarQuebra(ctx.pasta, cel); }
              }
            });
          }(l));
        }
        cols.forEach(function (col) { apagarCelula(ctx.pasta, ctx.ultima, col); });
        return { linhasEsperadas: ctx.ultima - 1 };
      });
  }

  /* ============================================ ORGANIZAR A PLANILHA ===== */
  /* Converte o que ja esta na planilha para o formato novo: as varias linhas
     do mesmo ticket viram UMA, com o historico datado. Serve para a semana em
     que voce trocou de versao, e para arrumar a casa quando quiser. Mesma
     disciplina de sempre: backup, conferencia do arquivo gerado, disco no fim. */
  function planejarOrganizacao(pasta, mapa, ultima) {
    var c = mapa.colunas, porChave = {}, ordem = [];
    for (var l = mapa.linha + 1; l <= ultima; l++) {
      var bruta = {
        data: pasta.ler(l, c.data), ticket: pasta.ler(l, c.ticket),
        problema: textoDaCelula(pasta.ler(l, c.problema)),
        feito: textoDaCelula(pasta.ler(l, c.feito)),
        status: textoDaCelula(pasta.ler(l, c.status))
      };
      if (!textoDaCelula(bruta.ticket) && !bruta.problema && !bruta.feito) { continue; }
      var iso = dataDaCelula(bruta.data);
      var digitos = soDigitos(bruta.ticket);
      var chave = digitos ? ('t' + digitos + '@' + (iso ? semanaDe(new Date(iso + 'T12:00:00')).inicio : ''))
                          : ('l' + l);
      if (!porChave[chave]) {
        porChave[chave] = { data: bruta.data, ticket: bruta.ticket, problema: '',
                            status: '', entradas: [], linhas: 0 };
        ordem.push(chave);
      }
      var g = porChave[chave];
      g.linhas++;
      if (bruta.problema) { g.problema = bruta.problema; }
      if (bruta.status) { g.status = bruta.status; }
      lerHistorico(bruta.feito, iso).forEach(function (e) {
        var t = digitos ? limparComentario(e.texto) : limparMarkdown(e.texto).replace(/\s+/g, ' ').trim();
        if (t && !jaConsta(g.entradas, t)) { g.entradas.push({ dia: e.dia, texto: t }); }
      });
    }
    return ordem.map(function (k) {
      var g = porChave[k];
      if (CFG.ignorarCortesia && g.entradas.length > 1) {
        var uteis = g.entradas.filter(function (e) { return !ehSoCortesia(e.texto); });
        if (uteis.length) { g.entradas = uteis; }
      }
      g.feito = g.entradas.length ? escreverHistorico(g.entradas) : '';
      return g;
    });
  }

  function organizarPlanilha() {
    return abrirPlanilha().then(function (previa) {
      var total = previa.ultima - previa.mapa.linha;
      var plano = planejarOrganizacao(previa.pasta, previa.mapa, previa.ultima);
      if (total <= 0) { registrar('Organizar: a planilha esta vazia.'); return false; }
      if (!confirm('Organizar a planilha no formato "uma linha por demanda"?\n\n' +
                   '  hoje ......... ' + total + ' linha(s)\n' +
                   '  vai ficar .... ' + plano.length + ' demanda(s)\n\n' +
                   'As linhas do mesmo ticket viram uma so, com o historico datado.\n' +
                   'Um backup e criado antes e nada e gravado se a conferencia falhar.')) {
        return false;
      }
      return editar('Planilha organizada: ' + total + ' linha(s) viraram ' +
                    plano.length + ' demanda(s).', function (ctx) {
        var c = ctx.mapa.colunas;
        var grupos = planejarOrganizacao(ctx.pasta, ctx.mapa, ctx.ultima);
        var checagens = [];
        grupos.forEach(function (g, i) {
          var l = ctx.mapa.linha + 1 + i;
          if (typeof g.data === 'number') { ctx.pasta.escreverNumero(l, c.data, g.data); }
          else if (textoDaCelula(g.data)) { ctx.pasta.escreverTexto(l, c.data, String(g.data)); }
          else { apagarCelula(ctx.pasta, l, c.data); }
          var tk = textoDaCelula(g.ticket);
          if (tk && soDigitos(tk) === tk) { ctx.pasta.escreverNumero(l, c.ticket, tk); }
          else { ctx.pasta.escreverTexto(l, c.ticket, tk); }
          ctx.pasta.escreverTexto(l, c.problema, g.problema);
          aplicarQuebra(ctx.pasta, ctx.pasta.escreverTexto(l, c.feito, g.feito));
          ctx.pasta.escreverTexto(l, c.status, g.status);
          alturaAutomatica(ctx.pasta, l);
          checagens.push({ linha: l, col: c.ticket, valor: tk });
          checagens.push({ linha: l, col: c.problema, valor: g.problema });
        });
        for (var l2 = ctx.mapa.linha + 1 + grupos.length; l2 <= ctx.ultima; l2++) {
          (function (linha) {
            CAMPOS.forEach(function (k) { apagarCelula(ctx.pasta, linha, c[k]); });
          }(l2));
        }
        return { checagens: checagens, linhasEsperadas: ctx.mapa.linha + grupos.length };
      });
    })['catch'](function (e) {
      registrar('ERRO ao organizar: ' + String((e && e.message) || e));
      return false;
    });
  }

  /* ========================================================= PAINEL ====== */
  var painel = null, corpo = null, barraBotoes = null, recolhido = false;
  var modoPainel = 'status';     /* 'status' | 'corrigir' */
  var recadoCorrecao = '';

  function estilo(el, css) { el.setAttribute('style', css); }

  /* Tema "terminal" - ciano neon sobre azul-marinho quase preto. Usado tanto
     na montagem do painel quanto nas cores de status/log em atualizarPainel(). */
  var TEMA = {
    fundo: 'linear-gradient(165deg,#060a12 0%,#0a1120 55%,#081420 100%)',
    fundoTopo: 'linear-gradient(90deg,#081018,#0c1c2c)',
    fundoBarra: '#060d16',
    fundoInput: '#050a10',
    borda: '#123449',
    bordaForte: '#1c5578',
    brilho: '0 0 0 1px rgba(0,229,255,.18),0 0 30px rgba(0,229,255,.12),0 14px 44px rgba(0,0,0,.55)',
    texto: '#bdeeff',
    textoFraco: '#5b7f92',
    acento: '#00e5ff',
    ok: '#00ffa2',
    aviso: '#ffce3d',
    erro: '#ff4d7e'
  };

  /* Injeta uma unica vez: hover/focus e a animacao do ponto "ao vivo" no
     titulo. So estilos cosmeticos do PROPRIO painel - nao toca em nada do
     Zendesk. */
  function injetarEstiloGlobal() {
    if (document.getElementById('aa-estilos')) { return; }
    var s = document.createElement('style');
    s.id = 'aa-estilos';
    s.textContent =
      '@keyframes aa-pulse{0%,100%{opacity:1}50%{opacity:.3}}' +
      '.aa-dot{width:7px;height:7px;border-radius:50%;display:inline-block;margin-right:8px;' +
        'vertical-align:1px;background:' + TEMA.ok + ';box-shadow:0 0 6px ' + TEMA.ok +
        ',0 0 16px ' + TEMA.ok + ';animation:aa-pulse 1.8s ease-in-out infinite}' +
      '.aa-btn{transition:background .15s ease,box-shadow .15s ease,color .15s ease}' +
      '.aa-btn:hover{background:#0f2f42!important;border-color:' + TEMA.acento + '!important;' +
        'box-shadow:0 0 10px rgba(0,229,255,.35),inset 0 0 8px rgba(0,229,255,.13)!important;' +
        'color:#eafeff!important}' +
      '.aa-input:focus{outline:none!important;border-color:' + TEMA.acento + '!important;' +
        'box-shadow:0 0 10px rgba(0,229,255,.3)!important}' +
      '.aa-input::placeholder{color:' + TEMA.textoFraco + '}' +
      '.aa-corpo::-webkit-scrollbar{width:6px}' +
      '.aa-corpo::-webkit-scrollbar-track{background:transparent}' +
      '.aa-corpo::-webkit-scrollbar-thumb{background:' + TEMA.borda + ';border-radius:3px}' +
      '.aa-corpo::-webkit-scrollbar-thumb:hover{background:' + TEMA.bordaForte + '}';
    document.documentElement.appendChild(s);
  }

  /* ------------------------------------------- mover e recolher ---------- */
  /* O painel nasce no canto inferior direito, mas voce arrasta pelo cabecalho
     para onde quiser e ele lembra o lugar. A posicao e sempre recolocada dentro
     da janela - na abertura e a cada redimensionamento - entao ele nunca some
     fora da tela. */
  var CH_POS = 'aa2_painel_pos';

  function dentroDaTela(x, y) {
    var maxX = Math.max(0, window.innerWidth - (painel.offsetWidth || 340));
    var maxY = Math.max(0, window.innerHeight - (painel.offsetHeight || 200));
    return { x: Math.min(Math.max(0, x), maxX), y: Math.min(Math.max(0, y), maxY) };
  }

  function posicionar(x, y, salvar) {
    var p = dentroDaTela(x, y);
    painel.style.left = p.x + 'px';
    painel.style.top = p.y + 'px';
    painel.style.right = 'auto';
    painel.style.bottom = 'auto';
    if (salvar) { gravar(CH_POS, { x: p.x, y: p.y }); }
  }

  function restaurarPosicao() {
    var p = ler(CH_POS, null);
    if (p && typeof p.x === 'number' && typeof p.y === 'number') {
      posicionar(p.x, p.y, false);
    }
  }

  function tornarArrastavel(alca) {
    estilo(alca, (alca.getAttribute('style') || '') + ';cursor:move;user-select:none');
    alca.addEventListener('mousedown', function (ev) {
      if (ev.button !== 0 || (ev.target && ev.target.getAttribute('data-aa-fixo'))) { return; }
      ev.preventDefault();
      ev.stopPropagation();
      var r = painel.getBoundingClientRect();
      var dx = ev.clientX - r.left, dy = ev.clientY - r.top;

      function mover(e) { posicionar(e.clientX - dx, e.clientY - dy, false); }
      function soltar(e) {
        document.removeEventListener('mousemove', mover, true);
        document.removeEventListener('mouseup', soltar, true);
        posicionar(e.clientX - dx, e.clientY - dy, true);
      }
      document.addEventListener('mousemove', mover, true);
      document.addEventListener('mouseup', soltar, true);
    }, true);

    window.addEventListener('resize', function () {
      if (!painel || painel.style.display === 'none' || !painel.style.left) { return; }
      posicionar(parseInt(painel.style.left, 10) || 0,
                 parseInt(painel.style.top, 10) || 0, true);
    });
  }

  /* Marca do painel: uma prancheta com um ponto que pisca enquanto a
     automacao esta viva. Desenhado em SVG pelo DOM, sem imagem externa - o
     CSP do Zendesk nao deixaria carregar uma. */
  function iconePainel() {
    var NS = 'http://www.w3.org/2000/svg';
    function no(tag, attrs) {
      var e = document.createElementNS(NS, tag);
      Object.keys(attrs).forEach(function (k) { e.setAttribute(k, attrs[k]); });
      return e;
    }
    var svg = no('svg', { viewBox: '0 0 24 24', width: '18', height: '18' });
    svg.setAttribute('style', 'vertical-align:-4px;margin-right:8px;overflow:visible;' +
      'filter:drop-shadow(0 0 4px rgba(0,229,255,.7))');
    svg.appendChild(no('rect', { x: '4', y: '3.5', width: '14', height: '17', rx: '2.5',
      fill: 'none', stroke: TEMA.acento, 'stroke-width': '1.5' }));
    svg.appendChild(no('line', { x1: '7.5', y1: '9', x2: '14.5', y2: '9',
      stroke: TEMA.acento, 'stroke-width': '1.4', 'stroke-linecap': 'round' }));
    svg.appendChild(no('line', { x1: '7.5', y1: '13', x2: '14.5', y2: '13',
      stroke: TEMA.acento, 'stroke-width': '1.4', 'stroke-linecap': 'round' }));
    var ponto = no('circle', { cx: '19.5', cy: '5', r: '2.6', fill: TEMA.ok });
    ponto.setAttribute('style', 'animation:aa-pulse 1.8s ease-in-out infinite');
    svg.appendChild(ponto);
    return svg;
  }

  function criarPainel() {
    injetarEstiloGlobal();

    painel = document.createElement('div');
    estilo(painel, 'position:fixed;right:14px;bottom:14px;z-index:2147483647;width:340px;' +
      'background:' + TEMA.fundo + ';color:' + TEMA.texto + ';' +
      'font:12px/1.6 "Cascadia Code",Consolas,Menlo,monospace;letter-spacing:.2px;' +
      'border:1px solid ' + TEMA.borda + ';border-radius:10px;box-shadow:' + TEMA.brilho +
      ';overflow:hidden');

    var topo = document.createElement('div');
    estilo(topo, 'padding:10px 12px;background:' + TEMA.fundoTopo + ';font-weight:700;' +
      'display:flex;justify-content:space-between;align-items:center;' +
      'border-bottom:1px solid ' + TEMA.borda);
    var titulo = document.createElement('span');
    estilo(titulo, 'color:' + TEMA.acento + ';text-shadow:0 0 8px rgba(0,229,255,.6);' +
      'letter-spacing:2px;font-size:11px');
    titulo.appendChild(iconePainel());
    titulo.appendChild(document.createTextNode('PAINEL V2'));
    topo.appendChild(titulo);

    var controles = document.createElement('span');
    estilo(controles, 'display:flex;gap:10px;align-items:center');
    function botaoTopo(simbolo, titulo2, aoClicar) {
      var b = document.createElement('span');
      b.textContent = simbolo;
      b.title = titulo2;
      b.setAttribute('data-aa-fixo', '1');   /* clique aqui nao arrasta o painel */
      estilo(b, 'cursor:pointer;opacity:.55;color:' + TEMA.texto + ';font-size:13px;' +
        'line-height:1');
      b.onmouseenter = function () { b.style.opacity = '1'; };
      b.onmouseleave = function () { b.style.opacity = '.55'; };
      b.onclick = function (ev) { ev.stopPropagation(); aoClicar(b); };
      controles.appendChild(b);
      return b;
    }
    botaoTopo('—', 'Recolher / expandir', function (b) {
      recolhido = !recolhido;
      /* display explicito na volta: a barra e flex, e limpar a propriedade a
         faria virar block, empilhando os botoes. */
      corpo.style.display = recolhido ? 'none' : 'block';
      barraBotoes.style.display = recolhido ? 'none' : 'flex';
      b.textContent = recolhido ? '▢' : '—';
    });
    botaoTopo('✕', 'Fechar (o favorito reabre)', function () {
      painel.style.display = 'none';
    });
    topo.appendChild(controles);
    painel.appendChild(topo);
    tornarArrastavel(topo);

    corpo = document.createElement('div');
    corpo.className = 'aa-corpo';
    estilo(corpo, 'padding:10px 12px;max-height:60vh;overflow:auto');
    painel.appendChild(corpo);

    var barra = document.createElement('div');
    barraBotoes = barra;
    estilo(barra, 'padding:9px 12px;background:' + TEMA.fundoBarra + ';display:flex;' +
      'flex-wrap:wrap;gap:6px;border-top:1px solid ' + TEMA.borda);
    [
      ['Conectar planilha', function () { escolherPasta(); }],
      ['Verificar agora', function () { varrer().then(function () { return gravarNoExcel(); }); }],
      ['Gravar', function () { gravarNoExcel(); }],
      ['Corrigir', function () { abrirCorrecao(); }],
      ['Organizar', function () { organizarPlanilha(); }],
      ['Ativar resumo', function () {
        if (modoResumo !== 'regras') { registrar('Resumo ja esta ativo.'); return; }
        if (typeof Summarizer === 'undefined' && typeof LanguageModel === 'undefined') {
          registrar('Este Chrome nao tem a IA local. Seguindo com a limpeza por regra.');
          return;
        }
        if (confirm('Ativar o resumo inteligente?\n\nO Chrome vai baixar um modelo de ' +
                    'IA (alguns GB) que roda 100% na sua maquina. Nenhum texto do ' +
                    'Zendesk sai do computador. O download acontece uma unica vez.')) {
          prepararIA(true);
        }
      }],
      ['Resumo', function () { resumoAvulso(); }],
      ['Fechar semana', function () {
        if (confirm('Arquivar a semana atual e limpar a planilha?\n\n' +
                    'Serao criados em ' + CFG.pastaSemanais + ':\n' +
                    '  - a copia da planilha (.xlsx)\n' +
                    '  - o resumo para os gestores (.pdf)\n\n' +
                    'A copia e conferida antes de qualquer limpeza.')) {
          fecharSemana(true);
        }
      }]
    ].forEach(function (par) {
      var b = document.createElement('button');
      b.textContent = par[0];
      b.className = 'aa-btn';
      estilo(b, 'cursor:pointer;background:' + TEMA.fundoTopo + ';color:' + TEMA.acento + ';' +
        'border:1px solid ' + TEMA.borda + ';padding:5px 9px;border-radius:5px;' +
        'font:11px Consolas,monospace;letter-spacing:.3px');
      b.onclick = par[1];
      barra.appendChild(b);
    });

    var entrada = document.createElement('input');
    entrada.className = 'aa-input';
    entrada.placeholder = 'Atividade avulsa: Assunto | O que foi feito | Status';
    estilo(entrada, 'width:100%;margin-top:6px;background:' + TEMA.fundoInput + ';color:' +
      TEMA.texto + ';border:1px solid ' + TEMA.borda + ';border-radius:5px;padding:6px 8px;' +
      'font:11px Consolas,monospace');
    /* stopPropagation: o Zendesk tem atalhos de tecla unica. Sem isto, digitar
       aqui dispararia acoes na tela do ticket. */
    entrada.onkeydown = function (ev) {
      ev.stopPropagation();
      if (ev.key !== 'Enter' || !entrada.value.trim()) { return; }
      var txt = entrada.value;
      entrada.disabled = true;          /* a IA local leva um instante */
      adicionarManual(txt).then(function (ok) {
        entrada.disabled = false;
        if (ok) { entrada.value = ''; gravarNoExcel(); }
        try { entrada.focus(); } catch (e) {}
      });
    };
    entrada.onkeyup = function (ev) { ev.stopPropagation(); };
    entrada.onkeypress = function (ev) { ev.stopPropagation(); };
    barra.appendChild(entrada);
    painel.appendChild(barra);

    document.documentElement.appendChild(painel);
  }

  /* --------------------------------------- tela de correcao de linhas ---- */
  var editando = false;

  function abrirCorrecao() {
    if (!corpo) { return; }
    modoPainel = 'corrigir';
    corpo.textContent = '';
    var espera = document.createElement('div');
    estilo(espera, 'color:' + TEMA.textoFraco);
    espera.textContent = 'lendo a planilha...';
    corpo.appendChild(espera);

    listarUltimasLinhas().then(function (itens) {
      desenharLista(itens);
    })['catch'](function (e) {
      modoPainel = 'status';
      registrar('ERRO ao abrir a correcao: ' + String((e && e.message) || e));
      atualizarPainel();
    });
  }

  function voltarAoStatus() {
    modoPainel = 'status';
    recadoCorrecao = '';
    atualizarPainel();
  }

  function cabecalhoCorrecao(texto) {
    var t = document.createElement('div');
    estilo(t, 'display:flex;justify-content:space-between;align-items:center;' +
      'margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid ' + TEMA.borda);
    var h = document.createElement('span');
    estilo(h, 'color:' + TEMA.acento + ';letter-spacing:1px;font-size:10px;' +
      'text-transform:uppercase');
    h.textContent = texto;
    var v = document.createElement('span');
    v.textContent = '← voltar';
    estilo(v, 'cursor:pointer;color:' + TEMA.textoFraco + ';font-size:10px');
    v.onclick = voltarAoStatus;
    t.appendChild(h);
    t.appendChild(v);
    corpo.appendChild(t);
  }

  function desenharLista(itens) {
    corpo.textContent = '';
    cabecalhoCorrecao('corrigir linha');

    if (recadoCorrecao) {
      var r = document.createElement('div');
      estilo(r, 'color:' + TEMA.ok + ';margin-bottom:6px;font-size:11px');
      r.textContent = recadoCorrecao;
      corpo.appendChild(r);
      recadoCorrecao = '';
    }

    if (!itens.length) {
      var vazio = document.createElement('div');
      estilo(vazio, 'color:' + TEMA.textoFraco);
      vazio.textContent = 'A planilha nao tem linhas preenchidas.';
      corpo.appendChild(vazio);
      return;
    }

    itens.forEach(function (it) {
      var d = document.createElement('div');
      estilo(d, 'cursor:pointer;padding:4px 6px;margin-bottom:3px;border-radius:4px;' +
        'border:1px solid transparent;white-space:nowrap;overflow:hidden;' +
        'text-overflow:ellipsis;color:' + TEMA.texto);
      d.textContent = 'L' + it.linha + '  ' + (it.data || '--/--') + '  ' +
                      (it.ticket || '-') + '  ' + (it.problema || '(sem assunto)');
      d.title = it.feito || '';
      d.onmouseenter = function () {
        d.style.background = '#0f2f42'; d.style.borderColor = TEMA.borda;
      };
      d.onmouseleave = function () {
        d.style.background = ''; d.style.borderColor = 'transparent';
      };
      d.onclick = function () { desenharFormulario(it); };
      corpo.appendChild(d);
    });

    var dica = document.createElement('div');
    estilo(dica, 'margin-top:8px;color:' + TEMA.textoFraco + ';font-size:10px;' +
      'white-space:pre-wrap');
    dica.textContent = 'Clique numa linha para corrigir o texto ou excluir.\n' +
                       'Um backup e criado antes de qualquer alteracao.';
    corpo.appendChild(dica);
  }

  /* stopPropagation em tudo: o Zendesk tem atalhos de tecla unica e sem isso
     digitar aqui dispararia acoes na tela do ticket. */
  function segurarTeclas(el) {
    el.onkeydown = function (ev) { ev.stopPropagation(); };
    el.onkeyup = function (ev) { ev.stopPropagation(); };
    el.onkeypress = function (ev) { ev.stopPropagation(); };
  }

  function campoEdicao(rotulo, valor, linhas) {
    var bloco = document.createElement('div');
    estilo(bloco, 'margin-bottom:6px');
    var r = document.createElement('div');
    estilo(r, 'color:' + TEMA.textoFraco + ';font-size:10px;letter-spacing:.5px;' +
      'text-transform:uppercase;margin-bottom:2px');
    r.textContent = rotulo;
    var el = document.createElement(linhas ? 'textarea' : 'input');
    el.className = 'aa-input';
    el.value = valor || '';
    if (linhas) { el.rows = linhas; }
    estilo(el, 'width:100%;box-sizing:border-box;background:' + TEMA.fundoInput + ';color:' +
      TEMA.texto + ';border:1px solid ' + TEMA.borda + ';border-radius:5px;padding:5px 7px;' +
      'font:11px Consolas,monospace;resize:vertical');
    segurarTeclas(el);
    bloco.appendChild(r);
    bloco.appendChild(el);
    corpo.appendChild(bloco);
    return el;
  }

  function campoStatus(valor) {
    var bloco = document.createElement('div');
    estilo(bloco, 'margin-bottom:6px');
    var r = document.createElement('div');
    estilo(r, 'color:' + TEMA.textoFraco + ';font-size:10px;letter-spacing:.5px;' +
      'text-transform:uppercase;margin-bottom:2px');
    r.textContent = 'status / andamento';
    var sel = document.createElement('select');
    sel.className = 'aa-input';
    estilo(sel, 'width:100%;box-sizing:border-box;background:' + TEMA.fundoInput + ';color:' +
      TEMA.texto + ';border:1px solid ' + TEMA.borda + ';border-radius:5px;padding:5px 7px;' +
      'font:11px Consolas,monospace');
    var opcoes = [CFG.statusResolvido, CFG.statusAndamento];
    /* Se a linha tem um status fora dos dois padroes, ele aparece na lista para
       nao ser trocado sem voce querer. */
    if (valor && opcoes.indexOf(valor) < 0) { opcoes.push(valor); }
    opcoes.forEach(function (o) {
      var op = document.createElement('option');
      op.value = o;
      op.textContent = o;
      if (o === valor) { op.selected = true; }
      sel.appendChild(op);
    });
    segurarTeclas(sel);
    bloco.appendChild(r);
    bloco.appendChild(sel);
    corpo.appendChild(bloco);
    return sel;
  }

  function desenharFormulario(it) {
    corpo.textContent = '';
    cabecalhoCorrecao('linha ' + it.linha);

    var fData = campoEdicao('data', it.data);
    var fTicket = campoEdicao('n ticket', it.ticket);
    var fProblema = campoEdicao('problema / solicitacao', it.problema, 2);
    var fFeito = campoEdicao('o que foi feito (historico)', it.feito, 6);
    var fNova = campoEdicao('nova tratativa (entra com a data de hoje)', '', 2);
    var fStatus = campoStatus(it.status);

    var acoes = document.createElement('div');
    estilo(acoes, 'display:flex;gap:6px;margin-top:8px;flex-wrap:wrap');
    corpo.appendChild(acoes);

    var botoes = [];
    function botao(texto, cor, aoClicar) {
      var b = document.createElement('button');
      b.textContent = texto;
      b.className = 'aa-btn';
      estilo(b, 'cursor:pointer;background:' + TEMA.fundoTopo + ';color:' + cor + ';' +
        'border:1px solid ' + TEMA.borda + ';padding:5px 9px;border-radius:5px;' +
        'font:11px Consolas,monospace;letter-spacing:.3px');
      b.onclick = aoClicar;
      acoes.appendChild(b);
      botoes.push(b);
      return b;
    }

    /* Enquanto grava, todos os botoes ficam travados e o que foi clicado diz o
       que esta acontecendo. Sem isso um duplo clique dispararia duas gravacoes. */
    function ocupar(btn, texto) {
      editando = true;
      btn.__rotulo = btn.textContent;
      btn.textContent = texto;
      botoes.forEach(function (b) { b.disabled = true; b.style.opacity = '.5'; });
    }
    function liberar(btn) {
      editando = false;
      btn.textContent = btn.__rotulo || btn.textContent;
      botoes.forEach(function (b) { b.disabled = false; b.style.opacity = ''; });
    }

    var bSalvar = botao('Salvar', TEMA.ok, function () {
      if (editando) { return; }
      ocupar(bSalvar, 'Salvando...');
      salvarCorrecao(it.linha, {
        data: fData.value, ticket: fTicket.value, problema: fProblema.value,
        feito: fFeito.value, novaTratativa: fNova.value, status: fStatus.value
      }).then(function (ok) {
        liberar(bSalvar);
        if (ok) {
          recadoCorrecao = 'Linha ' + it.linha + ' corrigida.';
          abrirCorrecao();
        } else {
          alert('Nao foi possivel salvar. Clique em "voltar" para ver o motivo ' +
                'na ultima mensagem do painel. A planilha nao foi alterada.');
        }
      });
    });

    var bExcluir = botao('Excluir linha', TEMA.erro, function () {
      if (editando) { return; }
      if (!confirm('Excluir a linha ' + it.linha + ' da planilha?\n\n' +
                   (it.problema || '(sem assunto)') + '\n\n' +
                   'As linhas de baixo sobem uma posicao. Um backup e criado antes ' +
                   'e nada e gravado se a conferencia falhar.')) { return; }
      ocupar(bExcluir, 'Excluindo...');
      excluirLinha(it.linha).then(function (ok) {
        liberar(bExcluir);
        if (ok) {
          recadoCorrecao = 'Linha ' + it.linha + ' excluida.';
          abrirCorrecao();
        } else {
          alert('Nao foi possivel excluir. Clique em "voltar" para ver o motivo ' +
                'na ultima mensagem do painel. A planilha nao foi alterada.');
        }
      });
    });

    botao('Cancelar', TEMA.textoFraco, function () {
      if (editando) { return; }
      abrirCorrecao();
    });
  }

  function atualizarPainel() {
    if (!corpo) { return; }
    /* Enquanto voce edita uma linha, o painel nao se redesenha por baixo: um
       log chegando no meio da digitacao apagaria o formulario. */
    if (modoPainel !== 'status') { return; }
    var a = ler(CH.atividades, {});
    var semana = semanaDe(agora());
    var daSemana = Object.keys(a).filter(function (id) {
      var dia = diaISO(new Date(a[id].quando));
      return dia >= semana.inicio && dia <= semana.fim;
    });
    var naoGravadas = daSemana.filter(function (id) { return !a[id].gravado; });
    var est = ler(CH.estado, {});

    corpo.textContent = '';
    function linha(txt, cor) {
      var d = document.createElement('div');
      d.textContent = txt;
      if (cor) { estilo(d, 'color:' + cor); }
      corpo.appendChild(d);
      return d;
    }

    linha('planilha ..... ' + (dirHandle ? 'conectada' : 'NAO conectada'),
          dirHandle ? TEMA.ok : TEMA.aviso);
    linha('semana ....... ' + semana.inicio + ' a ' + semana.fim);
    var demandas = {};
    daSemana.forEach(function (id) {
      demandas[soDigitos(a[id].ticket) ? 't' + a[id].ticket : id] = true;
    });
    linha('demandas ..... ' + Object.keys(demandas).length + ' na semana (' +
          daSemana.length + ' tratativas, ' + naoGravadas.length + ' pendentes)');
    linha('varredura .... ' + (varrendo ? 'em andamento' :
      (est.ultimaVarredura ? new Date(est.ultimaVarredura).toLocaleTimeString() : 'nunca')));
    linha('tickets ...... ' + Object.keys(ler(CH.tickets, {})).length + ' acompanhados');
    linha('alcance ...... ' + (CFG.descobrirPelaBusca && !buscaFalhou
            ? 'todas as abas (busca do Zendesk)' : 'so os tickets abertos nesta aba'),
          CFG.descobrirPelaBusca && !buscaFalhou ? TEMA.ok : TEMA.aviso);
    linha('resumo ....... ' + estadoIA + (modoResumo === 'prompt' ? ' [modelo]' : ''),
          modoResumo !== 'regras' ? TEMA.ok : TEMA.textoFraco);

    var sep = document.createElement('div');
    estilo(sep, 'margin:8px 0 5px;border-top:1px solid ' + TEMA.borda + ';padding-top:6px;' +
      'opacity:.7;color:' + TEMA.textoFraco + ';letter-spacing:1px;font-size:10px;' +
      'text-transform:uppercase');
    sep.textContent = 'ultimas mensagens';
    corpo.appendChild(sep);

    LOG.slice(-8).forEach(function (l) {
      var d = linha(l, l.indexOf('ERRO') > 0 ? TEMA.erro :
        (l.indexOf('AVISO') > 0 ? TEMA.aviso : TEMA.textoFraco));
      estilo(d, (d.getAttribute('style') || '') + ';white-space:pre-wrap;font-size:11px');
    });
  }

  function mostrarPainel() {
    if (!painel) { criarPainel(); restaurarPosicao(); }
    painel.style.display = '';
    /* Reabrir pelo favorito sempre volta para a tela de status: uma lista de
       correcao velha na tela levaria a editar a linha errada. */
    if (!editando) { modoPainel = 'status'; }
    /* A janela pode ter mudado de tamanho enquanto o painel estava fechado. */
    if (painel.style.left) {
      posicionar(parseInt(painel.style.left, 10) || 0,
                 parseInt(painel.style.top, 10) || 0, false);
    }
    atualizarPainel();
  }

  /* ========================================================= INICIO ====== */
  function iniciar() {
    var herdou = herdarEstadoAntigo();
    mostrarPainel();
    registrar('Automacao armada nesta aba (v2: uma linha por demanda).');
    if (herdou) {
      registrar('Historico de tickets da versao anterior aproveitado: nada sera ' +
                'lancado em dobro.');
    }
    vigiarNavegacao();

    prepararIA(false);   /* so verifica; nunca baixa nada sem voce mandar */

    souLider();
    conectarPasta(false).then(function (ok) {
      if (ok) { registrar('Pasta reconectada automaticamente.'); }
      else { registrar('Clique em "Conectar planilha" para autorizar o acesso.'); }
      atualizarPainel();
      return varrer();
    }).then(function () {
      return gravarNoExcel();
    }).then(function () {
      return fecharSemana(false);
    })['catch'](function (e) { registrar('ERRO na inicializacao: ' + e.message); });

    var avisouReserva = false;
    setInterval(function () { if (!avisouReserva) { souLider(); } }, 30000);   /* batimento */
    setInterval(function () {
      if (!souLider()) {
        if (!avisouReserva) {
          registrar('Outra aba ja esta com o painel ativo e cuida da captura. Esta fica de reserva.');
        }
        avisouReserva = true;
        return;
      }
      if (avisouReserva) { registrar('A outra aba saiu: esta assumiu a captura.'); }
      avisouReserva = false;
      varrer().then(function (novas) {
        if (novas || pendentes().length) { return gravarNoExcel(); }
      });
    }, CFG.intervaloVarreduraSeg * 1000);
    window.addEventListener('beforeunload', function () {
      var l = ler('aa2_lider', null);
      if (l && l.dono === ID_ABA) { gravar('aa2_lider', null); }
    });
  }

  window.__automacaoAtividades2 = {
    mostrarPainel: mostrarPainel,
    varrer: varrer,
    gravar: gravarNoExcel,
    fechar: fecharSemana,
    corrigir: abrirCorrecao,
    organizar: organizarPlanilha,
    resumo: resumoAvulso,
    cfg: CFG,
    /* Porta de diagnostico: deixa conferir as regras de texto pelo console
       (F12) sem tocar na planilha. Nenhuma delas grava nada. */
    _regras: {
      tirarNarrador: tirarNarrador,
      limparSaidaIA: limparSaidaIA,
      limparComentario: limparComentario,
      limparMarkdown: limparMarkdown,
      ehSoCortesia: ehSoCortesia,
      lerHistorico: lerHistorico,
      escreverHistorico: escreverHistorico,
      agruparDemandas: agruparDemandas,
      planejarOrganizacao: planejarOrganizacao,
      ufDoAssunto: ufDoAssunto,
      pareceIngles: pareceIngles,
      interpretarAvulsa: interpretarAvulsa,
      perdeuConteudo: perdeuConteudo,
      dataDigitada: dataDigitada,
      definirNome: function (n) { meuNome = n; }
    }
  };

  iniciar();
})();
