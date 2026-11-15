/* ==========================================================================
 * pdf_relatorio.js
 *
 * Gera PDF SEM nenhuma biblioteca externa. Mesmo motivo do xlsx_cirurgico.js:
 * o CSP do Zendesk nao deixa carregar CDN, e nada pode ser instalado.
 *
 * O PDF e escrito byte a byte. Usa as fontes padrao do formato (Helvetica),
 * que todo leitor de PDF ja tem embutido - por isso o arquivo nao precisa
 * carregar fonte nenhuma e abre igual em qualquer maquina. A acentuacao
 * portuguesa vem do WinAnsiEncoding.
 *
 * DUAS CAMADAS
 *   Documento  - primitivas: texto, retangulo, linha, imagem, paginas.
 *   Relatorio  - a identidade visual do relatorio: faixa de topo, rotulo de destaque,
 *                titulo, secao, tabela, bullets, painel e rodape. Quebra de
 *                pagina automatica.
 * ========================================================================== */

var PDFS = (function () {
  'use strict';

  var A4 = { largura: 595.28, altura: 841.89 };

  /* ===================================================== PALETA DO RELATORIO ==
     Paleta do relatorio. Para mudar a cor de destaque do documento, troque
     SO a linha "destaque" e tudo no documento acompanha. */
  var COR = {
    destaque:   [0.122, 0.435, 0.922],   /* #1F6FEB - cor de destaque */
    grafite:    [0.247, 0.247, 0.247],   /* #3F3F3F - cabecalho de tabela */
    texto:      [0.106, 0.106, 0.106],   /* #1B1B1B - corpo */
    azulAco:    [0.278, 0.416, 0.541],   /* #476A8A - linha de apoio */
    cinza:      [0.467, 0.467, 0.467],   /* #777777 - secundario */
    borda:      [0.800, 0.800, 0.800],   /* #CCCCCC - grade da tabela */
    bordaClara: [0.898, 0.898, 0.898],   /* #E5E5E5 */
    fundoPainel:[0.945, 0.965, 0.996],   /* #F1F6FE - painel de situacao */
    zebra:      [0.976, 0.976, 0.976],   /* #F9F9F9 */
    verde:      [0.118, 0.490, 0.196],   /* #1E7D32 - concluido */
    branco:     [1, 1, 1]
  };

  /* ======================================================== WinAnsiEncoding
     Para 0x20-0x7E e 0xA0-0xFF o byte e o proprio code point. A faixa
     0x80-0x9F e onde o WinAnsi difere do Latin-1: ali moram os tipograficos
     (travessao, aspas curvas, reticencias). Sem este mapa, um travessao
     colado de um e-mail viraria lixo no PDF. */
  var WINANSI = {
    0x20AC: 0x80, 0x201A: 0x82, 0x0192: 0x83, 0x201E: 0x84, 0x2026: 0x85,
    0x2020: 0x86, 0x2021: 0x87, 0x02C6: 0x88, 0x2030: 0x89, 0x0160: 0x8A,
    0x2039: 0x8B, 0x0152: 0x8C, 0x017D: 0x8E, 0x2018: 0x91, 0x2019: 0x92,
    0x201C: 0x93, 0x201D: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97,
    0x02DC: 0x98, 0x2122: 0x99, 0x0161: 0x9A, 0x203A: 0x9B, 0x0153: 0x9C,
    0x017E: 0x9E, 0x0178: 0x9F
  };

  /* Texto -> bytes WinAnsi, ja com escape dos caracteres que a sintaxe do PDF
     reserva. Caractere que nao existe no WinAnsi vira '?': perder um simbolo
     exotico e melhor do que gerar um PDF que nao abre. */
  function literalPdf(texto) {
    var s = String(texto === null || texto === undefined ? '' : texto);
    var saida = '';
    for (var i = 0; i < s.length; i++) {
      var cp = s.codePointAt(i);
      if (cp > 0xFFFF) { i++; saida += '?'; continue; }
      var b;
      if (cp === 0x0A) { b = 0x20; }                      /* quebra vira espaco */
      else if (cp < 0x20) { b = 0x20; }
      else if (cp <= 0x7E) { b = cp; }
      else if (WINANSI[cp] !== undefined) { b = WINANSI[cp]; }
      else if (cp >= 0xA0 && cp <= 0xFF) { b = cp; }
      else { b = 0x3F; }                                   /* '?' */
      if (b === 0x28 || b === 0x29 || b === 0x5C) { saida += '\\'; }
      saida += String.fromCharCode(b);
    }
    return saida;
  }

  /* =================================================== MEDIDA DE TEXTO =====
     A largura de um texto em Helvetica e a mesma em Arial - as duas familias
     compartilham as metricas para o alfabeto latino. Entao em vez de embutir a
     tabela de larguras do formato, a medicao usa o proprio canvas do Chrome,
     que ja tem Arial. Medir em 'px' e usar o mesmo numero em pontos funciona
     porque a conta e identica: largura = (avanco/1000) x corpo. */
  var ctxMedida = null;

  function contexto() {
    if (!ctxMedida) {
      var c = document.createElement('canvas');
      c.width = 10; c.height = 10;
      ctxMedida = c.getContext('2d');
    }
    return ctxMedida;
  }

  function medir(texto, tamanho, negrito, espacamento) {
    var s = String(texto || '');
    if (!s) { return 0; }
    var ctx = contexto();
    ctx.font = (negrito ? 'bold ' : '') + tamanho + 'px Arial, Helvetica, sans-serif';
    var w = ctx.measureText(s).width;
    /* Tc do PDF soma o espacamento depois de CADA glifo, inclusive o ultimo. */
    return w + (espacamento || 0) * s.length;
  }

  /* ======================================================= DOCUMENTO =======
     Montador de baixo nivel. Guarda os objetos numerados e escreve a tabela
     xref no fim - e a xref com deslocamentos errados que faz um PDF "nao
     abrir", entao ela e calculada a partir do tamanho real ja escrito. */
  function Documento() {
    this.paginas = [];
    this.imagens = [];          /* {apelido, bytes, largura, altura} */
    this.larguraPagina = A4.largura;
    this.alturaPagina = A4.altura;
  }

  Documento.prototype.pagina = function () {
    var pg = { ops: [] };
    this.paginas.push(pg);
    return pg;
  };

  function corPdf(cor, preenche) {
    var c = cor || COR.texto;
    return c[0].toFixed(3) + ' ' + c[1].toFixed(3) + ' ' + c[2].toFixed(3) +
           ' ' + (preenche ? 'rg' : 'RG');
  }

  Documento.prototype.retangulo = function (pg, x, y, larg, alt, cor) {
    pg.ops.push(corPdf(cor, true));
    pg.ops.push(x.toFixed(2) + ' ' + y.toFixed(2) + ' ' +
                larg.toFixed(2) + ' ' + alt.toFixed(2) + ' re f');
  };

  Documento.prototype.linha = function (pg, x1, y1, x2, y2, cor, espessura) {
    pg.ops.push(corPdf(cor, false));
    pg.ops.push((espessura || 0.5).toFixed(2) + ' w');
    pg.ops.push(x1.toFixed(2) + ' ' + y1.toFixed(2) + ' m ' +
                x2.toFixed(2) + ' ' + y2.toFixed(2) + ' l S');
  };

  /* y e a LINHA DE BASE do texto, nao o topo. */
  Documento.prototype.texto = function (pg, x, y, txt, op) {
    var o = op || {};
    var fonte = o.negrito ? '/F2' : (o.italico ? '/F3' : '/F1');
    pg.ops.push('BT');
    pg.ops.push(corPdf(o.cor, true));
    pg.ops.push(fonte + ' ' + (o.tamanho || 9).toFixed(2) + ' Tf');
    if (o.espacamento) { pg.ops.push(o.espacamento.toFixed(3) + ' Tc'); }
    pg.ops.push(x.toFixed(2) + ' ' + y.toFixed(2) + ' Td');
    pg.ops.push('(' + literalPdf(txt) + ') Tj');
    if (o.espacamento) { pg.ops.push('0 Tc'); }
    pg.ops.push('ET');
  };

  Documento.prototype.registrarImagem = function (apelido, bytes, largura, altura) {
    this.imagens.push({ apelido: apelido, bytes: bytes, largura: largura, altura: altura });
  };

  Documento.prototype.desenharImagem = function (pg, apelido, x, y, larg, alt) {
    pg.ops.push('q');
    pg.ops.push(larg.toFixed(2) + ' 0 0 ' + alt.toFixed(2) + ' ' +
                x.toFixed(2) + ' ' + y.toFixed(2) + ' cm');
    pg.ops.push('/' + apelido + ' Do');
    pg.ops.push('Q');
  };

  /* ---------------------------------------------------------- serializacao */
  function bytesDeLatin1(texto) {
    var b = new Uint8Array(texto.length);
    for (var i = 0; i < texto.length; i++) { b[i] = texto.charCodeAt(i) & 0xFF; }
    return b;
  }

  Documento.prototype.gerar = function () {
    var self = this;
    var objetos = [];            /* objetos[i] = string OU {cabecalho, dados} */

    function novo(conteudo) { objetos.push(conteudo); return objetos.length; }

    /* 1 catalogo, 2 paginas, 3-5 fontes: numeros fixos para simplificar. */
    var idCatalogo = novo(null);
    var idPaginas = novo(null);
    var idF1 = novo('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica ' +
                    '/Encoding /WinAnsiEncoding >>');
    var idF2 = novo('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold ' +
                    '/Encoding /WinAnsiEncoding >>');
    var idF3 = novo('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Oblique ' +
                    '/Encoding /WinAnsiEncoding >>');

    var recursosImagem = '';
    this.imagens.forEach(function (img) {
      var id = novo({
        cabecalho: '<< /Type /XObject /Subtype /Image /Width ' + img.largura +
                   ' /Height ' + img.altura + ' /ColorSpace /DeviceRGB ' +
                   '/BitsPerComponent 8 /Filter /DCTDecode /Length ' +
                   img.bytes.length + ' >>',
        dados: img.bytes
      });
      recursosImagem += '/' + img.apelido + ' ' + id + ' 0 R ';
    });

    var recursos = '<< /Font << /F1 ' + idF1 + ' 0 R /F2 ' + idF2 + ' 0 R /F3 ' +
                   idF3 + ' 0 R >>' +
                   (recursosImagem ? ' /XObject << ' + recursosImagem + '>>' : '') + ' >>';

    var idsPagina = [];
    this.paginas.forEach(function (pg) {
      var fluxo = pg.ops.join('\n');
      var idFluxo = novo({
        cabecalho: '<< /Length ' + fluxo.length + ' >>',
        dados: bytesDeLatin1(fluxo)
      });
      var idPg = novo('<< /Type /Page /Parent ' + idPaginas + ' 0 R /MediaBox ' +
                      '[0 0 ' + self.larguraPagina.toFixed(2) + ' ' +
                      self.alturaPagina.toFixed(2) + '] /Resources ' + recursos +
                      ' /Contents ' + idFluxo + ' 0 R >>');
      idsPagina.push(idPg);
    });

    objetos[idCatalogo - 1] = '<< /Type /Catalog /Pages ' + idPaginas + ' 0 R >>';
    objetos[idPaginas - 1] = '<< /Type /Pages /Count ' + idsPagina.length +
      ' /Kids [' + idsPagina.map(function (i) { return i + ' 0 R'; }).join(' ') + '] >>';

    /* ------- escrita, guardando o deslocamento real de cada objeto -------- */
    var pedacos = [];
    var posicao = 0;
    var deslocamentos = [];

    function escrever(algo) {
      var b = (algo instanceof Uint8Array) ? algo : bytesDeLatin1(algo);
      pedacos.push(b);
      posicao += b.length;
    }

    escrever('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');

    objetos.forEach(function (obj, i) {
      deslocamentos[i] = posicao;
      escrever((i + 1) + ' 0 obj\n');
      if (typeof obj === 'string') {
        escrever(obj + '\nendobj\n');
      } else {
        escrever(obj.cabecalho + '\nstream\n');
        escrever(obj.dados);
        escrever('\nendstream\nendobj\n');
      }
    });

    var inicioXref = posicao;
    var xref = 'xref\n0 ' + (objetos.length + 1) + '\n0000000000 65535 f \n';
    for (var i = 0; i < objetos.length; i++) {
      xref += ('0000000000' + deslocamentos[i]).slice(-10) + ' 00000 n \n';
    }
    escrever(xref);
    escrever('trailer\n<< /Size ' + (objetos.length + 1) + ' /Root ' +
             idCatalogo + ' 0 R >>\nstartxref\n' + inicioXref + '\n%%EOF\n');

    var total = pedacos.reduce(function (s, p) { return s + p.length; }, 0);
    var saida = new Uint8Array(total);
    var off = 0;
    pedacos.forEach(function (p) { saida.set(p, off); off += p.length; });
    return saida;
  };

  /* =================================================== TEXTO COM NEGRITO ===
     Marcacao minima: *assim* fica em negrito. Serve para destacar nomes e
     numeros dentro de bullets e celulas, como no modelo. */
  function separarRuns(texto) {
    var partes = String(texto || '').split('*');
    var runs = [];
    for (var i = 0; i < partes.length; i++) {
      if (partes[i] === '') { continue; }
      runs.push({ t: partes[i], negrito: (i % 2) === 1 });
    }
    return runs.length ? runs : [{ t: '', negrito: false }];
  }

  /* Quebra em linhas respeitando a largura. Palavra maior que a linha inteira
     e cortada na forca - melhor do que estourar a coluna. */
  function quebrar(texto, largura, tamanho, espacamento) {
    var runs = separarRuns(texto);
    var linhas = [];
    var atual = [];
    var larguraAtual = 0;

    function fecharLinha() {
      linhas.push(atual);
      atual = [];
      larguraAtual = 0;
    }

    runs.forEach(function (run) {
      var palavras = run.t.split(/(\s+)/);
      palavras.forEach(function (palavra) {
        if (palavra === '') { return; }
        var w = medir(palavra, tamanho, run.negrito, espacamento);

        if (/^\s+$/.test(palavra)) {
          if (larguraAtual === 0) { return; }        /* nao comeca linha com espaco */
          atual.push({ t: palavra, negrito: run.negrito, w: w });
          larguraAtual += w;
          return;
        }

        if (larguraAtual + w > largura && larguraAtual > 0) { fecharLinha(); }

        while (w > largura && palavra.length > 1) {   /* palavra gigante */
          var corte = palavra.length;
          while (corte > 1 && medir(palavra.slice(0, corte), tamanho, run.negrito,
                                    espacamento) > largura) { corte--; }
          var pedaco = palavra.slice(0, corte);
          atual.push({ t: pedaco, negrito: run.negrito,
                       w: medir(pedaco, tamanho, run.negrito, espacamento) });
          fecharLinha();
          palavra = palavra.slice(corte);
          w = medir(palavra, tamanho, run.negrito, espacamento);
        }

        if (palavra) {
          atual.push({ t: palavra, negrito: run.negrito, w: w });
          larguraAtual += w;
        }
      });
    });

    if (atual.length) { fecharLinha(); }
    return linhas.length ? linhas : [[]];
  }

  /* ======================================================== RELATORIO ======
     A identidade visual do modelo "Resumo de demandas", componente a
     componente, com quebra de pagina automatica. */
  function Relatorio(opcoes) {
    var o = opcoes || {};
    this.doc = new Documento();
    this.margemX = 52;
    /* Sem logo nem selo o topo encolhe: nao faz sentido reservar 116pt para uma
       faixa vazia. Assim o documento sai bem proporcionado antes de o logo
       existir, e se ajusta sozinho quando ele aparecer na pasta. */
    this.topo = (o.logo || o.selo) ? 116 : 74;
    this.baixo = 62;            /* espaco reservado para o rodape */
    this.largura = A4.largura - this.margemX * 2;
    this.pg = null;
    this.y = 0;
    this.logo = o.logo || null;         /* {apelido, largura, altura} */
    this.selo = o.selo || null;
    this.rodapeTexto = o.rodape || '';
    this.numeroPagina = 0;
    this.paginasParaNumerar = [];
  }

  Relatorio.prototype.novaPagina = function () {
    this.pg = this.doc.pagina();
    this.numeroPagina++;
    this.desenharTopo();
    this.paginasParaNumerar.push({ pg: this.pg, n: this.numeroPagina });
    this.y = A4.altura - this.topo;
    return this.pg;
  };

  /* Faixa de topo: logo a esquerda, selo a direita, filete claro embaixo. */
  Relatorio.prototype.desenharTopo = function () {
    var yTopo = A4.altura - 40;
    if (this.logo) {
      var hL = 34;
      var wL = hL * (this.logo.largura / this.logo.altura);
      this.doc.desenharImagem(this.pg, this.logo.apelido, this.margemX, yTopo - hL, wL, hL);
    }
    if (this.selo) {
      var hS = 30;
      var wS = hS * (this.selo.largura / this.selo.altura);
      this.doc.desenharImagem(this.pg, this.selo.apelido,
        A4.largura - this.margemX - wS, yTopo - hS + 2, wS, hS);
    }
    var yFilete = A4.altura - this.topo + 24;
    this.doc.linha(this.pg, this.margemX, yFilete,
                   A4.largura - this.margemX, yFilete, COR.bordaClara, 0.8);
  };

  /* Paragrafo de corpo, texto normal. Aceita a marcacao *negrito*. */
  Relatorio.prototype.paragrafo = function (texto, opcoes) {
    var o = opcoes || {};
    var tamanho = o.tamanho || 9.5;
    var alturaLinha = tamanho * 1.45;
    var linhas = quebrar(texto, this.largura, tamanho, 0);
    var self = this;
    linhas.forEach(function (linha) {
      self.garantir(alturaLinha);
      self.y -= alturaLinha;
      self.escreverLinha(linha, self.margemX, self.y, tamanho, o.cor || COR.texto, 0);
    });
    this.y -= 8;
  };

  Relatorio.prototype.garantir = function (altura) {
    if (!this.pg || this.y - altura < this.baixo) { this.novaPagina(); }
  };

  Relatorio.prototype.pular = function (h) { this.y -= h; };

  /* Rotulo de destaque em caixa alta, bem espacado - o "RESUMO DE DEMANDAS". */
  Relatorio.prototype.rotulo = function (texto) {
    this.garantir(22);
    this.y -= 9;
    this.doc.texto(this.pg, this.margemX, this.y, String(texto).toUpperCase(),
      { tamanho: 8, negrito: true, cor: COR.destaque, espacamento: 2.1 });
    this.y -= 6;
  };

  Relatorio.prototype.titulo = function (texto) {
    var linhas = quebrar(texto, this.largura, 23, 0);
    this.garantir(linhas.length * 28 + 16);
    var self = this;
    linhas.forEach(function (linha) {
      self.y -= 25;
      self.escreverLinha(linha, self.margemX, self.y, 23, COR.texto, 0);
    });
    this.y -= 10;
    this.doc.linha(this.pg, this.margemX, this.y,
                   A4.largura - this.margemX, this.y, COR.destaque, 1.6);
    this.y -= 4;
  };

  /* Paragrafo de apoio em azul-aco, com barra clara a esquerda. */
  Relatorio.prototype.apoio = function (texto) {
    var linhas = quebrar(texto, this.largura - 12, 9.5, 0);
    this.garantir(linhas.length * 13 + 14);
    var yInicio = this.y - 4;
    var self = this;
    linhas.forEach(function (linha) {
      self.y -= 13;
      self.escreverLinha(linha, self.margemX + 12, self.y, 9.5, COR.azulAco, 0);
    });
    this.doc.retangulo(this.pg, this.margemX, this.y - 3, 2.5,
                       yInicio - this.y + 3, COR.bordaClara);
    this.y -= 10;
  };

  /* Cabecalho de secao: cor de destaque, caixa alta, com regua embaixo. */
  Relatorio.prototype.secao = function (texto) {
    this.garantir(34);
    this.y -= 16;
    this.doc.texto(this.pg, this.margemX, this.y, String(texto).toUpperCase(),
      { tamanho: 9.5, negrito: true, cor: COR.destaque, espacamento: 1.3 });
    this.y -= 5;
    this.doc.linha(this.pg, this.margemX, this.y,
                   A4.largura - this.margemX, this.y, COR.destaque, 1.1);
    this.y -= 9;
  };

  /* Encurta o texto ate caber na largura, terminando em reticencias. */
  function encaixar(texto, largura, tamanho, negrito) {
    var s = String(texto || '');
    if (medir(s, tamanho, negrito, 0) <= largura) { return s; }
    while (s.length > 1 && medir(s + '…', tamanho, negrito, 0) > largura) {
      s = s.slice(0, -1);
    }
    return s.replace(/[\s,;.-]+$/, '') + '…';
  }

  /* Faixa cheia na cor de destaque - abre cada demanda no detalhamento. O titulo e o
     subtitulo sao encurtados para nunca vazar a faixa: texto branco saindo
     para fora da faixa fica ilegivel. */
  Relatorio.prototype.faixa = function (rotulo, titulo, subtitulo) {
    var altura = subtitulo ? 62 : 50;
    var util = this.largura - 32;
    this.garantir(altura + 14);
    this.y -= altura;
    this.doc.retangulo(this.pg, this.margemX, this.y, this.largura, altura, COR.destaque);
    var yTexto = this.y + altura - 17;
    if (rotulo) {
      this.doc.texto(this.pg, this.margemX + 16, yTexto,
        encaixar(String(rotulo).toUpperCase(), util, 7.5, true),
        { tamanho: 7.5, negrito: true, cor: COR.branco, espacamento: 2 });
      yTexto -= 21;
    }
    this.doc.texto(this.pg, this.margemX + 16, yTexto, encaixar(titulo, util, 19, true),
      { tamanho: 19, negrito: true, cor: COR.branco });
    if (subtitulo) {
      this.doc.texto(this.pg, this.margemX + 16, yTexto - 14,
        encaixar(subtitulo, util, 9, false), { tamanho: 9, cor: COR.branco });
    }
    this.y -= 12;
  };

  Relatorio.prototype.escreverLinha = function (runs, x, y, tamanho, cor, espacamento) {
    var cx = x;
    var self = this;
    runs.forEach(function (r) {
      if (r.t.trim() !== '') {
        self.doc.texto(self.pg, cx, y, r.t, {
          tamanho: tamanho, negrito: r.negrito, cor: cor, espacamento: espacamento
        });
      }
      cx += r.w;
    });
  };

  /* Lista com marcador de destaque. */
  Relatorio.prototype.bullets = function (itens, opcoes) {
    var o = opcoes || {};
    var tamanho = o.tamanho || 9;
    var recuo = 14;
    var self = this;
    itens.forEach(function (item) {
      var linhas = quebrar(item, self.largura - recuo - (o.recuoExtra || 0), tamanho, 0);
      self.garantir(linhas.length * 12.5 + 5);
      linhas.forEach(function (linha, i) {
        self.y -= 12.5;
        if (i === 0) {
          self.doc.texto(self.pg, self.margemX + (o.recuoExtra || 0) + 3, self.y,
            '•', { tamanho: tamanho + 1, negrito: true, cor: COR.destaque });
        }
        self.escreverLinha(linha, self.margemX + recuo + (o.recuoExtra || 0),
                           self.y, tamanho, o.cor || COR.texto, 0);
      });
      self.y -= 3.5;
    });
  };

  /* Painel claro com barra de destaque - a "SITUACAO ATUAL" do modelo. */
  Relatorio.prototype.painel = function (itens) {
    var tamanho = 9;
    var recuo = 16;
    var larguraTexto = this.largura - recuo - 18;
    var self = this;

    /* Mede tudo antes para nao partir o painel no meio da pagina. */
    var blocos = itens.map(function (i) {
      return quebrar(i, larguraTexto, tamanho, 0);
    });
    var altura = blocos.reduce(function (s, b) { return s + b.length * 12.5 + 4; }, 0) + 16;
    this.garantir(altura + 8);

    var yTopo = this.y;
    this.doc.retangulo(this.pg, this.margemX, yTopo - altura, this.largura, altura, COR.fundoPainel);
    this.doc.retangulo(this.pg, this.margemX, yTopo - altura, 3, altura, COR.destaque);

    this.y -= 8;
    blocos.forEach(function (linhas) {
      linhas.forEach(function (linha, i) {
        self.y -= 12.5;
        if (i === 0) {
          self.doc.texto(self.pg, self.margemX + recuo - 9, self.y, '•',
            { tamanho: tamanho + 1, negrito: true, cor: COR.destaque });
        }
        self.escreverLinha(linha, self.margemX + recuo, self.y, tamanho, COR.texto, 0);
      });
      self.y -= 4;
    });
    this.y = yTopo - altura - 10;
  };

  /* Tabela com cabecalho grafite. `larguras` sao proporcoes que somam 1.
     Uma linha alta demais para caber na pagina e quebrada: o cabecalho se
     repete na pagina seguinte, para a tabela nunca virar um bloco orfao. */
  Relatorio.prototype.tabela = function (cabecalhos, proporcoes, linhas, opcoes) {
    var o = opcoes || {};
    var tamanho = o.tamanho || 8.5;
    var pad = 6;
    var self = this;
    var cols = proporcoes.map(function (p) { return p * self.largura; });

    function desenharCabecalho() {
      self.garantir(40);
      self.y -= 19;
      self.doc.retangulo(self.pg, self.margemX, self.y, self.largura, 19, COR.grafite);
      var x = self.margemX;
      cabecalhos.forEach(function (c, i) {
        self.doc.texto(self.pg, x + pad, self.y + 6.5, c,
          { tamanho: tamanho, negrito: true, cor: COR.branco });
        x += cols[i];
      });
    }

    desenharCabecalho();

    linhas.forEach(function (linha, idx) {
      var celulas = linha.map(function (c, i) {
        return quebrar(c, cols[i] - pad * 2, tamanho, 0);
      });
      var alturaLinha = Math.max.apply(null, celulas.map(function (c) {
        return c.length * 11.5;
      })) + pad * 2;

      if (self.y - alturaLinha < self.baixo) {
        self.novaPagina();
        desenharCabecalho();
      }

      self.y -= alturaLinha;
      if (o.zebra && idx % 2 === 1) {
        self.doc.retangulo(self.pg, self.margemX, self.y, self.largura,
                           alturaLinha, COR.zebra);
      }

      var x = self.margemX;
      celulas.forEach(function (linhasCelula, i) {
        var yc = self.y + alturaLinha - pad - 8;
        var cor = (o.colunaDestaque === i) ? COR.destaque : COR.texto;
        var negritoCol = (o.colunaDestaque === i) || (o.colunasNegrito || []).indexOf(i) >= 0;
        linhasCelula.forEach(function (lc) {
          if (negritoCol) { lc.forEach(function (r) { r.negrito = true; }); }
          self.escreverLinha(lc, x + pad, yc, tamanho, cor, 0);
          yc -= 11.5;
        });
        /* grade vertical */
        self.doc.linha(self.pg, x, self.y, x, self.y + alturaLinha, COR.borda, 0.5);
        x += cols[i];
      });
      self.doc.linha(self.pg, x, self.y, x, self.y + alturaLinha, COR.borda, 0.5);
      self.doc.linha(self.pg, self.margemX, self.y, self.margemX + self.largura,
                     self.y, COR.borda, 0.5);
    });

    this.y -= 12;
  };

  /* Quadro de indicadores: numero grande e rotulo embaixo, lado a lado.
     itens = [{ valor, rotulo, cor }]. */
  Relatorio.prototype.indicadores = function (itens) {
    var altura = 50, vao = 8;
    var larg = (this.largura - vao * (itens.length - 1)) / itens.length;
    this.garantir(altura + 12);
    this.y -= altura;
    var self = this;
    itens.forEach(function (it, i) {
      var x = self.margemX + i * (larg + vao);
      var cor = it.cor || COR.destaque;
      self.doc.retangulo(self.pg, x, self.y, larg, altura, COR.zebra);
      self.doc.retangulo(self.pg, x, self.y + altura - 2.5, larg, 2.5, cor);
      self.doc.texto(self.pg, x + 11, self.y + 20, String(it.valor),
        { tamanho: 21, negrito: true, cor: COR.texto });
      self.doc.texto(self.pg, x + 11, self.y + 8, String(it.rotulo).toUpperCase(),
        { tamanho: 6.8, negrito: true, cor: COR.cinza, espacamento: 0.9 });
    });
    this.y -= 12;
  };

  /* Ficha de uma demanda no detalhamento:
       #ticket  Titulo em negrito (quebra em quantas linhas precisar)
       periodo . N tratativas . STATUS
       | sintese em azul-aco (opcional)
       dd/mm   tratativa...
       dd/mm   tratativa...
     O cabecalho da ficha nunca fica sozinho no pe da pagina; o historico pode
     continuar na pagina seguinte, linha a linha. */
  Relatorio.prototype.demanda = function (d) {
    var self = this;
    var wTicket = d.ticket ? medir(d.ticket, 10, true, 0) + 9 : 0;
    var titulo = quebrarNegrito(String(d.titulo || '').replace(/\*/g, ''),
                            this.largura - wTicket, 10);

    this.garantir(titulo.length * 13 + 13 + 34);
    this.y -= 6;
    titulo.forEach(function (linha, i) {
      self.y -= 13;
      if (i === 0 && d.ticket) {
        self.doc.texto(self.pg, self.margemX, self.y, d.ticket,
          { tamanho: 10, negrito: true, cor: COR.destaque });
      }
      self.escreverLinha(linha, self.margemX + wTicket, self.y, 10, COR.texto, 0);
    });

    this.y -= 12;
    var meta = String(d.meta || '');
    this.doc.texto(this.pg, this.margemX + wTicket, this.y, meta, { tamanho: 8, cor: COR.cinza });
    if (d.status) {
      var xs = this.margemX + wTicket + medir(meta + '   ·   ', 8, false, 0);
      this.doc.texto(this.pg, this.margemX + wTicket + medir(meta, 8, false, 0), this.y,
        '   ·   ', { tamanho: 8, cor: COR.cinza });
      this.doc.texto(this.pg, xs, this.y, String(d.status).toUpperCase(),
        { tamanho: 7.5, negrito: true, cor: d.corStatus || COR.azulAco, espacamento: 0.6 });
    }
    this.y -= 4;

    if (d.sintese) {
      var ls = quebrar(d.sintese, this.largura - wTicket - 10, 9, 0);
      this.garantir(ls.length * 12.5 + 8);
      this.y -= 3;
      var yIni = this.y;
      ls.forEach(function (linha) {
        self.y -= 12.5;
        self.escreverLinha(linha, self.margemX + wTicket + 10, self.y, 9, COR.azulAco, 0);
      });
      this.doc.retangulo(this.pg, this.margemX + wTicket, this.y - 3, 2.2,
                         yIni - this.y + 1, COR.azulAco);
      this.y -= 5;
    }

    var entradas = d.entradas || [];
    var comData = entradas.length > 1 || (entradas.length === 1 && d.sintese);
    var wDia = comData ? 34 : 0;
    entradas.forEach(function (e) {
      var linhas = quebrar(e.texto, self.largura - wTicket - wDia, 8.8, 0);
      linhas.forEach(function (linha, i) {
        self.garantir(12);
        self.y -= 12;
        if (i === 0 && comData && e.dia) {
          self.doc.texto(self.pg, self.margemX + wTicket, self.y, e.dia,
            { tamanho: 8.2, negrito: true, cor: COR.cinza });
        }
        self.escreverLinha(linha, self.margemX + wTicket + wDia, self.y, 8.8, COR.texto, 0);
      });
      self.y -= 2.5;
    });
    if (!entradas.length) {
      this.y -= 12;
      this.doc.texto(this.pg, this.margemX + wTicket, this.y,
        'Sem tratativa descrita na planilha.', { tamanho: 8.5, italico: true, cor: COR.cinza });
    }

    this.y -= 7;
    this.doc.linha(this.pg, this.margemX, this.y, A4.largura - this.margemX, this.y,
                   COR.bordaClara, 0.6);
    this.y -= 3;
  };

  /* quebrar() mede em peso normal; titulo em negrito e mais largo. */
  function quebrarNegrito(texto, largura, tamanho) {
    return quebrar('*' + texto + '*', largura, tamanho, 0);
  }

  /* Regua + texto centralizado, como o "FIM DO TICKET" do modelo. */
  Relatorio.prototype.encerramento = function (texto, complemento) {
    this.garantir(40);
    this.y -= 14;
    this.doc.linha(this.pg, this.margemX, this.y,
                   A4.largura - this.margemX, this.y, COR.destaque, 1.1);
    this.y -= 15;
    var t = String(texto).toUpperCase();
    var w = medir(t, 9, true, 1.8);
    var wc = complemento ? medir('   |   ' + complemento, 8.5, false, 0) : 0;
    var x = (A4.largura - (w + wc)) / 2;
    this.doc.texto(this.pg, x, this.y, t,
      { tamanho: 9, negrito: true, cor: COR.destaque, espacamento: 1.8 });
    if (complemento) {
      this.doc.texto(this.pg, x + w, this.y, '   |   ' + complemento,
        { tamanho: 8.5, cor: COR.cinza });
    }
    this.y -= 8;
    this.doc.linha(this.pg, this.margemX, this.y,
                   A4.largura - this.margemX, this.y, COR.destaque, 1.1);
    this.y -= 10;
  };

  /* Rodape so no fim: a numeracao precisa saber o total de paginas. */
  Relatorio.prototype.fecharRodapes = function () {
    var total = this.paginasParaNumerar.length;
    var self = this;
    this.paginasParaNumerar.forEach(function (p) {
      self.doc.linha(p.pg, self.margemX, 46, A4.largura - self.margemX, 46,
                     COR.bordaClara, 0.8);
      if (self.rodapeTexto) {
        self.doc.texto(p.pg, self.margemX, 34, self.rodapeTexto,
          { tamanho: 7.5, cor: COR.cinza });
      }
      var num = p.n + ' de ' + total;
      self.doc.texto(p.pg, A4.largura - self.margemX - medir(num, 7.5, false, 0), 34,
        num, { tamanho: 7.5, cor: COR.cinza });
    });
  };

  Relatorio.prototype.gerar = function () {
    this.fecharRodapes();
    return this.doc.gerar();
  };

  /* ==================================================== IMAGEM -> JPEG =====
     O PDF recebe a imagem como JPEG (DCTDecode), que e o formato que entra
     direto, sem recodificar. Passar pelo canvas resolve duas coisas de uma
     vez: aceita PNG de entrada e achata a transparencia sobre branco - sem
     isso, um logo PNG com fundo transparente sairia sobre preto. */
  function imagemParaJpeg(blobOuArquivo) {
    return createImageBitmap(blobOuArquivo).then(function (bmp) {
      var largura = bmp.width, altura = bmp.height;
      var max = 900;
      if (largura > max) { altura = Math.round(altura * max / largura); largura = max; }
      var c = document.createElement('canvas');
      c.width = largura; c.height = altura;
      var ctx = c.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, largura, altura);
      ctx.drawImage(bmp, 0, 0, largura, altura);
      try { bmp.close(); } catch (e) {}

      var url = c.toDataURL('image/jpeg', 0.92);
      var base64 = url.slice(url.indexOf(',') + 1);
      var bin = atob(base64);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) { bytes[i] = bin.charCodeAt(i); }
      return { bytes: bytes, largura: largura, altura: altura };
    });
  }

  /* ==================================================== CONFERENCIA =======
     O erro classico de um gerador de PDF escrito a mao e a tabela xref apontar
     para o lugar errado: o leitor abre e diz apenas "arquivo danificado", sem
     nenhuma pista. Entao, antes de os bytes irem para o disco, a xref e relida
     e cada deslocamento tem que cair exatamente em "N 0 obj". Mesma disciplina
     da gravacao do .xlsx: conferir primeiro, gravar depois. */
  function conferir(bytes) {
    var txt = new TextDecoder('latin1').decode(bytes);

    if (txt.indexOf('%PDF-') !== 0) { throw new Error('falta o cabecalho %PDF'); }

    var posStart = txt.lastIndexOf('startxref');
    if (posStart < 0) { throw new Error('falta o startxref'); }
    var inicio = parseInt(txt.slice(posStart + 9).trim(), 10);
    if (!(inicio > 0) || inicio >= bytes.length) {
      throw new Error('o startxref aponta para fora do arquivo');
    }

    var m = /^xref\s*\n0 (\d+)\s*\n/.exec(txt.slice(inicio));
    if (!m) { throw new Error('a xref nao esta onde o startxref aponta'); }

    var total = parseInt(m[1], 10);
    var corpo = txt.slice(inicio + m[0].length);
    for (var n = 1; n < total; n++) {
      var entrada = corpo.slice(n * 20, (n + 1) * 20);
      if (entrada.length < 18) { throw new Error('xref truncada no objeto ' + n); }
      var off = parseInt(entrada.slice(0, 10), 10);
      var esperado = n + ' 0 obj';
      if (txt.slice(off, off + esperado.length) !== esperado) {
        throw new Error('o objeto ' + n + ' nao esta no deslocamento que a xref informa');
      }
    }

    if (txt.lastIndexOf('%%EOF') < inicio) { throw new Error('falta o %%EOF'); }
    if (txt.indexOf('/Type /Page') < 0) { throw new Error('nenhuma pagina no documento'); }
    return total - 1;
  }

  return {
    Documento: Documento,
    Relatorio: Relatorio,
    COR: COR,
    medir: medir,
    quebrar: quebrar,
    literalPdf: literalPdf,
    imagemParaJpeg: imagemParaJpeg,
    conferir: conferir,
    A4: A4
  };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = PDFS; }
