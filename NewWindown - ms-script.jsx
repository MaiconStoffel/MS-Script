(function () {
  if (app.documents.length === 0) {
    alert("Abra um documento antes de rodar o script.");
    return;
  }

  var doc = app.activeDocument;

  var AspasAbre = "“«„";
  var AspasFecha = "”“»";
  var aspasAbertura = "";
  var aspasFechamento = "";
  var aplicarAspas = false;

  var ignorarTabelasBarra = false;
  var ajustarEspacamentoDoisPontos_Antes = false;
  var ajustarBarraEspaco = false;
  var ajustarVirgulaEspaco = false;
  var ajustarIgualEspaco = false;
  var ajustarEspacamentoDoisPontos = false;
  var ajustarPontoFinalEspaco = false;
  var ajustarEspacoAntesPercent = false;
  var ajustarEspacamentoPontoVirgula = false;
  var idiomaEscolhido = "";

  //Capitalizacao
  var idiomaSelecionado = null; //eh diferente do idiomaEscolhido
  var capitalizacaoSelecionada = null;
  var estilosSelecionados = [];
  var contadorCapitalizacao = 0;

  // Variáveis contagem
  var contadorAspas = 0;
  var contadorPontoVirgula = 0;
  var contadorVirgulaPonto = 0;
  var contadorVirgulaEspaco = 0;
  var contadorBarraEspaco = 0;
  var contadorIgualEspaco = 0;
  var contadorEspacamentoDoisPontos = 0;
  var contadorPontoFinalEspaco = 0;
  var contadorEspacoAntesPercent = 0;
  var contadorEspacoDepoisPontoVirgula = 0;

  //variaveis da troca do decimal
  var trocarPontoPorVirgula = false;
  var trocarVirgulaPorPonto = false;
  var paginasSelecionadasDecimal = ""; // var q conteem o que o usuario escreveu na troca decimal, porém mais filtrada e arrumada
  var aplicarEmTodasAsPaginas = true; // padrao pra aplicar a mudança decimal em todas as pgns
  var paginasArray = []; // array q conteem as paginas que o usuario digitou
  var inputPaginas; // var pra receber o que o usuario escreveu nas pgns da troca decimal

  function TelaInicial() {
    var janela = new Window("dialog", "");

    var titulo = janela.add(
      "statictext",
      undefined,
      "Por favor, selecione uma das funções abaixo:"
    );
    titulo.alignment = "center";

    var grupoPrincipal = janela.add("group");
    grupoPrincipal.orientation = "row";
    grupoPrincipal.alignChildren = "top";

    var coluna1 = grupoPrincipal.add("group");
    coluna1.orientation = "column";
    coluna1.alignChildren = "fill";

    var coluna2 = grupoPrincipal.add("group");
    coluna2.orientation = "column";
    coluna2.alignChildren = "fill";

    var tamanhoBotao = [180, 40];

    // Função para criar botão com texto simples e tooltip
    function criarBotao(grupo, texto, onClick, helpTip) {
      var botao = grupo.add("button", undefined, texto);
      botao.preferredSize = tamanhoBotao;
      botao.helpTip = helpTip || "";
      botao.onClick = function () {
        janela.close();
        onClick();
      };
      return botao;
    }

    criarBotao(
      coluna1,
      "⇄ Troca de Aspas ⇄",
      function () {
        telaAspas();
      },
      "Altera as aspas do documento com base no idioma selecionado"
    );

    criarBotao(
      coluna1,
      "↑ Alteração de Caixa ↑",
      function () {
        telaCapitalizacaoTitulos();
      },
      "Altera a Caixa de títulos com base no idioma e estilos de parágrafo selecionados"
    );

    criarBotao(
      coluna2,
      "∞ Separador Decimal ∞",
      function () {
        telaDecimais();
      },
      "Troca do separador decimal (ponto/vírgula)"
    );

    criarBotao(
      coluna2,
      " ✇ Ajustes Gerais ✇",
      function () {
        telaGerais();
      },
      "Ajustes gerais de espaçamento para pontuação e símbolos"
    );

    janela.add(
      "statictext",
      undefined,
      "Developed by: @mshcwertz - Version 2.5"
    );

    janela.center();
    janela.show();
  }

  function telaAspas() {
    var win = new Window("dialog", "Escolha de Aspas");
    win.orientation = "column";
    win.alignChildren = "left";

    var grupoTexto = win.add("group");
    grupoTexto.alignment = "center";
    grupoTexto.add(
      "statictext",
      undefined,
      "Escolha o idioma para a alteração das aspas:"
    );

    var idiomaGrupo = win.add("group");
    idiomaGrupo.orientation = "column";
    idiomaGrupo.alignChildren = "left";

    var idiomaPolones = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Polonês („ ”)"
    );
    var idiomaAlemao = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Alemão („ “)"
    );
    var idiomaFrances = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Francês (« »)"
    );
    var idiomaRusso = idiomaGrupo.add("radiobutton", undefined, "Russo (« »)");
    var idiomaPadrao = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Padrão (“ ”)"
    );
    var idiomaOutro = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Não Efetuar Alteração"
    );
    idiomaOutro.value = true;

    var botoes = win.add("group");
    botoes.alignment = "center";
    var voltar = botoes.add("button", undefined, "Voltar");
    var proximo = botoes.add("button", undefined, "Próximo", { name: "ok" });
    var cancelar = botoes.add("button", undefined, "Cancelar", {
      name: "cancel",
    });

    voltar.onClick = function () {
      win.close();
      TelaInicial();
    };

    cancelar.onClick = function () {
      win.close();
    };

    proximo.onClick = function () {
      if (!idiomaOutro.value) {
        aplicarAspas = true;
        if (idiomaPolones.value) {
          aspasAbertura = "„";
          aspasFechamento = "”";
          idiomaEscolhido = "Polonês („ ”)";
        }
        if (idiomaAlemao.value) {
          aspasAbertura = "„";
          aspasFechamento = "“";
          idiomaEscolhido = "Alemão („ “)";
        }
        if (idiomaFrances.value) {
          aspasAbertura = "« ";
          aspasFechamento = " »";
          idiomaEscolhido = "Francês (« »)";
        }
        if (idiomaRusso.value) {
          aspasAbertura = "«";
          aspasFechamento = "»";
          idiomaEscolhido = "Russo (« »)";
        }
        if (idiomaPadrao.value) {
          aspasAbertura = "“";
          aspasFechamento = "”";
          idiomaEscolhido = "Padrão (“ ”)";
        }
      }
      win.close();
      telaFinal();
    };

    win.center();
    win.show();
  }

  function telaDecimais() {
    var win = new Window("dialog", "Troca de Decimais");
    win.orientation = "column";
    win.alignChildren = "center";

    win.add("statictext", undefined, "Escolha o tipo de troca decimal:");

    // Grupo de radiobuttons
    var grupoDecimal = win.add("group");
    grupoDecimal.orientation = "column";
    grupoDecimal.alignChildren = "left";

    var txtPgnsCenter = win.add("group");
    txtPgnsCenter.orientation = "column";
    txtPgnsCenter.alignChildren = "center";

    var rbPontoVirgula = grupoDecimal.add(
      "radiobutton",
      undefined,
      "Trocar ponto por vírgula decimal (3.14 → 3,14)"
    );
    var rbVirgulaPonto = grupoDecimal.add(
      "radiobutton",
      undefined,
      "Trocar vírgula por ponto decimal (3,14 → 3.14)"
    );
    var rbNenhum = grupoDecimal.add(
      "radiobutton",
      undefined,
      "Não Efetuar Alteração"
    );
    rbNenhum.value = true;

    var chkTodasPaginas = grupoDecimal.add(
      "checkbox",
      undefined,
      "Aplicar a todas as páginas"
    );
    chkTodasPaginas.value = true;

    txtPgnsCenter.add(
      "statictext",
      undefined,
      "Páginas a afetar (ex: 1 - 3 , 5 , 8):"
    );
    inputPaginas = win.add("edittext", undefined, "");
    inputPaginas.characters = 30;
    inputPaginas.enabled = false;

    // Quando o checkbox muda, habilita/desabilita o campo
    chkTodasPaginas.onClick = function () {
      inputPaginas.enabled = !chkTodasPaginas.value;
    };

    // Botões
    var botoes = win.add("group");
    botoes.alignment = "center";
    var voltar = botoes.add("button", undefined, "Voltar");
    var proximo = botoes.add("button", undefined, "Próximo");
    var cancelar = botoes.add("button", undefined, "Cancelar", {
      name: "cancel",
    });

    cancelar.onClick = function () {
      win.close();
    };

    voltar.onClick = function () {
      win.close();
      TelaInicial();
    };

    proximo.onClick = function () {
      trocarPontoPorVirgula = rbPontoVirgula.value;
      trocarVirgulaPorPonto = rbVirgulaPonto.value;
      aplicarEmTodasAsPaginas = chkTodasPaginas.value;

      if (!aplicarEmTodasAsPaginas) {
        paginasSelecionadasDecimal = inputPaginas.text
          ? trim(inputPaginas.text)
          : "";

        if (
          !paginasSelecionadasDecimal ||
          paginasSelecionadasDecimal.length === 0
        ) {
          alert(
            "Por favor, digite a(s) página(s) a ser(em) alterada(s) ou marque 'Aplicar a todas as páginas'.",
            "Um Problema foi Detectado",
            true
          );
          return;
        }

        var totalDePaginas = app.activeDocument.pages.length;

        paginasArray = parsePaginas(paginasSelecionadasDecimal, totalDePaginas);
        if (paginasArray.length === 0) {
          alert(
            "Nenhuma página válida foi detectada. Verifique o formato.",
            "Um Problema foi Detectado",
            true
          );
          return;
        }

        paginasArray = validarPaginasExistentes(paginasArray, totalDePaginas);

        if (paginasArray.length === 0) {
          alert(
            "Verifique se você preencheu corretamente as páginas que deseja fazer a mudança",
            "Um Problema foi Detectado",
            true
          );
          return;
        }
      } else {
        paginasSelecionadasDecimal = "";
        paginasArray = [];
      }

      win.close();
      telaFinal();
    };

    win.center();
    win.show();
  }

  function telaGerais() {
    var win = new Window("dialog", "Ajustes Gerais");
    win.orientation = "column";
    win.alignChildren = "left";

    var grupoTexto = win.add("group");
    grupoTexto.alignment = "center";
    grupoTexto.add(
      "statictext",
      undefined,
      "Escolha os ajustes de espaços para pontuação e símbolos:"
    );

    var chkSelecionarTodos = win.add(
      "checkbox",
      undefined,
      "Selecionar Todas as Padrozinações"
    );

    var chkPontoFinalEspaco = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Ponto → '.  '"
    );
    var chkVirgulaEspaco = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento da Vírgula → ',  '."
    );
    var chkIgualEspaco = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Igual → '  =  '"
    );

    var grupo2P = win.add("group");
    grupo2P.orientation = "row";
    grupo2P.alignChildren = "left";
    var chkDoisPontosEspaco = grupo2P.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Dois-Pontos → ':  '"
    );
    var chk2PEspacoAntes = grupo2P.add("checkbox", undefined, "Espaço Antes?");
    chk2PEspacoAntes.helpTip =
      "Se marcado, o espaçamento será aplicado também antes do Dois-Pontos, como em '  :  '";

    var chkEspacoDepoisPontoVirgula = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Ponto e Vírgula → ';  '"
    );
    var chkEspacoAntesPercent = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento da Porcentagem → '  %  '"
    );

    var grupoBarra = win.add("group");
    grupoBarra.orientation = "row";
    grupoBarra.alignChildren = "left";
    var chkBarraEspaco = grupoBarra.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento da Barra → '  /  '"
    );
    var chkTabelas = grupoBarra.add("checkbox", undefined, "Ignorar Tabelas?");
    chkTabelas.helpTip =
      "Se marcado, a padronização da barra '  /  ' será aplicada apenas fora de tabelas.";

    var checkboxes = [
      chkPontoFinalEspaco,
      chkVirgulaEspaco,
      chkIgualEspaco,
      chkDoisPontosEspaco,
      chkEspacoDepoisPontoVirgula,
      chkEspacoAntesPercent,
      chkBarraEspaco,
    ];

    // Evento: quando "Selecionar Todos" for clicado
    chkSelecionarTodos.onClick = function () {
      for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].value = chkSelecionarTodos.value;
      }
    };

    // Evento: se algum dos checkboxes for desmarcado manualmente, desmarca o "Selecionar Todos"
    for (var i = 0; i < checkboxes.length; i++) {
      checkboxes[i].onClick = function () {
        if (!this.value && chkSelecionarTodos.value) {
          chkSelecionarTodos.value = false;
        }
      };
    }

    var botoes = win.add("group");
    botoes.alignment = "center";
    var voltar = botoes.add("button", undefined, "Voltar");
    var proximo = botoes.add("button", undefined, "Próximo", { name: "ok" });
    var cancelar = botoes.add("button", undefined, "Cancelar", {
      name: "cancel",
    });

    cancelar.onClick = function () {
      win.close();
    };

    voltar.onClick = function () {
      win.close();
      TelaInicial();
    };

    proximo.onClick = function () {
      ajustarVirgulaEspaco = chkVirgulaEspaco.value;
      ajustarIgualEspaco = chkIgualEspaco.value;
      ajustarBarraEspaco = chkBarraEspaco.value;
      ignorarTabelasBarra = chkTabelas.value;
      ajustarEspacamentoDoisPontos = chkDoisPontosEspaco.value;
      ajustarEspacamentoDoisPontos_Antes = chk2PEspacoAntes.value;
      ajustarPontoFinalEspaco = chkPontoFinalEspaco.value;
      ajustarEspacoAntesPercent = chkEspacoAntesPercent.value;
      ajustarEspacamentoPontoVirgula = chkEspacoDepoisPontoVirgula.value;

      win.close();
      telaFinal();
    };

    win.center();
    win.show();
  }

  function telaCapitalizacaoTitulos() {
    var win = new Window("dialog", "Alteração de Caixa");
    win.orientation = "column";
    win.alignChildren = "center";

    win.add(
      "statictext",
      undefined,
      "Escolha o idioma e os estilos de parágrafo para aplicar a Alteração de Caixa:"
    );

    // Checkbox para ignorar Alteração de Caixa
    var chkNaoAplicar = win.add(
      "checkbox",
      undefined,
      "Não aplicar alteração de caixa de títulos"
    );
    chkNaoAplicar.alignment = "center";
    chkNaoAplicar.value = true;

    // Dropdown de idiomas
    var idiomas = ["Português", "Inglês", "Polonês", "Francês"];
    var grupoIdioma = win.add("group");
    grupoIdioma.add("statictext", undefined, "Idioma:");
    var dropdownIdioma = grupoIdioma.add("dropdownlist", undefined, idiomas);
    dropdownIdioma.selection = 0;

    // Dropdown de estilos
    var grupoCapitalizacao = win.add("group");
    grupoCapitalizacao.orientation = "row";
    grupoCapitalizacao.add("statictext", undefined, "Alterar a Caixa para  →");
    var dropdownCapitalizacao = grupoCapitalizacao.add(
      "dropdownlist",
      undefined,
      ["Maiúsculas", "Minúsculas", "Title Case"]
    );
    dropdownCapitalizacao.selection = 2; // default em Title Case

    // Estilos de parágrafo
    var listaEstilos = getAllParagraphStyles(doc);

    var grupoEstilos = win.add("panel", undefined, "");
    grupoEstilos.orientation = "column";
    grupoEstilos.alignChildren = "fill";

    // Grupo para centralizar o título
    var tituloGrupo = grupoEstilos.add("group");
    tituloGrupo.alignment = "center";
    tituloGrupo.add(
      "statictext",
      undefined,
      "    Selecione o(s) estilos de parágrafo para a mudança:    ."
    );

    // Grupo com scroll para os estilos
    var scrollGroup = grupoEstilos.add("group");
    scrollGroup.orientation = "column";
    scrollGroup.alignChildren = "fill";
    scrollGroup.preferredSize = [300, 150];

    var listaEstilosBox = scrollGroup.add("listbox", undefined, [], {
      multiselect: true,
    });
    listaEstilosBox.preferredSize = [300, 150];

    for (var i = 0; i < listaEstilos.length; i++) {
      if (listaEstilos[i] !== "[Parágrafo padrão]") {
        listaEstilosBox.add("item", listaEstilos[i]);
      }
    }

    //faz com q o usuario n consiga mudar as opcoes até desmarcar a opcao padrao
    dropdownCapitalizacao.enabled = !chkNaoAplicar.value;
    dropdownIdioma.enabled = !chkNaoAplicar.value;
    listaEstilosBox.enabled = !chkNaoAplicar.value;

    chkNaoAplicar.onClick = function () {
      var ativo = !chkNaoAplicar.value;
      dropdownCapitalizacao.enabled = ativo;
      dropdownIdioma.enabled = ativo;
      listaEstilosBox.enabled = ativo;
    };

    // Botões
    var botoes = win.add("group");
    botoes.alignment = "center";
    var voltar = botoes.add("button", undefined, "Voltar");
    var adicionarSigla = botoes.add("button", undefined, "Adicionar Sigla");
    var proximo = botoes.add("button", undefined, "Próximo", { name: "ok" });
    var cancelar = botoes.add("button", undefined, "Cancelar", {
      name: "cancel",
    });

    cancelar.onClick = function () {
      win.close();
    };
    voltar.onClick = function () {
      win.close();
      TelaInicial();
    };

    adicionarSigla.onClick = function () {
      abrirJanelaAdicionarSigla();
    };

    proximo.onClick = function () {
      if (chkNaoAplicar.value) {
        idiomaSelecionado = null;
        capitalizacaoSelecionada = null;
        estilosSelecionados = [];
        win.close();
        telaFinal();
        return;
      }

      if (!dropdownIdioma.selection) {
        alert("Selecione um idioma para continuar.");
        return;
      }

      if (
        !listaEstilosBox.selection ||
        listaEstilosBox.selection.length === 0
      ) {
        alert(
          "Selecione ao menos um estilo ou marque 'Não aplicar alteração de caixa de títulos'."
        );
        return;
      }

      idiomaSelecionado = dropdownIdioma.selection.text;
      capitalizacaoSelecionada = dropdownCapitalizacao.selection.text;
      estilosSelecionados = [];

      for (var j = 0; j < listaEstilosBox.selection.length; j++) {
        estilosSelecionados.push(listaEstilosBox.selection[j].text);
      }

      win.close();
      telaFinal();
    };

    win.center();
    win.show();
  }

  function telaFinal() {
    var win = new Window("dialog", "Confirmação Final ⌛");
    win.orientation = "column";
    win.alignChildren = "center";

    var resumo = "           As seguintes ações serão aplicadas:\n\n";
    if (aplicarAspas)
      resumo +=
        " -  Correção das aspas para o idioma: " + idiomaEscolhido + "\n";
    if (trocarPontoPorVirgula)
      resumo += "\n -  Trocar decimal, ponto por vírgula (3.14 → 3,14)\n";
    if (trocarVirgulaPorPonto)
      resumo += "\n -  Trocar decimal, vírgula por ponto (3,14 → 3.14)\n";

    /*---------------------------------------------------------------------------------*/

    if (ajustarPontoFinalEspaco)
      resumo += "\n -  Padronizar espaçamento do Ponto '.  '\n";
    if (ajustarVirgulaEspaco)
      resumo += "\n -  Padronizar espaçamento da Vírgula ',  '\n";
    if (ajustarIgualEspaco)
      resumo += "\n -  Padronizar espaçamento do Igual '  =  '\n";
    if (ajustarEspacamentoDoisPontos)
      resumo += "\n -  Padronizar espaçamento do Dois-Pontos ':  '\n";
    if (ajustarEspacamentoPontoVirgula)
      resumo += "\n -  Padronizar espaçamento do Ponto e Vírgula ';  '\n";
    if (ajustarEspacoAntesPercent)
      resumo += "\n -  Padronizar espaçamento da Porcentagem '  %  '\n";
    if (ajustarBarraEspaco)
      resumo += "\n -  Padronizar espaçamento da Barra '  /  '\n";

    /*---------------------------------------------------------------------------------*/

    if (idiomaSelecionado && estilosSelecionados.length > 0) {
      resumo +=
        "\n -  Alteração de Caixa:  • Modelo: " +
        capitalizacaoSelecionada +
        "\n\t\t      • Idioma: " +
        idiomaSelecionado +
        "\n";
      resumo += " -  Estilos selecionados:";
      for (var i = 0; i < estilosSelecionados.length; i++) {
        resumo += "  • " + estilosSelecionados[i] + "\n";
      }
    }

    /*---------------------------------------------------------------------------------*/

    if (
      !aplicarAspas &&
      !trocarPontoPorVirgula &&
      !trocarVirgulaPorPonto &&
      !ajustarVirgulaEspaco &&
      !ajustarIgualEspaco &&
      !ajustarBarraEspaco &&
      !ajustarEspacamentoDoisPontos &&
      !ajustarPontoFinalEspaco &&
      !ajustarEspacoAntesPercent &&
      !ajustarEspacamentoPontoVirgula &&
      !(idiomaSelecionado && estilosSelecionados.length > 0)
    ) {
      resumo += "- Nenhuma ação selecionada.\n";
    }

    var textoResumo = win.add("edittext", undefined, resumo, {
      multiline: true,
      readonly: true,
    });
    textoResumo.preferredSize.width = 300;
    textoResumo.preferredSize.height = 180;

    var botoes = win.add("group");
    botoes.alignment = "center";
    var voltar = botoes.add("button", undefined, "Voltar");
    var executar = botoes.add("button", undefined, "Executar", { name: "ok" });
    var cancelar = botoes.add("button", undefined, "Cancelar", {
      name: "cancel",
    });

    cancelar.onClick = function () {
      win.close();
    };

    voltar.onClick = function () {
      win.close();
      TelaInicial();
    };

    executar.onClick = function () {
      win.close();
      app.doScript(
        executarCorrecao,
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "Correções Projeto TRI"
      );
    };

    win.center();
    win.show();
  }

  function executarCorrecao() {
    if (aplicarAspas) {
      removerEspacoDepoisAspaAbertura();

      removerEspacoAntesAspaFechamento();

      contadorAspasAberto = corrigirAspasAbertura(aspasAbertura);

      contadorAspasFechado = corrigirAspasFechamento(aspasFechamento);

      contadorAspas = contadorAspasAberto + contadorAspasFechado;
    }

    if (trocarPontoPorVirgula)
      contadorPontoVirgula = substituirPontoPorVirgula();
    if (trocarVirgulaPorPonto)
      contadorVirgulaPonto = substituirVirgulaPorPonto();

    if (ajustarVirgulaEspaco) {
      var contagemV1 = removerEspacoAntesVirgula();
      var contagemV2 = corrigirEspacoDepoisVirgula();
      contadorVirgulaEspaco = contagemV1 + contagemV2;
    }

    if (ajustarBarraEspaco) contadorBarraEspaco = ajustarBarraIgual();

    if (ajustarIgualEspaco) contadorIgualEspaco = ajustarEspacoIgual();

    if (ajustarPontoFinalEspaco) {
      var contagemPF1 = removerEspacoAntesPontoFinal();
      var contagemPF2 = corrigirEspacoDepoisPontoFinal();
      contadorPontoFinalEspaco = contagemPF1 + contagemPF2;
    }

    if (ajustarEspacoAntesPercent)
      contadorEspacoAntesPercent = ajustarEspacoAntesPercentual();

    if (ajustarEspacamentoDoisPontos) {
      var contagemDP1 = removerEspacoAntesDoisPontos();
      var contagemDP2 = corrigirEspacoDepoisDoisPontos();
      contadorEspacamentoDoisPontos = contagemDP1 + contagemDP2;
    }

    if (ajustarEspacamentoPontoVirgula) {
      var contagemPV1 = removerEspacoAntesPontoVirgula();
      var contagemPV2 = corrigirEspacoDepoisPontoVirgula();
      contadorEspacoDepoisPontoVirgula = contagemPV1 + contagemPV2;
    }

    if (estilosSelecionados && estilosSelecionados.length > 0) {
      contadorCapitalizacao = aplicarCapitalizacaoComGREP(
        estilosSelecionados,
        idiomaSelecionado,
        capitalizacaoSelecionada
      );
    }

    var mensagem = "\n                    Correções concluídas!\n\n";
    if (aplicarAspas)
      mensagem +=
        "✅ " +
        contadorAspas +
        " - Aspas alteradas (" +
        contadorAspasAberto +
        " de abertura e " +
        contadorAspasFechado +
        " de fechamento).\n";

    if (trocarPontoPorVirgula)
      mensagem +=
        "✅ " + contadorPontoVirgula + " - Pontos trocados por vírgula.\n";
    if (trocarVirgulaPorPonto)
      mensagem +=
        "✅ " + contadorVirgulaPonto + " - Vírgulas trocadas por ponto.\n";

    if (ajustarPontoFinalEspaco)
      mensagem +=
        "✅ " +
        contadorPontoFinalEspaco +
        " - Ajustes em '   .   ' executados com sucesso.\n";
    if (ajustarVirgulaEspaco)
      mensagem +=
        "✅ " +
        contadorVirgulaEspaco +
        " - Ajustes em '   ,   ' executados com sucesso.\n";
    if (ajustarIgualEspaco)
      mensagem +=
        "✅ " +
        contadorIgualEspaco +
        " - Ajustes em '  =  ' executados com sucesso.\n";
    if (ajustarEspacamentoDoisPontos)
      mensagem +=
        "✅ " +
        contadorEspacamentoDoisPontos +
        " - Ajustes em '   :   ' executados com sucesso.\n";
    if (ajustarEspacamentoPontoVirgula)
      mensagem +=
        "✅ " +
        contadorEspacoDepoisPontoVirgula +
        " - Ajustes em '   ;   ' executados com sucesso. \n";
    if (ajustarEspacoAntesPercent)
      mensagem +=
        "✅ " +
        contadorEspacoAntesPercent +
        " - Ajustes em '  %  ' executados com sucesso.\n";
    if (ajustarBarraEspaco)
      mensagem +=
        "✅ " +
        contadorBarraEspaco +
        " - Ajustes em '   /   ' executados com sucesso.\n";

    if (contadorCapitalizacao > 0) {
      mensagem +=
        "✅ " +
        contadorCapitalizacao +
        " - Alterações de Caixa concluídas: \n       -  Idioma: " +
        idiomaSelecionado +
        ", \n       -  Modelo: " +
        capitalizacaoSelecionada +
        ", \n       -  Estilos: " +
        estilosSelecionados.join(", ") +
        "\n";
    }

    if (
      !aplicarAspas &&
      !trocarPontoPorVirgula &&
      !trocarVirgulaPorPonto &&
      !ajustarVirgulaEspaco &&
      !ajustarBarraEspaco &&
      !ajustarIgualEspaco &&
      !ajustarEspacamentoDoisPontos &&
      !ajustarPontoFinalEspaco &&
      !ajustarEspacoAntesPercent &&
      !ajustarEspacamentoPontoVirgula &&
      contadorCapitalizacao === 0
    ) {
      mensagem += "- Nenhuma alteração feita.\n";
    }
    telaResultadoFinal(mensagem);
  }

  function telaResultadoFinal(resumoExecucao) {
    var win = new Window("dialog", "Resultado Final! ☑");
    win.orientation = "column";
    win.alignChildren = "center";

    win.add("statictext", undefined, "Execução concluída com sucesso!");

    var textoResumo = win.add("edittext", undefined, resumoExecucao, {
      multiline: true,
      readonly: true,
    });
    textoResumo.preferredSize.width = 300;
    textoResumo.preferredSize.height = 180;

    var botoes = win.add("group");
    botoes.alignment = "center";
    var ok = botoes.add("button", undefined, "OK", { name: "ok" });

    ok.onClick = function () {
      win.close();
    };

    win.center();
    win.show();
  }

  //FUNCOES DE CORRECAO DE ASPAS

  //    Foi necessario adicionarmos funcoes para retirar os espacos antes e depois das aspas,
  // pois quando não havia essas funcoes acontecia o erro de duplicacao das aspas de abertura.

  function removerEspacoDepoisAspaAbertura() {
    app.findGrepPreferences = app.changeGrepPreferences = null;

    // Inclui espaços comuns (\x20), inseparáveis (\xA0) e estreitos inseparáveis (\x{202F})
    app.findGrepPreferences.findWhat =
      "(?<![A-Za-zÀ-ÖØ-öø-ÿ0-9])([" +
      AspasAbre +
      "])[\x20\xA0\u202F]+(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])";
    app.changeGrepPreferences.changeTo = "$1";

    var changed = doc.changeGrep();
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return changed ? changed.length : 0;
  }

  function removerEspacoAntesAspaFechamento() {
    var totalChanges = 0;
    var changes;

    do {
      app.findGrepPreferences = app.changeGrepPreferences = null;

      app.findGrepPreferences.findWhat =
        "([A-Za-zÀ-ÖØ-öø-ÿ0-9])\\s+([" +
        AspasFecha +
        "])(?=(\\s+[A-Za-zÀ-ÖØ-öø-ÿ])|\\s*$|\\s*[.,;:!?)]|$)";
      app.changeGrepPreferences.changeTo = "$1$2";

      var result = doc.changeGrep();
      changes = result ? result.length : 0;
      totalChanges += changes;

      app.findGrepPreferences = app.changeGrepPreferences = null;
    } while (changes > 0);

    return totalChanges;
  }

  function corrigirAspasAbertura(aspasAbertura) {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    var totalAlteracoes = 0;

    var listaAspas = AspasAbre.replace(aspasAbertura, "");
    var aspasAberturaEscapada = aspasAbertura.replace(
      /([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])/g,
      "\\$1"
    );

    app.findGrepPreferences.findWhat =
      "(?<=[\\s\\(\\[\\{—])[" + listaAspas + "]";
    app.changeGrepPreferences.changeTo = aspasAberturaEscapada;

    var changed = doc.changeGrep();
    totalAlteracoes = changed ? changed.length : 0;

    app.findGrepPreferences = app.changeGrepPreferences = null;
    return totalAlteracoes;
  }

  function corrigirAspasFechamento(aspasFechamento) {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    var totalAlteracoes = 0;

    var listaAspas = AspasFecha.replace(aspasFechamento, "");
    var aspasFechamentoEscapada = aspasFechamento.replace(
      /([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])/g,
      "\\$1"
    );

    app.findGrepPreferences.findWhat =
      "(?<![\\s\\(\\[\\{—])[" + listaAspas + "]";
    app.changeGrepPreferences.changeTo = aspasFechamentoEscapada;

    var changed = doc.changeGrep();
    totalAlteracoes = changed ? changed.length : 0;

    app.findGrepPreferences = app.changeGrepPreferences = null;
    return totalAlteracoes;
  }

  /*--------------------------------------------------------------------------*/

  //FUNCOES DE EXECUCAO, FAZEM AS ALTERACOES NO ARQUIVO UTILIZANDO GREP

  //------------------------------FUNÇÕES TROCA DE DECIMAIS----------------------------------------

  function estaEmReferenciaCruzada(trecho) {
    var doc = app.activeDocument;
    var refs = doc.crossReferenceSources;

    try {
      for (var i = 0; i < refs.length; i++) {
        var fonte = refs[i].sourceText;

        if (!fonte || !trecho) continue;

        // Verifica se estão no mesmo story (história de texto)
        if (fonte.parentStory !== trecho.parentStory) continue;

        // Pega os índices absolutos no story
        var inicioFonte = fonte.insertionPoints[0].index;
        var fimFonte = fonte.insertionPoints[-1].index;

        var inicioTrecho = trecho.insertionPoints[0].index;
        var fimTrecho = trecho.insertionPoints[-1].index;

        if (inicioTrecho >= inicioFonte && fimTrecho <= fimFonte) {
          return true;
        }
      }
    } catch (e) {
      $.writeln("Erro na verificação de referência cruzada: " + e.message);
    }

    return false;
  }

  function substituirPontoPorVirgula() {
    var doc = app.activeDocument;
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

    app.findGrepPreferences.findWhat = "(?<=\\d)\\.(?=\\d)";
    var substituirPor = ",";
    var alterados = 0;
    var ignorados = 0;

    var resultados = [];

    if (aplicarEmTodasAsPaginas) {
      // Aplica em todo o documento
      resultados = doc.findGrep();
    } else {
      // Aplica apenas nas páginas especificadas que o user colocou manualmente
      for (var i = 0; i < paginasArray.length; i++) {
        var pagina = doc.pages[paginasArray[i] - 1];
        if (!pagina) continue;

        var itens = pagina.textFrames;
        for (var j = 0; j < itens.length; j++) {
          var tf = itens[j];
          app.findGrepPreferences = app.changeGrepPreferences =
            NothingEnum.nothing;
          app.findGrepPreferences.findWhat = "(?<=\\d)\\.(?=\\d)";
          var encontrados = tf.findGrep();

          for (var k = 0; k < encontrados.length; k++) {
            resultados.push(encontrados[k]);
          }
        }
      }
    }

    //  Agora aplica a substituição
    for (var i = 0; i < resultados.length; i++) {
      var match = resultados[i];
      var fontStyle = "";
      var isBold = false;
      var isItalic = false;

      try {
        fontStyle = match.appliedFont.fontStyleName.toLowerCase();
        isBold = fontStyle.indexOf("bold") !== -1;
        isItalic =
          fontStyle.indexOf("italic") !== -1 ||
          fontStyle.indexOf("oblique") !== -1;
      } catch (e) {}

      if (!isBold && !isItalic && !estaEmReferenciaCruzada(match)) {
        match.contents = substituirPor;
        alterados++;
      } else {
        ignorados++;
      }
    }

    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    return alterados;
  }

  function substituirVirgulaPorPonto() {
    var doc = app.activeDocument;
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

    app.findGrepPreferences.findWhat = "(?<=\\d)\\,(?=\\d)";
    var substituirPor = ".";
    var alterados = 0;
    var ignorados = 0;

    var resultados = [];

    if (aplicarEmTodasAsPaginas) {
      // Aplica em todo o documento
      resultados = doc.findGrep();
    } else {
      // Aplica apenas nas páginas especificadas que o user escreveu
      for (var i = 0; i < paginasArray.length; i++) {
        var pagina = doc.pages[paginasArray[i] - 1];
        if (!pagina) continue;

        var itens = pagina.textFrames;
        for (var j = 0; j < itens.length; j++) {
          var tf = itens[j];
          app.findGrepPreferences = app.changeGrepPreferences =
            NothingEnum.nothing;
          app.findGrepPreferences.findWhat = "(?<=\\d)\\,(?=\\d)";
          var encontrados = tf.findGrep();

          for (var k = 0; k < encontrados.length; k++) {
            resultados.push(encontrados[k]);
          }
        }
      }
    }

    //  Agora aplica a substituição
    for (var i = 0; i < resultados.length; i++) {
      var match = resultados[i];
      var fontStyle = "";
      var isBold = false;
      var isItalic = false;

      try {
        fontStyle = match.appliedFont.fontStyleName.toLowerCase();
        isBold = fontStyle.indexOf("bold") !== -1;
        isItalic =
          fontStyle.indexOf("italic") !== -1 ||
          fontStyle.indexOf("oblique") !== -1;
      } catch (e) {}

      if (!isBold && !isItalic && !estaEmReferenciaCruzada(match)) {
        match.contents = substituirPor;
        alterados++;
      } else {
        ignorados++;
      }
    }

    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    return alterados;
  }

  function trim(str) {
    return str.replace(/^\s+|\s+$/g, "");
  }

  function arrayContem(arr, val) {
    for (var i = 0; i < arr.length; i++) {
      if (arr[i] === val) return true;
    }
    return false;
  }

  function parsePaginas(str, totalPaginas) {
    var resultado = [];
    if (!str) return resultado;

    try {
      var partes = str.split(",");
      for (var i = 0; i < partes.length; i++) {
        var parte = partes[i] && partes[i].toString();
        if (!parte) continue;

        parte = parte.replace(/^\s+|\s+$/g, ""); // trim manual

        if (parte.indexOf("-") > -1) {
          var limites = parte.split("-");
          var inicio = parseInt(limites[0], 10);
          var fim = parseInt(limites[1], 10);

          // Se o intervalo for inválido ou qualquer das pontas estiver fora do total
          if (
            isNaN(inicio) ||
            isNaN(fim) ||
            fim < inicio ||
            inicio < 1 ||
            fim > totalPaginas
          ) {
            alert("Intervalo inválido ou fora do limite: " + parte);
            continue;
          }

          for (var n = inicio; n <= fim; n++) {
            if (!arrayContem(resultado, n)) {
              resultado.push(n);
            }
          }
        } else {
          var num = parseInt(parte, 10);
          if (
            !isNaN(num) &&
            num >= 1 &&
            num <= totalPaginas &&
            !arrayContem(resultado, num)
          ) {
            resultado.push(num);
          } else {
            alert("Número inválido ou fora do limite: " + parte);
          }
        }
      }

      resultado.sort(function (a, b) {
        return a - b;
      });
      //alert("Resultado final: " + resultado.join(", ")); //descomentar caso queira ver o numero de páginas que serão alteradas
      return resultado;
    } catch (erro) {
      alert("Erro em parsePaginas: " + erro.message);
      return [];
    }
  }

  function validarPaginasExistentes(arrayPaginas, totalPaginas) {
    var paginasValidas = [];
    for (var i = 0; i < arrayPaginas.length; i++) {
      var pg = arrayPaginas[i];
      if (pg >= 1 && pg <= totalPaginas) {
        paginasValidas.push(pg);
      } else {
        //alert("Página inválida ignorada: " + pg);
      }
    }
    return paginasValidas;
  }

  //-----------------------------------------------------------------------------------------------

  function removerEspacoAntesVirgula() {
    //remove qualquer espaço antes de uma virgula
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\s+,";
    app.changeGrepPreferences.changeTo = ",";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
  }

  function corrigirEspacoDepoisVirgula() {
    //corrige a "," depois de uma LETRA, nao funciona de for numeros
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = ",(?!\\s|\\d)(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])";
    app.changeGrepPreferences.changeTo = ", ";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
  }

  function removerEspacoAntesDoisPontos() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\s+:"; // espaços antes do :
    app.changeGrepPreferences.changeTo = ":";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
  }

  function corrigirEspacoDepoisDoisPontos() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    var totalAlteracoes = 0;

    if (ajustarEspacamentoDoisPontos_Antes) {
      // Corrige todos os casos para " : "
      app.findGrepPreferences.findWhat = "\\s*:\\s*";
      app.changeGrepPreferences.changeTo = " : ";
      var changed = doc.changeGrep();
      totalAlteracoes += changed ? changed.length : 0;
    } else {
      // Corrige apenas o espaço depois dos dois pontos
      app.findGrepPreferences.findWhat = ":(?! )(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])";
      app.changeGrepPreferences.changeTo = ": ";
      var changed = doc.changeGrep();
      totalAlteracoes += changed ? changed.length : 0;
    }

    app.findGrepPreferences = app.changeGrepPreferences = null;
    return totalAlteracoes;
  }

  function removerEspacoAntesPontoVirgula() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\s+;"; // espaços antes do ;
    app.changeGrepPreferences.changeTo = ";";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
  }

  function corrigirEspacoDepoisPontoVirgula() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = ";(?! )(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])"; // ; sem espaço depois
    app.changeGrepPreferences.changeTo = "; ";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
  }

  function ajustarBarraIgual() {
    app.findGrepPreferences = app.changeGrepPreferences = null;

    var totalAlteracoes = 0;

    // Se a opção de ignorar tabelas estiver marcada
    if (ignorarTabelasBarra) {
      var stories = doc.stories;
      for (var i = 0; i < stories.length; i++) {
        var story = stories[i];

        if (story.tables.length > 0) continue;

        app.findGrepPreferences.findWhat = "([^\\s])/";
        app.changeGrepPreferences.changeTo = "$1 /";
        var changed1 = story.changeGrep();

        app.findGrepPreferences.findWhat = "/(\\S)";
        app.changeGrepPreferences.changeTo = "/ $1";
        var changed2 = story.changeGrep();

        totalAlteracoes +=
          (changed1 ? changed1.length : 0) + (changed2 ? changed2.length : 0);
      }
    } else {
      // Aplica no documento inteiro
      app.findGrepPreferences.findWhat = "([^\\s])/";
      app.changeGrepPreferences.changeTo = "$1 /";
      var changed1 = doc.changeGrep();

      app.findGrepPreferences.findWhat = "/(\\S)";
      app.changeGrepPreferences.changeTo = "/ $1";
      var changed2 = doc.changeGrep();

      totalAlteracoes =
        (changed1 ? changed1.length : 0) + (changed2 ? changed2.length : 0);
    }

    app.findGrepPreferences = app.changeGrepPreferences = null;

    return totalAlteracoes;
  }

  function ajustarEspacoIgual(doc) {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    var doc = app.activeDocument;

    // Corrige espaços antes do '='
    app.findGrepPreferences.findWhat = "([^\\s])=";
    app.changeGrepPreferences.changeTo = "$1 =";
    var changed1 = doc.changeGrep();

    // Corrige espaços depois do '='
    app.findGrepPreferences.findWhat = "=(\\S)";
    app.changeGrepPreferences.changeTo = "= $1";
    var changed2 = doc.changeGrep();

    app.findGrepPreferences = app.changeGrepPreferences = null;

    return (changed1 ? changed1.length : 0) + (changed2 ? changed2.length : 0);
  }

  function removerEspacoAntesPontoFinal() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\s+\\."; // espaços antes do ponto
    app.changeGrepPreferences.changeTo = ".";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
  }

  function corrigirEspacoDepoisPontoFinal() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\.(?!\\s|\\d)(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])";
    app.changeGrepPreferences.changeTo = ". ";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
  }
  function ajustarEspacoAntesPercentual(doc) {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    var doc = app.activeDocument;

    // Espaço antes do % quando falta
    app.findGrepPreferences.findWhat = "([^\\s])%";
    app.changeGrepPreferences.changeTo = "$1 %";
    var changed1 = doc.changeGrep();

    // Espaço depois do % quando falta
    app.findGrepPreferences.findWhat = "%(\\S)";
    app.changeGrepPreferences.changeTo = "% $1";
    var changed2 = doc.changeGrep();

    app.findGrepPreferences = app.changeGrepPreferences = null;

    return (changed1 ? changed1.length : 0) + (changed2 ? changed2.length : 0);
  }

  /*
    CAPITALIZAÇÃO 
    */

  function getAllParagraphStyles(doc) {
    var estilos = [];

    function coletar(estilosLista) {
      for (var i = 0; i < estilosLista.length; i++) {
        var estilo = estilosLista[i];
        if (!estilo.name.match(/^\[.*\]$/)) {
          estilos.push(estilo.name);
        }
      }
    }

    coletar(doc.paragraphStyles);

    var grupos = doc.paragraphStyleGroups;
    for (var j = 0; j < grupos.length; j++) {
      coletar(grupos[j].paragraphStyles);
    }

    return estilos;
  }

  function abrirJanelaAdicionarSigla() {
    var caminho =
      "Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/SiglasPadrao.txt";

    function salvarSigla(sigla) {
      var arquivo = File(caminho);
      var siglas = [];

      if (arquivo.exists) {
        if (arquivo.open("r")) {
          var conteudo = arquivo.read();
          arquivo.close();
          siglas = conteudo.split(/[\r\n]+/);
        } else {
          alert("Erro ao abrir o arquivo para leitura.");
          return;
        }
      } else {
        alert("Arquivo NÃO existe. Será criado.");
      }

      // Remove espaços e repete uppercase
      sigla = sigla.replace(/\s+/g, "").toUpperCase();

      // Verifica duplicata sem usar .indexOf
      var jaExiste = false;
      for (var i = 0; i < siglas.length; i++) {
        if (siglas[i].toUpperCase() === sigla) {
          jaExiste = true;
          break;
        }
      }

      if (jaExiste) {
        alert("Essa sigla já existe nesse banco de dados! ");
        return;
      }

      siglas.push(sigla);

      if (arquivo.open("w")) {
        arquivo.write(siglas.join("\n"));
        arquivo.close();
        alert("Sigla salva com sucesso.");
      } else {
        alert("Erro ao abrir o arquivo para escrita.");
      }
    }

    // Cria a janela
    var win = new Window("dialog", "Adicionar Sigla");
    win.orientation = "column";
    win.alignChildren = "fill";

    win.add("statictext", undefined, "Digite a sigla:");
    var input = win.add("edittext", undefined, "");
    input.characters = 20;
    input.active = true;

    var botoes = win.add("group");
    var salvar = botoes.add("button", undefined, "Salvar", { name: "ok" });
    var cancelar = botoes.add("button", undefined, "Cancelar", {
      name: "cancel",
    });

    salvar.onClick = function () {
      var sigla = input.text;
      salvarSigla(sigla);
      win.close();
    };

    cancelar.onClick = function () {
      win.close();
    };

    win.center();
    win.show();
  }

  function carregarSiglasDoArquivo() {
    try {
      var arquivo = File(
        "Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/SiglasPadrao.txt"
      );

      if (!arquivo.exists) {
        alert("Arquivo de siglas não encontrado.");
        return [];
      }

      arquivo.open("r");
      var conteudo = arquivo.read();
      arquivo.close();

      var linhasBrutas = conteudo.split(/[\r\n]+/);
      var siglas = [];

      for (var i = 0; i < linhasBrutas.length; i++) {
        var linha = linhasBrutas[i];
        if (linha && linha.replace(/\s+/g, "") !== "") {
          siglas.push(linha.replace(/^\s+|\s+$/g, "")); // trim manual
        }
      }

      return siglas;
    } catch (e) {
      alert("Erro ao carregar siglas:\n" + e.message);
      return [];
    }
  }

  function detectarSiglaPresente(palavra, siglas) {
    for (var i = 0; i < siglas.length; i++) {
      var sigla = siglas[i];
      var siglaLower = sigla.toLowerCase();
      var palavraLower = palavra.toLowerCase();

      if (palavraLower === siglaLower) {
        return sigla; // palavra inteira é a sigla
      }

      if (
        palavraLower.indexOf(siglaLower) === 0 && // começa com a sigla
        palavra.length > sigla.length // há mais coisa depois
      ) {
        var proximoChar = palavra.charAt(sigla.length);

        // Só aplica se o próximo caractere não for uma letra (A-Z ou a-z)
        if (!/[a-zA-Z]/.test(proximoChar)) {
          return sigla;
        }
      }
    }
    return null;
  }

  function aplicarCapitalizacaoComGREP(
    estilosSelecionados,
    idiomaSelecionado,
    capitalizacaoSelecionada
  ) {
    try {
      var doc = app.activeDocument;
      var paragrafos = doc.stories
        .everyItem()
        .paragraphs.everyItem()
        .getElements();

      if (!estilosSelecionados || !(estilosSelecionados instanceof Array)) {
        alert("estilosSelecionados não é um array válido.");
        return 0;
      }

      var excecoes = {
        Português: [
          "de",
          "na",
          "pra",
          "da",
          "do",
          "das",
          "dos",
          "e",
          "em",
          "com",
          "por",
          "para",
          "a",
          "o",
          "as",
          "os",
          "nos",
        ],
        Inglês: [
          "a",
          "an",
          "and",
          "but",
          "or",
          "for",
          "nor",
          "as",
          "at",
          "by",
          "in",
          "of",
          "on",
          "per",
          "to",
          "the",
          "with",
        ],
        Polonês: [
          "i",
          "oraz",
          "a",
          "ale",
          "lecz",
          "lub",
          "czy",
          "nie",
          "w",
          "na",
          "do",
          "z",
          "od",
          "pod",
          "nad",
          "przed",
          "po",
          "bez",
          "dla",
          "o",
          "u",
        ],
        Francês: [
          "à",
          "au",
          "aux",
          "avec",
          "chez",
          "dans",
          "de",
          "des",
          "du",
          "en",
          "et",
          "la",
          "le",
          "les",
          "ou",
          "par",
          "pour",
          "sur",
          "un",
          "une",
        ],
      };

      var palavrasMin = excecoes[idiomaSelecionado] || [];
      var siglasMaiusculas = carregarSiglasDoArquivo();

      function estaNaLista(palavra, lista) {
        for (var i = 0; i < lista.length; i++) {
          if (palavra.toLowerCase() === lista[i].toLowerCase()) {
            return true;
          }
        }
        return false;
      }

      // Definir tipo de capitalização
      var modoChangecase = null;
      if (capitalizacaoSelecionada === "Maiúsculas") {
        modoChangecase = ChangecaseMode.UPPERCASE;
      } else if (capitalizacaoSelecionada === "Minúsculas") {
        modoChangecase = ChangecaseMode.LOWERCASE;
      } else if (capitalizacaoSelecionada === "Title Case") {
        modoChangecase = "TITULO_MANUAL";
      } else {
        alert("Tipo de Alteração de Caixa não reconhecido.");
        return 0;
      }

      var contador = 0;

      for (var i = 0; i < paragrafos.length; i++) {
        var p = paragrafos[i];
        if (!p.appliedParagraphStyle || !p.appliedParagraphStyle.name) continue;

        var nomeEstilo = String(p.appliedParagraphStyle.name).toLowerCase();
        var estiloEncontrado = false;

        for (var j = 0; j < estilosSelecionados.length; j++) {
          if (nomeEstilo === String(estilosSelecionados[j]).toLowerCase()) {
            estiloEncontrado = true;
            break;
          }
        }

        if (!estiloEncontrado) continue;

        if (modoChangecase === "TITULO_MANUAL") {
          var palavras = p.words;

          // Primeiro, coloca tudo em maiúsculas
          for (var k = 0; k < palavras.length; k++) {
            palavras[k].changecase(ChangecaseMode.UPPERCASE);
          }

          // Depois, aplica o Title Case manual
          for (var k = 0; k < palavras.length; k++) {
            var palavraOriginal = palavras[k].contents;
            var palavraTexto = palavraOriginal.toLowerCase();

            var siglaEncontrada = detectarSiglaPresente(
              palavraOriginal,
              siglasMaiusculas
            );

            if (siglaEncontrada) {
              var novaPalavra = palavraOriginal.replace(
                new RegExp("^" + siglaEncontrada, "i"),
                siglaEncontrada.toUpperCase()
              );
              palavras[k].contents = novaPalavra;
            } else if (k === 0 || !estaNaLista(palavraTexto, palavrasMin)) {
              // Verifica se é tipo d'Expansion, com a letra após o apóstrofo maiúscula
              if (palavraOriginal.match(/^[dlmnstcj][’'][a-zA-ZÀ-ÖØ-öø-ÿ]/i)) {
                var partes = palavraOriginal.split(/['’]/);
                if (partes.length === 2) {
                  var ap = palavraOriginal.charAt(1); // mantém o apóstrofo original (’ ou ')
                  var prefixo = partes[0].toLowerCase();
                  var principal = partes[1];

                  var siglaEncontrada = detectarSiglaPresente(
                    principal,
                    siglasMaiusculas
                  );
                  if (siglaEncontrada) {
                    var novaPalavra =
                      prefixo + ap + siglaEncontrada.toUpperCase();
                  } else {
                    var novaPalavra =
                      prefixo +
                      ap +
                      principal.charAt(0).toUpperCase() +
                      principal.slice(1).toLowerCase();
                  }

                  palavras[k].contents = novaPalavra;
                  continue; // já tratou essa palavra
                }
              } else {
                palavras[k].changecase(ChangecaseMode.TITLECASE);
              }
            } else {
              palavras[k].changecase(ChangecaseMode.LOWERCASE);
            }
          }
        } else {
          p.changecase(modoChangecase);
        }

        contador++;
      }

      return contador;
    } catch (e) {
      alert("Erro ao aplicar Alteração de Caixa:\n" + e.message);
      return 0;
    }
  }

  TelaInicial();
})();
