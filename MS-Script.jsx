/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * @file           : MS-Script
  * @brief          : Script for InDesign Automations
  * @author 	      : @mschwertz / 2025
  ******************************************************************************
  * @attention

  ******************************************************************************
  */
/* USER CODE END Header */

(function () {
  ("highlight Force Decorate");
  if (app.documents.length === 0) {
    alert("Abra um documento antes de rodar o script.");
    return;
  }

  var doc = app.activeDocument;

  // PL -  „|”
  // G  -  „|“
  // F  - « | »
  // R  -  «|»
  // PD -  “|”

  var AspasAbre = "“«„";
  var AspasFecha = "”“»";
  var aspasAbertura = "";
  var aspasFechamento = "";
  var aplicarAspas = false;

  var ignorarTabelasBarra = false;
  var ajustarEspacamentoDoisPontos_Antes = false;
  var ignorarExclusaoParenteses = false;
  var ignorarExclusaoColchetes = false;
  var ajustarBarraEspaco = false;
  var ajustarVirgulaEspaco = false;
  var ajustarIgualEspaco = false;
  var ajustarEspacamentoDoisPontos = false;
  var ajustarPontoFinalemTabela = false;
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
  var contadorPontoFinalemTabela = 0;
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

    var linha = win.add("panel");
    linha.alignment = "fill";

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

    var linha1 = win.add("panel");
    linha1.alignment = "fill";

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

    var linha = grupoDecimal.add("panel");
    linha.alignment = "fill";

    var grupoExclusao = grupoDecimal.add("group");
    grupoExclusao.orientation = "row";
    grupoExclusao.alignment = "center";

    var ExclusaoParenteses = grupoExclusao.add(
      "checkbox",
      undefined,
      "Manter → (  )"
    );
    ExclusaoParenteses.helpTip =
      "Se marcado, o conteúdo dentro dos parênteses não será afetado pela troca decimal.";

    var ExclusaoColchetes = grupoExclusao.add(
      "checkbox",
      undefined,
      "Manter → [  ]"
    );
    ExclusaoColchetes.helpTip =
      "Se marcado, o conteúdo dentro dos colchetes não será afetado pela troca decimal.";

    var linha = grupoDecimal.add("panel");
    linha.alignment = "fill";

    var chkTodasPaginas = grupoDecimal.add(
      "checkbox",
      undefined,
      "Aplicar a todas as páginas?"
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
      ignorarExclusaoParenteses = ExclusaoParenteses.value;
      ignorarExclusaoColchetes = ExclusaoColchetes.value;

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

    var linha = win.add("panel");
    linha.alignment = "fill";

    var chkPontoFinalTabela = win.add(
      "checkbox",
      undefined,
      "Verificar os Pontos Finais dentro Tabelas?"
    );

    chkPontoFinalTabela.onClick = function () {
      if (chkPontoFinalTabela.value) {
        // Desmarca todos os outros checkboxes
        for (var i = 0; i < checkboxes.length; i++) {
          checkboxes[i].value = false;
          checkboxes[i].enabled = false;
        }
        chk2PEspacoAntes.value = false;
        chk2PEspacoAntes.enabled = false;
        chkTabelas.value = false;
        chkTabelas.enabled = false;
        chkSelecionarTodos.value = false;
        chkSelecionarTodos.enabled = false;
      } else {
        // Reabilita todos quando desmarcado
        for (var i = 0; i < checkboxes.length; i++) {
          checkboxes[i].enabled = true;
        }
        chk2PEspacoAntes.enabled = true;
        chkTabelas.enabled = true;
        chkSelecionarTodos.enabled = true;
      }
    };

    var chkSelecionarTodos = win.add(
      "checkbox",
      undefined,
      "Selecionar Todas as Padrozinações abaixo"
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
      // adicionar as checkboxes que devem ser selecionadas com o "Selecionar Todos"
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
      ajustarPontoFinalemTabela = chkPontoFinalTabela.value;
      ajustarVirgulaEspaco = chkVirgulaEspaco.value;
      ajustarIgualEspaco = chkIgualEspaco.value;
      ajustarBarraEspaco = chkBarraEspaco.value;

      ignorarTabelasBarra = chkTabelas.value;
      ajustarEspacamentoDoisPontos_Antes = chk2PEspacoAntes.value;

      ajustarEspacamentoDoisPontos = chkDoisPontosEspaco.value;
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
    win.alignChildren = "fill";

    var grupoTexto = win.add("group");
    grupoTexto.alignment = "fill";
    grupoTexto.alignChildren = "center"; // filhos centralizados
    grupoTexto.add(
      "statictext",
      undefined,
      "     Escolha o idioma e os estilos de parágrafo para aplicar a Alteração de Caixa::::::"
    );

    var linha = win.add("panel");
    linha.alignment = "fill";

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
    grupoIdioma.alignment = "center";
    var dropdownIdioma = grupoIdioma.add("dropdownlist", undefined, idiomas);
    dropdownIdioma.selection = 0;

    // Dropdown de estilos
    var grupoCapitalizacao = win.add("group");
    grupoCapitalizacao.orientation = "row";
    grupoCapitalizacao.alignment = "center";
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

    if (ajustarPontoFinalemTabela)
      resumo += "\n -  Ajustes em ' . ' dentro de tabelas\n";
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
      !ajustarPontoFinalemTabela &&
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
      removerEspacoAntesAspaRetas(doc);
      removerEspacoDepoisAspaRetas(doc);
      substituirAspasRetas(doc);

      removerEspacoDepoisAspaAbertura(doc);
      removerEspacoAntesAspaFechamento(doc);

      contadorAspasAberto = corrigirAspasAbertura(aspasAbertura, doc);
      contadorAspasFechado = corrigirAspasFechamento(aspasFechamento, doc);

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

    if (ajustarBarraEspaco) {
      contadorBarraEspaco = ajustarBarrEspaco(
        app.activeDocument,
        ignorarTabelasBarra
      );
    }

    if (ajustarIgualEspaco) contadorIgualEspaco = ajustarEspacoIgual();

    if (ajustarPontoFinalemTabela) {
      ajustarEspacoPontoFinalTabela(function (count) {
        contadorPontoFinalemTabela = count;

        // monta a mensagem completa aqui, incluindo todos os outros contadores
        var mensagem = "\nCorreções concluídas!\n\n";
        if (aplicarAspas)
          mensagem += "✅ " + contadorAspas + " - Aspas alteradas\n";
        if (contadorPontoFinalemTabela)
          mensagem +=
            "✅ " +
            contadorPontoFinalemTabela +
            " - Ajustes em Pontos finais dentro de tabelas\n";

        // chama a tela de resultado final
        telaResultadoFinal(mensagem);
      });

      return; // evita continuar antes do usuário finalizar
    }

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
      !ajustarEspacoPontoFinalTabela &&
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

    var contadoresAlteracoes = {
      "Aspas alteradas": contadorAspas,
      "Pontos trocados por vírgula": contadorPontoVirgula,
      "Vírgulas trocadas por ponto": contadorVirgulaPonto,
      "Ajustes em ' . '": contadorPontoFinalEspaco,
      "Ajustes em ' , '": contadorVirgulaEspaco,
      "Ajustes em ' = '": contadorIgualEspaco,
      "Ajustes em ' : '": contadorEspacamentoDoisPontos,
      "Ajustes em ' ; '": contadorEspacoDepoisPontoVirgula,
      "Ajustes em ' % '": contadorEspacoAntesPercent,
      "Ajustes em ' / '": contadorBarraEspaco,
      "Alterações de Caixa": contadorCapitalizacao,
    };

    telaResultadoFinal(mensagem, contadoresAlteracoes);
  }

  function telaResultadoFinal(resumoExecucao, contadoresAlteracoes) {
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

    var temposalvo = botoes.add("button", undefined, "Tempo Salvo");
    temposalvo.helpTip = "Exibir o tempo salvo pela execução do script.";

    var ok = botoes.add("button", undefined, "OK");
    ok.helpTip = "Manter alterações e fechar o script.";

    var cancelar = botoes.add("button", undefined, "Cancelar");

    ok.onClick = function () {
      win.close();
    };

    cancelar.onClick = function () {
      alert(
        "Se você quiser desfazer todas as alterações realizadas, use o comando CTRL + Z no InDesign."
      );
    };

    temposalvo.onClick = function () {
      var ops = {
        aplicarAspas: aplicarAspas,
        trocarPontoPorVirgula: trocarPontoPorVirgula,
        trocarVirgulaPorPonto: trocarVirgulaPorPonto,
        ajustarEspacamentoDoisPontos: ajustarEspacamentoDoisPontos,
        ajustarEspacamentoPontoVirgula: ajustarEspacamentoPontoVirgula,
        ajustarBarraEspaco: ajustarBarraEspaco,
        ajustarIgualEspaco: ajustarIgualEspaco,
        ajustarEspacoAntesPercent: ajustarEspacoAntesPercent,
      };

      calcularTempoSalvo(contadoresAlteracoes, ops); // abre a nova
      win.close(); // só fecha a original depois que a nova for exibida
    };

    win.center();
    win.show();
  }

  function calcularTempoSalvo(contadoresAlteracoes, ops) {
    // --- Configuração de tarefas e GREPs relevantes por operação ---
    var tarefas = [
      {
        op: "trocarPontoPorVirgula",
        label: "Pontos trocados por vírgula",
        greps: ["\\.", "\\,"],
      },
      {
        op: "trocarVirgulaPorPonto",
        label: "Vírgulas trocadas por ponto",
        greps: ["\\,", "\\."],
      },
      {
        op: "ajustarEspacamentoDoisPontos",
        label: "Ajustes em ' : '",
        greps: ["\\:"],
      },
      {
        op: "ajustarEspacamentoPontoVirgula",
        label: "Ajustes em ' ; '",
        greps: ["\\;"],
      },
      { op: "ajustarBarraEspaco", label: "Ajustes em ' / '", greps: ["\\/"] },
      { op: "ajustarIgualEspaco", label: "Ajustes em ' = '", greps: ["\\="] },
      {
        op: "ajustarEspacoAntesPercent",
        label: "Ajustes em ' % '",
        greps: ["\\%"],
      },
      // Aspas é opcional; aqui considero aspas retas (")
      { op: "aplicarAspas", label: "Aspas alteradas", greps: ['"'] },
    ];

    // Nome "bonito" para cada grep (pra listar na seção de totais)
    var grep2Nome = {
      "\\.": "Ponto(s) final",
      "\\,": "Vírgula(s)",
      "\\:": "Dois-pontos",
      "\\;": "Ponto e vírgula",
      "\\/": "Barra(s)",
      "\\=": "Igual",
      "\\%": "Porcentagem",
      '"': 'Aspas (")',
    };

    // Considera só operações ligadas
    var tarefasAtivas = [];
    for (var t = 0; t < tarefas.length; t++) {
      var tf = tarefas[t];
      if (ops && ops[tf.op]) tarefasAtivas.push(tf);
    }

    // Se nada estiver ativo, só mostra mensagem simples
    if (!tarefasAtivas.length) {
      var w0 = new Window("dialog", "Resultado da Correção ⏱");
      w0.orientation = "column";
      w0.alignChildren = "center";
      w0.add("statictext", undefined, "Nenhuma operação foi selecionada.");
      var fechar0 = w0.add("button", undefined, "Fechar");
      fechar0.onClick = function () {
        w0.close();
      };
      w0.show();
      return;
    }

    // ---------- Seção: totais globais apenas do que é relevante ----------
    var resultado =
      "---- Quantidade total (apenas do que foi selecionado) ----\n\n";

    // Deduplica greps globais (ex.: se as duas tarefas de decimal estiverem ativas)
    var grepsGlobais = {};
    for (var i = 0; i < tarefasAtivas.length; i++) {
      var g = tarefasAtivas[i].greps;
      for (var j = 0; j < g.length; j++) grepsGlobais[g[j]] = true;
    }

    var totalOcorrenciasGlobais = 0;
    for (var grepStr in grepsGlobais) {
      app.findGrepPreferences = NothingEnum.nothing;
      app.findGrepPreferences.findWhat = grepStr;
      var encontrados = app.activeDocument.findGrep();
      var qtd = encontrados.length;
      if (qtd > 0) {
        var nome = grep2Nome[grepStr] || grepStr;
        resultado += nome + ": " + qtd + "\n";
        totalOcorrenciasGlobais += qtd;
      }
    }

    resultado +=
      "\n-------------- Alterações realizadas pelo script --------------\n\n";

    // Lista apenas contadores > 0
    if (contadoresAlteracoes) {
      for (var key in contadoresAlteracoes) {
        if (contadoresAlteracoes[key] > 0) {
          resultado += "✅ " + contadoresAlteracoes[key] + " - " + key + "\n";
        }
      }
    } else {
      resultado += "Nenhuma alteração feita.\n";
    }

    // ---------- Tempo estimado: só tarefas ativas ----------
    var tempoNaoAlterado = 3; // s p/ item inspecionado sem mudança
    var tempoAlterado = 8; // s p/ item que precisou mudar
    var tempoTotalSegundos = 0;

    for (var x = 0; x < tarefasAtivas.length; x++) {
      var tarefa = tarefasAtivas[x];

      // total relevante da tarefa (somatório dos seus GREPs)
      var totalTarefa = 0;
      for (var y = 0; y < tarefa.greps.length; y++) {
        app.findGrepPreferences = NothingEnum.nothing;
        app.findGrepPreferences.findWhat = tarefa.greps[y];
        totalTarefa += app.activeDocument.findGrep().length;
      }

      // quantas mudanças essa tarefa fez de fato
      var alteradas = 0;
      if (
        contadoresAlteracoes &&
        contadoresAlteracoes[tarefa.label] !== undefined
      ) {
        alteradas = Number(contadoresAlteracoes[tarefa.label]) || 0;
      }

      var naoAlteradas = Math.max(totalTarefa - alteradas, 0);
      tempoTotalSegundos +=
        naoAlteradas * tempoNaoAlterado + alteradas * tempoAlterado;
    }

    // Converte tempo
    var tempoTotalMinutos = tempoTotalSegundos / 60;
    var tempoTotalHoras = tempoTotalMinutos / 60;

    resultado +=
      "\n-------------------- Tempo total estimado --------------------\n\n";
    resultado += "- " + tempoTotalSegundos.toFixed(0) + " segundos\n";
    resultado += "- " + tempoTotalMinutos.toFixed(2) + " minutos\n";
    resultado += "- " + tempoTotalHoras.toFixed(2) + " horas\n";

    // ---------- Janela ----------
    var winTempoSalvo = new Window("dialog", "Resultado da Correção ⏱");
    winTempoSalvo.orientation = "column";
    winTempoSalvo.alignChildren = "center";

    winTempoSalvo.add(
      "statictext",
      undefined,
      "Execução concluída com sucesso!"
    );

    var caixaResultado = winTempoSalvo.add("edittext", undefined, resultado, {
      multiline: true,
      readonly: true,
    });
    caixaResultado.preferredSize.width = 320;
    caixaResultado.preferredSize.height = 300;

    var fechar = winTempoSalvo.add("button", undefined, "Fechar");
    fechar.onClick = function () {
      winTempoSalvo.close();
    };

    // limpa prefs no final por segurança
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

    winTempoSalvo.show();
  }

  /* FUNCOES DE CORRECAO DE ASPAS

          Funções para remover as aspas retas do arquivo.
  Basicamente, elas removem os espaços antes e depois das aspas retas,
  e depois substituem as aspas retas por aspas curvas, depois o codigo faz a substituição
  das aspas curvas por aspas de acordo com o idioma selecionado.

  Isso tudo porque o InDesign detecta qualquer aspa como uma aspa reta, se eu colocar pra ele buscar por aspas retas,
  ele vai substituir todas as aspas, polonesas,fracesas, alemãs, padrões, basicamente todas as aspas do arquivo.
  
*/
  function removerEspacoDepoisAspaRetas() {
    app.findGrepPreferences = app.changeGrepPreferences = null;

    app.findGrepPreferences.findWhat =
      '(?<![A-Za-zÀ-ÖØ-öø-ÿ0-9])(")[\x20\xA0\u202F]+(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])';
    app.changeGrepPreferences.changeTo = "$1";

    doc.changeGrep();

    app.findGrepPreferences = app.changeGrepPreferences = null;
  }

  function removerEspacoAntesAspaRetas() {
    do {
      app.findGrepPreferences = app.changeGrepPreferences = null;

      app.findGrepPreferences.findWhat =
        '([A-Za-zÀ-ÖØ-öø-ÿ0-9])\\s+(")(?=(\\s+["A-Za-zÀ-ÖØ-öø-ÿ])|\\s*$|\\s*[.,;:!?)]|$)';
      app.changeGrepPreferences.changeTo = "$1$2";

      var result = doc.changeGrep();
      var changes = result ? result.length : 0;

      app.findGrepPreferences = app.changeGrepPreferences = null;
    } while (changes > 0);
  }

  function substituirAspasRetas() {
    var doc = app.activeDocument;
    app.findGrepPreferences = app.changeGrepPreferences = null;

    // Busca somente aspas retas: "conteúdo"
    app.findGrepPreferences.findWhat = "\u0022([^\u0022]+?)\u0022";

    var encontrados = doc.findGrep();

    for (var i = 0; i < encontrados.length; i++) {
      var conteudo = encontrados[i].contents;
      var novoConteudo = "“" + conteudo.slice(1, -1) + "”";
      encontrados[i].contents = novoConteudo;
    }

    app.findGrepPreferences = app.changeGrepPreferences = null;
  }

  /*
  Foi necessario adicionarmos funcoes para retirar os espacos antes e depois das aspas,
  pois quando não havia essas funcoes acontecia o erro de duplicacao das aspas de abertura.
*/
  function removerEspacoDepoisAspaAbertura() {
    // alert("Atenção: Removendo espaços após aspas de abertura.");
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
    // alert("Atenção: Removendo espaços antes de aspas de fechamento.");
    var totalChanges = 0;
    var changes;

    do {
      app.findGrepPreferences = app.changeGrepPreferences = null;

      app.findGrepPreferences.findWhat =
        "([A-Za-zÀ-ÖØ-öø-ÿ0-9])\\s+([" +
        AspasFecha +
        "])(?=(\\s+[" +
        AspasAbre +
        "A-Za-zÀ-ÖØ-öø-ÿ])|\\s*$|\\s*[.,;:!?)]|$)";

      app.changeGrepPreferences.changeTo = "$1$2";

      var result = doc.changeGrep();
      changes = result ? result.length : 0;
      totalChanges += changes;

      app.findGrepPreferences = app.changeGrepPreferences = null;
    } while (changes > 0);

    return totalChanges;
  }

  function corrigirAspasAbertura(aspasAbertura) {
    // alert("Atenção: Corrigindo aspas de abertura.");
    app.findGrepPreferences = app.changeGrepPreferences = null;

    var listaAspas = AspasAbre.replace(aspasAbertura, "");

    var changes;
    do {
      app.findGrepPreferences.findWhat = "[" + listaAspas + "]";
      app.changeGrepPreferences.changeTo = aspasAbertura;

      var changed = doc.changeGrep();
      var totalAlteracoes = changed ? changed.length : 0;
    } while (changes > 0);

    app.findGrepPreferences = app.changeGrepPreferences = null;
    return totalAlteracoes;
  }

  function corrigirAspasFechamento(aspasFechamento) {
    // alert("Atenção: Corrigindo aspas de fechamento.");
    app.findGrepPreferences = app.changeGrepPreferences = null;

    var listaAspas = AspasFecha.replace(aspasFechamento, "");

    var aspasFechamentoEscapada = aspasFechamento.replace(
      /([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])/g,
      "\\$1"
    );

    app.findGrepPreferences.findWhat =
      "(?<=[^\\s\\(\\[\\{—])[" + listaAspas + "]";

    app.changeGrepPreferences.changeTo = aspasFechamentoEscapada;

    var changed = doc.changeGrep();
    var totalAlteracoes = changed ? changed.length : 0;

    app.findGrepPreferences = app.changeGrepPreferences = null;
    return totalAlteracoes;
  }

  /*--------------------------------------------------------------------------*/

  //------------------------------FUNÇÕES TROCA DE DECIMAIS----------------------------------------

  function estaDentroDeTabelaDecimais(objeto) {
    try {
      return objeto.parent.constructor.name === "Cell";
    } catch (e) {
      return false;
    }
  }

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

  function estaDentroDeExcecao(texto, pos, abre, fecha) {
    var dentro = false;
    for (var i = 0; i < pos; i++) {
      if (texto[i] === abre) {
        dentro = true;
      } else if (texto[i] === fecha) {
        dentro = false;
      }
    }
    return dentro;
  }

  function deveIgnorarPorExcecao(ponto, storyText, idx) {
    // Se estiver dentro de célula, usamos o texto da célula ao invés da story toda
    var textoBase = storyText;
    if (estaDentroDeTabelaDecimais(ponto)) {
      try {
        textoBase = ponto.parent.contents; // texto da célula
        idx = ponto.index; // índice relativo ao conteúdo da célula
      } catch (e) {}
    }

    var dentroDeParenteses =
      ignorarExclusaoParenteses &&
      estaDentroDeExcecao(textoBase, idx, "(", ")");
    var dentroDeColchetes =
      ignorarExclusaoColchetes && estaDentroDeExcecao(textoBase, idx, "[", "]");

    if (dentroDeParenteses || dentroDeColchetes) {
      return true;
    }

    return false;
  }

  /**
   * Versão “pura” da verificação de exceções, recebendo o texto-base e o índice relativo a esse texto.
   * (Assim você pode reaproveitar em qualquer escopo: story, frame ou célula.)
   */
  function deveIgnorarPorExcecaoComBase(textoBase, idx) {
    var dentroDeParenteses =
      ignorarExclusaoParenteses &&
      estaDentroDeExcecao(textoBase, idx, "(", ")");
    var dentroDeColchetes =
      ignorarExclusaoColchetes && estaDentroDeExcecao(textoBase, idx, "[", "]");
    return dentroDeParenteses || dentroDeColchetes;
  }

  function substituirPontoPorVirgula() {
    var doc = app.activeDocument;
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

    // GREP: ponto decimal entre dígitos
    app.findGrepPreferences.findWhat = "(?<=\\d)\\.(?=\\d)";
    var substituirPor = ",";
    var alterados = 0;
    var ignorados = 0;

    // Monta a lista de alvos de TEXTO (não stories!)
    // - Todas as páginas: usa stories -> texts[0]
    // - Páginas selecionadas: usa textFrames da página -> texts[0]
    var alvosTexto = [];

    if (aplicarEmTodasAsPaginas) {
      var todasStories = doc.stories.everyItem().getElements();
      for (var s = 0; s < todasStories.length; s++) {
        // Text do story inteiro
        alvosTexto.push(todasStories[s].texts[0]);
      }
    } else {
      for (var i = 0; i < paginasArray.length; i++) {
        // Índice físico (1-based -> 0-based). Se preferir, troque por itemByName(String(paginasArray[i])).
        var pagina = doc.pages[paginasArray[i] - 1];
        if (!pagina) continue;

        var textFrames = pagina.textFrames;
        for (var j = 0; j < textFrames.length; j++) {
          // Text do quadro daquela página
          alvosTexto.push(textFrames[j].texts[0]);
        }
      }
    }

    // Percorre SOMENTE os textos-alvo (restritos às páginas escolhidas)
    for (var t = 0; t < alvosTexto.length; t++) {
      var textoAlvo = alvosTexto[t];
      var matches = textoAlvo.findGrep();

      for (var k = matches.length - 1; k >= 0; k--) {
        var ponto = matches[k];

        // Índice relativo ao escopo atual (textoAlvo)
        var idx = ponto.insertionPoints[0].index;
        var baseText = textoAlvo.contents;

        // --- Ajuste para exceções dentro de TABELA ---
        // Se o match estiver numa célula, usamos o conteúdo da célula e o índice local da célula
        try {
          if (
            ponto.parent &&
            ponto.parent.constructor &&
            ponto.parent.constructor.name === "Cell"
          ) {
            baseText = ponto.parent.contents;
            idx = ponto.index; // índice relativo ao conteúdo da célula
          }
        } catch (e) {}

        // Verificação centralizada de exceções () e []
        if (deveIgnorarPorExcecaoComBase(baseText, idx)) {
          ignorados++;
          continue;
        }

        // Verificação de estilo (negrito/itálico) + referência cruzada
        var fontStyle = "";
        var isBold = false;
        var isItalic = false;
        try {
          fontStyle = (
            (ponto.appliedFont && ponto.appliedFont.fontStyleName) ||
            ""
          ).toLowerCase();
          isBold = fontStyle.indexOf("bold") !== -1;
          isItalic =
            fontStyle.indexOf("italic") !== -1 ||
            fontStyle.indexOf("oblique") !== -1;
        } catch (e) {}

        if (!isBold && !isItalic && !estaEmReferenciaCruzada(ponto)) {
          ponto.contents = substituirPor;
          alterados++;
        } else {
          ignorados++;
        }
      }
    }

    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    // Se quiser, mostre um resumo:
    // alert("Alterados: " + alterados + "\nIgnorados: " + ignorados);
    return alterados;
  }

  function substituirVirgulaPorPonto() {
    var doc = app.activeDocument;
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

    // GREP: ponto decimal entre dígitos
    app.findGrepPreferences.findWhat = "(?<=\\d)\\,(?=\\d)";
    var substituirPor = ".";
    var alterados = 0;
    var ignorados = 0;

    // Monta a lista de alvos de TEXTO (não stories!)
    // - Todas as páginas: usa stories -> texts[0]
    // - Páginas selecionadas: usa textFrames da página -> texts[0]
    var alvosTexto = [];

    if (aplicarEmTodasAsPaginas) {
      var todasStories = doc.stories.everyItem().getElements();
      for (var s = 0; s < todasStories.length; s++) {
        // Text do story inteiro
        alvosTexto.push(todasStories[s].texts[0]);
      }
    } else {
      for (var i = 0; i < paginasArray.length; i++) {
        // Índice físico (1-based -> 0-based). Se preferir, troque por itemByName(String(paginasArray[i])).
        var pagina = doc.pages[paginasArray[i] - 1];
        if (!pagina) continue;

        var textFrames = pagina.textFrames;
        for (var j = 0; j < textFrames.length; j++) {
          // Text do quadro daquela página
          alvosTexto.push(textFrames[j].texts[0]);
        }
      }
    }

    // Percorre SOMENTE os textos-alvo (restritos às páginas escolhidas)
    for (var t = 0; t < alvosTexto.length; t++) {
      var textoAlvo = alvosTexto[t];
      var matches = textoAlvo.findGrep();

      for (var k = matches.length - 1; k >= 0; k--) {
        var virgula = matches[k];

        // Índice relativo ao escopo atual (textoAlvo)
        var idx = virgula.insertionPoints[0].index;
        var baseText = textoAlvo.contents;

        // --- Ajuste para exceções dentro de TABELA ---
        // Se o match estiver numa célula, usamos o conteúdo da célula e o índice local da célula
        try {
          if (
            virgula.parent &&
            virgula.parent.constructor &&
            virgula.parent.constructor.name === "Cell"
          ) {
            baseText = virgula.parent.contents;
            idx = virgula.index; // índice relativo ao conteúdo da célula
          }
        } catch (e) {}

        // Verificação centralizada de exceções () e []
        if (deveIgnorarPorExcecaoComBase(baseText, idx)) {
          ignorados++;
          continue;
        }

        // Verificação de estilo (negrito/itálico) + referência cruzada
        var fontStyle = "";
        var isBold = false;
        var isItalic = false;
        try {
          fontStyle = (
            (virgula.appliedFont && virgula.appliedFont.fontStyleName) ||
            ""
          ).toLowerCase();
          isBold = fontStyle.indexOf("bold") !== -1;
          isItalic =
            fontStyle.indexOf("italic") !== -1 ||
            fontStyle.indexOf("oblique") !== -1;
        } catch (e) {}

        if (!isBold && !isItalic && !estaEmReferenciaCruzada(virgula)) {
          virgula.contents = substituirPor;
          alterados++;
        } else {
          ignorados++;
        }
      }
    }

    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    // Se quiser, mostre um resumo:
    // alert("Alterados: " + alterados + "\nIgnorados: " + ignorados);
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
      // Espaço antes dos dois-pontos
      app.findGrepPreferences = app.changeGrepPreferences = null;
      app.findGrepPreferences.findWhat = "(?<!\\s):";
      var encontrados1 = doc.findGrep();
      if (encontrados1.length > 0) {
        app.changeGrepPreferences.changeTo = " :";
        var changed1 = doc.changeGrep();
        totalAlteracoes += changed1 ? changed1.length : 0;
      }

      // Espaço depois dos dois-pontos
      app.findGrepPreferences = app.changeGrepPreferences = null;
      app.findGrepPreferences.findWhat = ":(?!\\s|\\r|\\n)";
      var encontrados2 = doc.findGrep();
      if (encontrados2.length > 0) {
        app.changeGrepPreferences.changeTo = ": ";
        var changed2 = doc.changeGrep();
        totalAlteracoes += changed2 ? changed2.length : 0;
      }
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

  /*ajuste espaçamento barra*/

  function ajustarBarrEspaco(doc, ignorarTabelasBarra) {
    var totalAlteracoes = 0;

    if (ignorarTabelasBarra) {
      var stories = doc.stories;
      for (var i = 0; i < stories.length; i++) {
        var story = stories[i];

        if (story.tables.length === 0) {
          totalAlteracoes += aplicarSeNecessario(
            story,
            "([^\\s])/([^\\s])",
            "$1 / $2",
            false
          );
        } else {
          for (var j = 0; j < story.paragraphs.length; j++) {
            var p = story.paragraphs[j];

            if (!estaDentroDeTabela(p)) {
              totalAlteracoes += aplicarSeNecessario(
                p,
                "([^\\s])/([^\\s])",
                "$1 / $2",
                false
              );
            }
          }
        }
      }
    } else {
      totalAlteracoes += aplicarSeNecessario(
        doc,
        "([^\\s])/([^\\s])",
        "$1 / $2",
        true
      );
    }

    return totalAlteracoes;
  }

  function estaDentroDeTabela(objeto) {
    try {
      return objeto.parent.constructor.name === "Cell";
    } catch (e) {
      return false;
    }
  }

  function aplicarSeNecessario(
    alvo,
    procurar,
    substituir,
    permitirDentroTabela
  ) {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = procurar;
    app.changeGrepPreferences.changeTo = substituir;

    var encontrados = alvo.findGrep();
    var total = 0;

    for (var i = 0; i < encontrados.length; i++) {
      var item = encontrados[i];

      var dentroTabela = estaDentroDeTabela(item);
      if (!permitirDentroTabela && dentroTabela) continue;

      item.changeGrep();
      total++;
    }

    return total;
  }

  ////////////////////////////

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

  function ajustarEspacoPontoFinalTabela(onFinish) {
    if (app.documents.length === 0) {
      alert("Nenhum documento aberto.");
      return;
    }

    var doc = app.activeDocument;

    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    app.findGrepPreferences.findWhat = "\\.$";

    var ocorrencias = [];

    for (var s = 0; s < doc.stories.length; s++) {
      var story = doc.stories[s];
      for (var t = 0; t < story.tables.length; t++) {
        var table = story.tables[t];
        var results = table.findGrep();
        for (var r = 0; r < results.length; r++) {
          ocorrencias.push(results[r]);
        }
      }
    }

    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    if (ocorrencias.length === 0) {
      alert("Nenhum ponto final encontrado em tabelas.");
      if (onFinish) onFinish();
      return;
    }

    var indiceAtual = 0;
    var count = 0;

    function selecionarOcorrencia(index) {
      try {
        var occ = ocorrencias[index];
        app.select(occ);

        var win = app.layoutWindows[0];
        var tf = occ.parentTextFrames[0];
        var page = tf.parentPage;
        win.activePage = page;

        win.zoom(ZoomOptions.FIT_SPREAD);
        win.zoomPercentage = 200; //modifica o valor do zoom

        //centraliza a tela
        var vb = tf.visibleBounds;
        var cx = (vb[1] + vb[3]) / 2;
        var cy = (vb[0] + vb[2]) / 2;
        win.scrollCoordinates = [cx, cy];
      } catch (e) {
        $.writeln("Erro ao selecionar ocorrência: " + e);
      }
    }

    var win = new Window("dialog", "Revisar Pontos Finais");
    win.orientation = "column";
    win.alignChildren = "center";

    var info = win.add(
      "statictext",
      undefined,
      "Ocorrência 1 de " + ocorrencias.length
    );

    var btnGroup = win.add("group");
    btnGroup.orientation = "row";

    var btnAnterior = btnGroup.add("button", undefined, "Anterior");
    var btnProximo = btnGroup.add("button", undefined, "Próximo");
    var btnExcluir = btnGroup.add("button", undefined, "Excluir");
    var btnFinalizar = btnGroup.add("button", undefined, "Finalizar");

    function atualizarInfo() {
      info.text =
        "Ocorrência " + (indiceAtual + 1) + " de " + ocorrencias.length;
    }

    btnAnterior.onClick = function () {
      if (indiceAtual > 0) {
        indiceAtual--;
        selecionarOcorrencia(indiceAtual);
        atualizarInfo();
      }
    };

    btnProximo.onClick = function () {
      if (indiceAtual < ocorrencias.length - 1) {
        indiceAtual++;
        selecionarOcorrencia(indiceAtual);
        atualizarInfo();
      }
    };

    btnExcluir.onClick = function () {
      ocorrencias[indiceAtual].contents = ocorrencias[
        indiceAtual
      ].contents.replace(/\.$/, "");
      count++;
      if (indiceAtual < ocorrencias.length - 1) {
        indiceAtual++;
        selecionarOcorrencia(indiceAtual);
        atualizarInfo();
      }
    };

    btnFinalizar.onClick = function () {
      win.close();
      if (onFinish) onFinish(count);
    };

    selecionarOcorrencia(indiceAtual);
    atualizarInfo();
    win.show();
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
