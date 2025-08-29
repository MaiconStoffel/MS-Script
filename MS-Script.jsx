/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * ^file           : MS-Script
  * ^brief          : Script for InDesign Automations
  * ^author 	      : mschwertz / 2025
  ******************************************************************************
  /* USER CODE END Header */

  var developerInfo = "Desenvolvido por: @mschwertz - Versão 1.0"; 
  
//#region to-do
//                        TODO - @mschwertz

  //* - [✓] - Arrumar funcoes de troca de decimal (as duas)
  //* - [✓] - Colocar o espanhol, russo, italiano, turco e holandes como opcao de idioma na Capitalizacao nas excessoes
  //* - [✓] - Arrumar a parte das siglas "de", "para" e tal da Capitalização, verificar pra colocar em um txt ?
  //* - [✓] - Tirar o botao de contabilizar o tempo e fazer como padrao
  //* - [✓] - Verificar a funcao de Capitalização para pegar arquivos de Programação que nao possuem o estilo.
  //* - [✓] - Tela de execução quando chama a funcao de "Aplicação de Estilos"
  //* - [✓] - Revisar função de capitalização para textos dentro de tabelas, nao está mudando.
  //* - [✓] - Verificar se vale a pena colocar aquela tela de execucao quando clicar em executar.
  //* - [✓] - Verificar pra colocar a função de Ponto Final em Tabela na soma de tempo total -> Passar ela pela tela final e executar
  //* - [✓] - Verificar o motivo da funcao de capitalização aparece 0 no tempo salvo
  //! - [X] - Verificar de colocar a funcao da aplicação de estilos na soma de tempo total - Não tem motivo de passar ela pela tela final e nem executar, mas podemos
  //!        fazer ela mostrar a quantidade de mudanças na tela de resultado final e fazer o calculo do tempo salvo
   
  // TODO - [] - Verificar se é possível criar uma função que gere o codigo 2D pra colocar no documento e que passe o 2D direto para a pasta de vínculos do documento 
  //^            Testar com o aplicativo "Zint", pelo visto ele cria o codigo 2d datamatrix, podemos automatizar para ele pegar o numero que o 
  //^            usuário digitar e receber na volta o png do codigo


  // TODO - [] - Passar funcao por funcao fazendo uma verificacao do codigo e vendo se funciona tudo certinho
   
//#endregion
  
(function () {  
  ("highlight Force Decorate");
  if (app.documents.length === 0) {
    alert("Abra um documento antes de rodar script.");
    return;
  }

  //#region Variáveis 
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
  
  var ContabilizarTempo = true;
  var contagensIniciais = {}; 

  var ajustarPontoFinalTabela = false;
  var ignorarTabelasBarra = false;
  var ajustarEspacamentoDoisPontos_Antes = false;
  var ignorarExclusaoParenteses = false;
  var ignorarExclusaoColchetes = false;
  var ajustarBarraEspaco = false;
  var ajustarVirgulaEspaco = false;
  var ajustarIgualEspaco = false;
  var ajustarEspacamentoDoisPontos = false;
  var ajustarPontoFinalEspaco = false;
  var ajustarEspacoAntesPercent = false;
  var ajustarEspacamentoPontoVirgula = false;
  var idiomaEscolhido = "";

  //? Capitalizacao

  var idiomaSelecionado = null; //eh diferente do idiomaEscolhido
  var capitalizacaoSelecionada = null;
  var estilosSelecionados = [];
  var contadorCapitalizacao = 0;

  //? Variáveis contagem

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
  var contadorPontoFinalemTabela = 0;

  //? AplicaçãoEstilos

  var prefixoEscolhido = "";
  var corEscolhida = "";
  var fonteEscolhida = "";
  var estiloEscolhido = "";
  var tamanhoEscolhido = "";

  //? variaveis da troca do decimal

  var dadosTempoSalvo = null;
  var trocarPontoPorVirgula = false;
  var trocarVirgulaPorPonto = false;
  var paginasSelecionadasDecimal = ""; // var q conteem o que o usuario escreveu na troca decimal, porém mais filtrada e arrumada
  var aplicarEmTodasAsPaginas = true; // padrao pra aplicar a mudança decimal em todas as pgns
  var paginasArray = []; // array q conteem as paginas que o usuario digitou
  var inputPaginas; // var pra receber o que o usuario escreveu nas pgns da troca decimal

  //#endregion

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
      coluna1,
      "✓ Aplicação de Estilos ✓",
      function () {
       telaAplicacaoEstilos ();
      },
      "Filtra textos com base nas definições do usuário e aplica estilos de parágrafos a ele"
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

    criarBotao(
      coluna2,
      "● Ponto Final em Tabela ●",
      function () {
      ajustarPontoFinalTabela = true;
      telaFinal();
      },
      "Mostra pontos finais dentro de tabelas para o usuário decidir se quer mantê-los ou excluí-los"
    );

    // Novo botão centralizado abaixo dos dois últimos botões
    var grupoCentral = janela.add("group");
    grupoCentral.orientation = "row";
    grupoCentral.alignment = "center"; // Centraliza no meio da janela

    criarBotao(
      grupoCentral,
      "📊 Gerar Código 2D 📊",
      function () {
        alert("Novo botão centralizado clicado!");
      },
      "Gera o código 2D e insere automaticamente na pasta vínculos e no seu documento"
    );

    janela.add(
      "statictext",
      undefined,
      developerInfo
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
    idiomaPolones.helpTip = "Irá alterar as aspas do documento para o formato Polonês: „ ”";
    var idiomaAlemao = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Alemão („ “)"
    );
    idiomaAlemao.helpTip = "Irá alterar as aspas do documento para o formato Alemão: „ “";
    var idiomaFrances = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Francês (« »)"
    );
    idiomaFrances.helpTip = "Irá alterar as aspas do documento para o formato Francês: « »";

    var idiomaRusso = idiomaGrupo.add("radiobutton", undefined, "Russo (« »)");
    idiomaRusso.helpTip = "Irá alterar as aspas do documento para o formato Russo: « »";
    var idiomaPadrao = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Padrão (“ ”)"
    );
    idiomaPadrao.helpTip = "Irá alterar as aspas do documento para o formato Padrão: “ ”";

    var idiomaOutro = idiomaGrupo.add(
      "radiobutton",
      undefined,
      "Não Efetuar Alteração"
    );
    idiomaOutro.value = true;
    idiomaOutro.helpTip = "Não irá efetuar nenhuma alteração nas aspas do documento.";

    var linha = win.add("panel");
    linha.alignment = "fill";

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
    rbPontoVirgula.helpTip = "Se marcado, o script irá trocar todos os pontos decimais (3.14) por vírgulas (3,14).";

    var rbVirgulaPonto = grupoDecimal.add(
      "radiobutton",
      undefined,
      "Trocar vírgula por ponto decimal (3,14 → 3.14)"
    );
    rbVirgulaPonto.helpTip = "Se marcado, o script irá trocar todas as vírgulas decimais (3,14) por pontos (3.14).";

    var rbNenhum = grupoDecimal.add(
      "radiobutton",
      undefined,
      "Não Efetuar Alteração"
    );
    rbNenhum.value = true;
    rbNenhum.helpTip = "Se marcado, nenhuma alteração será feita nos decimais do documento.";

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
      "Aplicar em todas as páginas?"
    );
    chkTodasPaginas.value = true;
    chkTodasPaginas.helpTip = "Se marcado, a troca decimal será aplicada em todas as páginas do documento.";

    txtPgnsCenter.add(
      "statictext",
      undefined,
      "Páginas para a mudança (ex: 1 - 3 , 5 , 8):"
    );
    inputPaginas = win.add("edittext", undefined, "");
    inputPaginas.characters = 30;
    inputPaginas.enabled = false;

    // Quando o checkbox muda, habilita/desabilita o campo
    chkTodasPaginas.onClick = function () {
      inputPaginas.enabled = !chkTodasPaginas.value;
    };

    var linha = win.add("panel");
    linha.alignment = "fill";

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

    var chkSelecionarTodos = win.add(
      "checkbox",
      undefined,
      "Selecionar todas as padrozinações abaixo"
    );
    chkSelecionarTodos.helpTip = "Se marcado, todas as padronizações abaixo serão aplicadas ao documento.";

    var chkPontoFinalEspaco = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Ponto → '.  '"
    );
    chkPontoFinalEspaco.helpTip = "Padroniza o espaçamento após o ponto final, deixando um espaço após o ponto.";

    var chkVirgulaEspaco = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento da Vírgula → ',  '."
    );
    chkVirgulaEspaco.helpTip = "Padroniza o espaçamento após a vírgula, deixando um espaço após a vírgula.";

    var chkIgualEspaco = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Igual → '  =  '"
    );
    chkIgualEspaco.helpTip = "Padroniza o espaçamento antes e depois do sinal de igual, deixando um espaço antes e depois do sinal.";

    var grupo2P = win.add("group");
    grupo2P.orientation = "row";
    grupo2P.alignChildren = "left";

    var chkDoisPontosEspaco = grupo2P.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Dois-Pontos → ':  '"
    );
    chkDoisPontosEspaco.helpTip = "Padroniza o espaçamento após o dois-pontos, deixando um espaço após o símbolo.";

    var chk2PEspacoAntes = grupo2P.add("checkbox", undefined, "Espaço Antes?");
    chk2PEspacoAntes.helpTip =
      "Se marcado, o espaçamento será aplicado também antes do Dois-Pontos, como no exemplo: '  :  '";

    var chkEspacoDepoisPontoVirgula = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento do Ponto e Vírgula → ';  '"
    );
    chkEspacoDepoisPontoVirgula.helpTip = "Padroniza o espaçamento após o ponto e vírgula, deixando um espaço após o símbolo.";

    var chkEspacoAntesPercent = win.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento da Porcentagem → '  %  '"
    );
    chkEspacoAntesPercent.helpTip = "Padroniza o espaçamento antes do símbolo de porcentagem, deixando um espaço antes e depois do símbolo.";

    var grupoBarra = win.add("group");
    grupoBarra.orientation = "row";
    grupoBarra.alignment = "fill"; // ocupa a largura total
    grupoBarra.alignChildren = ["left", "center"];

    // Checkbox esquerda
    var chkBarraEspaco = grupoBarra.add(
      "checkbox",
      undefined,
      "Padronizar espaçamento da Barra → '  /  '"
    );
    chkBarraEspaco.helpTip = "Padroniza o espaçamento antes e depois da barra, deixando um espaço antes e depois do símbolo.";

    // Espaçador "elástico"
    var espacoFlex = grupoBarra.add("statictext", undefined, "");
    espacoFlex.alignment = "fill";

    // Checkbox direita
    var chkTabelas = grupoBarra.add("checkbox", undefined, "Ignorar Tabelas?");
    chkTabelas.helpTip =
      "Se marcado, a padronização da barra '  /  ' será aplicada apenas fora de tabelas.";

    var BancoDadosBarra = win.add("group");
    BancoDadosBarra.orientation = "row";
    BancoDadosBarra.alignment = "right";

    var BancoDadosBarra = BancoDadosBarra.add(
      "button",
      undefined,
      "Adicionar Siglas?"
    );
    BancoDadosBarra.preferredSize = [125, 20]; // ajusta largura e altura
    BancoDadosBarra.helpTip = 
      "Se marcado, o usuário poderá adicionar siglas, ao banco de dados, para não serem afetadas pela padronização da barra '  /  '.";

    var linha = win.add("panel");
    linha.alignment = "fill";
    
    
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

    BancoDadosBarra.onClick = function () {
      abrirJanelaAdicionarSiglaBarra();
    }

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
      "     Escolha o idioma e os estilos de parágrafo para aplicar a alteração de caixa::::::"
    );

    var linha = win.add("panel");
    linha.alignment = "fill";

    // Checkbox para ignorar Alteração de Caixa
    var chkNaoAplicar = win.add(
      "checkbox",
      undefined,
      "Não aplicar alteração de caixa"
    );
    chkNaoAplicar.helpTip = "Se marcado, nenhuma alteração de caixa será aplicada no documento.";
    chkNaoAplicar.alignment = "center";
    chkNaoAplicar.value = true;

    // Dropdown de idiomas
    var idiomas = ["Português", "Inglês", "Espanhol", "Francês", "Alemão", "Italiano", "Polonês", "Russo", "Holandês", "Turco"];
    var grupoIdioma = win.add("group");
    grupoIdioma.add("statictext", undefined, "Idioma:");
    grupoIdioma.alignment = "center";
    var dropdownIdioma = grupoIdioma.add("dropdownlist", undefined, idiomas);
    dropdownIdioma.helpTip = "Selecione o idioma para a alteração de caixa.";
    dropdownIdioma.selection = 0;

    // Dropdown de estilos
    var grupoCapitalizacao = win.add("group");
    grupoCapitalizacao.orientation = "row";
    grupoCapitalizacao.alignment = "center";
    grupoCapitalizacao.add("statictext", undefined, "Alterar a caixa para  →");
    var dropdownCapitalizacao = grupoCapitalizacao.add(
      "dropdownlist",
      undefined,
      ["Maiúsculas", "Minúsculas", "Title Case"]
    );
    dropdownCapitalizacao.helpTip = "Selecione o tipo de caixa para a alteração.";
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
     // if (listaEstilos[i] !== "[Parágrafo padrão]") {
        listaEstilosBox.add("item", listaEstilos[i]);
      //}
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
    adicionarSigla.helpTip = "Abre a janela para adicionar siglas ao banco de dados, para que não sejam afetadas pela alteração de caixa e ficarem em maiúsculo.";
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

    if (ajustarPontoFinalTabela)
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
      !ajustarPontoFinalTabela &&
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
    executar.helpTip = "Inicia o processo de correção do documento com base nas opções selecionadas.";
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
      if (ajustarPontoFinalTabela) {
        win.close();

        var ops = {
          ajustarPontoFinalTabela:ajustarPontoFinalTabela,
          };

          if (ContabilizarTempo) {
            contagensIniciais = contarOcorrencias(app.activeDocument, ops);

            // TESTE: exibe os caracteres encontrados
            var testeResultado = "";
            for (var key in contagensIniciais) {
              alert("entrou");
              testeResultado += key + ": " + contagensIniciais[key] + "\n";
            }
          }

          fncAjustarPontoFinalTabela(function (count) {
            contadorPontoFinalemTabela = count;

            var mensagem = "\n                    Correções concluídas!\n\n";
            if (contadorPontoFinalemTabela)
              mensagem +=
                "✅ " +
                contadorPontoFinalemTabela +
                " - Ajustes em Pontos finais dentro de tabelas\n";

            // --- MONTA CONTADORES ---
            var contadoresAlteracoes = {
              "Pontos finais removidos em tabelas": contadorPontoFinalemTabela,
            };

            // --- SALVA TEMPO ---
            if (ContabilizarTempo) {
              dadosTempoSalvo = salvarTempo(
                contadoresAlteracoes,
                contagensIniciais,
                {
                  ajustarPontoFinalTabela: true,
                }
              );
            }

            // --- MOSTRA RESULTADO FINAL ---
            telaResultadoFinal(mensagem, contadoresAlteracoes);
          });


        return; // evita continuar antes do usuário finalizar
      } else {
        win.close();

        // Define quais operações o usuário selecionou
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

        if (ContabilizarTempo) {
          contagensIniciais = contarOcorrencias(app.activeDocument, ops);

          // TESTE: exibe os caracteres encontrados
          var testeResultado = "";
          for (var key in contagensIniciais) {
            alert("entrou");
            testeResultado += key + ": " + contagensIniciais[key] + "\n";
          }
        }

        app.doScript(
          executarCorrecao,
          ScriptLanguage.JAVASCRIPT,
          undefined,
          UndoModes.ENTIRE_SCRIPT,
          "Correções Projeto TRI"
        );
      }
    };

    win.center();
    win.show();
  }

  function executarCorrecao() {
    alertaTemporario("Iniciando as correções... ⌛", 1000);
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
      contadorCapitalizacao = aplicarCapitalizacaoComGREP(estilosSelecionados,idiomaSelecionado,capitalizacaoSelecionada);
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

    var contadoresAlteracoes = {
      "Aspas alteradas": contadorAspas,
      "Pontos trocados por vírgula": contadorPontoVirgula,
      "Vírgulas trocadas por ponto": contadorVirgulaPonto,
      "Ajustes em ' = '": contadorIgualEspaco,
      "Ajustes em ' : '": contadorEspacamentoDoisPontos,
      "Ajustes em ' ; '": contadorEspacoDepoisPontoVirgula,
      "Ajustes em ' % '": contadorEspacoAntesPercent,
      "Ajustes em ' / '": contadorBarraEspaco,
      "Alterações de Caixa": contadorCapitalizacao,
    };

    var dadosTempoSalvo = salvarTempo(contadoresAlteracoes, contagensIniciais, {
      aplicarAspas: aplicarAspas,
      trocarPontoPorVirgula: trocarPontoPorVirgula,
      trocarVirgulaPorPonto: trocarVirgulaPorPonto,
      ajustarEspacamentoDoisPontos: ajustarEspacamentoDoisPontos,
      ajustarEspacamentoPontoVirgula: ajustarEspacamentoPontoVirgula,
      ajustarBarraEspaco: ajustarBarraEspaco,
      ajustarIgualEspaco: ajustarIgualEspaco,
      ajustarEspacoAntesPercent: ajustarEspacoAntesPercent,
      aplicarCapitalizacao: (contadorCapitalizacao > 0), // 👈 agora com o nome certo

    });

    telaResultadoFinal(mensagem, contadoresAlteracoes);
  }

  function telaResultadoFinal(resumoExecucao,contadoresAlteracoes) {
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
      if (!ContabilizarTempo) {
        alert(
          "Para ver o tempo salvo, volte à tela inicial e selecione a opção 'Contabilizar Tempo Salvo'."
        );
        return;
      }
      mostrarJanelaTempo(
        contadoresAlteracoes,
        contagensIniciais,
        dadosTempoSalvo,
        );
    };



    win.center();
    win.show();
  }

  //#region Tempo Salvo
  //*----------------- FUNCAO DE CALCULO DO TEMPO SALVO -----------------

  var op2Greps = {
    trocarPontoPorVirgula: ["\\."],
    trocarVirgulaPorPonto: ["\\,"],
    ajustarEspacamentoDoisPontos: ["\\:"],
    ajustarEspacamentoPontoVirgula: ["\\;"],
    ajustarBarraEspaco: ["\\/"],
    ajustarIgualEspaco: ["\\="],
    ajustarEspacoAntesPercent: ["\\%"],
    aplicarAspas: ['"'],
  };

  var op2LabelContagem = {
    trocarPontoPorVirgula: "Quantidade de pontos",
    trocarVirgulaPorPonto: "Quantidade de vírgulas",
    ajustarEspacamentoDoisPontos: "Quantidade de ':'",
    ajustarEspacamentoPontoVirgula: "Quantidade de ';'",
    ajustarBarraEspaco: "Quantidade de '/'",
    ajustarIgualEspaco: "Quantidade de '='",
    ajustarEspacoAntesPercent: "Quantidade de '%'",
    aplicarAspas: "Quantidade de aspas",
    ajustarPontoFinalTabela: "Quantidade de pontos finais em tabelas",
    aplicarCapitalizacao: "Quantidade de títulos candidatos",
  };

  var op2LabelAlteracao = {
    trocarPontoPorVirgula: "Pontos trocados por vírgula",
    trocarVirgulaPorPonto: "Vírgulas trocadas por ponto",
    ajustarEspacamentoDoisPontos: "Ajustes em ' : '",
    ajustarEspacamentoPontoVirgula: "Ajustes em ' ; '",
    ajustarBarraEspaco: "Ajustes em ' / '",
    ajustarIgualEspaco: "Ajustes em ' = '",
    ajustarEspacoAntesPercent: "Ajustes em ' % '",
    aplicarAspas: "Aspas alteradas",
    ajustarPontoFinalTabela: "Pontos finais removidos em tabelas",
    aplicarCapitalizacao: "Alterações de Caixa",
  };

  function contarOcorrencias(doc, ops) {
  var resultado = {};
  if (!doc) doc = app.activeDocument; // fallback

  try {
    // zera prefs antes de começar
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    for (var op in op2Greps) {
      if (!op2Greps.hasOwnProperty(op)) continue;

      // se 'ops' foi passado, conta só as operações ativas
      if (ops && !ops[op]) continue;

      var total = 0;

      for (var i = 0; i < op2Greps[op].length; i++) {
        app.findGrepPreferences.findWhat = op2Greps[op][i];
        var encontrados = doc.findGrep();
        total += encontrados.length;
      }

      var label = op2LabelContagem[op] || op;
      resultado[label] = total;

      // limpa o find pra próxima iteração
      app.findGrepPreferences = NothingEnum.nothing;
    }
  } catch (e) {
    alert("Erro ao contar ocorrências: " + e);
  } finally {
    // SEMPRE limpar no final
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
  }

  return resultado;
  }

  function alertaTemporario(msg, tempoMs) {
  var win = new Window("palette", "Aviso");  
  win.add("statictext", undefined, msg);  
  win.show();  

  $.sleep(tempoMs || 1500); // espera (1.5s padrão)
  win.close();
  }

  function salvarTempo(contadoresAlteracoes, contagensIniciais, ops) {
  var tarefas = [
    { op: "trocarPontoPorVirgula" },
    { op: "trocarVirgulaPorPonto" },
    { op: "ajustarEspacamentoDoisPontos" },
    { op: "ajustarEspacamentoPontoVirgula" },
    { op: "ajustarBarraEspaco" },
    { op: "ajustarIgualEspaco" },
    { op: "ajustarEspacoAntesPercent" },
    { op: "aplicarAspas" },
    { op: "aplicarCapitalizacao" },
    { op: "ajustarPontoFinalTabela" },
  ];

  var tarefasAtivas = [];
  for (var t = 0; t < tarefas.length; t++) {
    if (ops && ops[tarefas[t].op]) tarefasAtivas.push(tarefas[t]);
  }

  var tempoCapitalizacao = 17;
  var tempoNaoAlterado = 3;
  var tempoAlterado = 8;
  var tempoTotalSegundos = 0;

  for (var x = 0; x < tarefasAtivas.length; x++) {
    var tarefa = tarefasAtivas[x];
    var labelContagem = op2LabelContagem[tarefa.op];
    var labelAlteracao = op2LabelAlteracao[tarefa.op];

    var totalTarefa = contagensIniciais[labelContagem] || 0;
    var alteradas = contadoresAlteracoes[labelAlteracao] || 0;
    var naoAlteradas = Math.max(totalTarefa - alteradas, 0);

    if (tarefa.op === "aplicarCapitalizacao") {
      // cálculo especial para capitalização
      tempoTotalSegundos += alteradas * tempoCapitalizacao;
    } else {
      // cálculo padrão
      tempoTotalSegundos += naoAlteradas * tempoNaoAlterado + alteradas * tempoAlterado;
    }
  }

  var tempoTotalMinutos = tempoTotalSegundos / 60;
  var tempoTotalHoras = tempoTotalMinutos / 60;

  // --- SALVAR TEMPO ACUMULADO EM BD ---
  var caminhoArquivo =
    "Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/TempoSalvo.txt";
  var horasAnteriores = lerTempoAcumulado(caminhoArquivo);
  var horasAtualizadas = horasAnteriores + tempoTotalHoras;
  salvarTempoAcumulado(caminhoArquivo, horasAtualizadas);

  dadosTempoSalvo = {
    segundos: tempoTotalSegundos,
    minutos: tempoTotalMinutos,
    horas: tempoTotalHoras,
    horasAcumuladas: horasAtualizadas,
  };

  return dadosTempoSalvo;
  }

  function mostrarJanelaTempo(contadoresAlteracoes, contagensIniciais, dados) {
  alertaTemporario("Carregando tempo salvo...", 500);

  contagensIniciais = contagensIniciais || {};
  contadoresAlteracoes = contadoresAlteracoes || {};
  dados = dados ||
    dadosTempoSalvo || {
      segundos: 0,
      minutos: 0,
      horas: 0,
      horasAcumuladas: 0,
    };

  var winTempoSalvo = new Window("dialog", "Resultado Tempo Salvo ⏱");
  winTempoSalvo.orientation = "column";
  winTempoSalvo.alignChildren = "center";

  var resultado =
    "----- Quantidade de caracteres antes das alterações -----\n\n";
  for (var key in contagensIniciais) {
    resultado += key + ": " + contagensIniciais[key] + "\n";
  }

  resultado +=
    "\n------------------------- Alterações feitas -------------------------\n\n";

  var algumaAlteracao = false;
  for (var k in contadoresAlteracoes) {
    if (contadoresAlteracoes[k] > 0) {
      resultado += "✅ " + contadoresAlteracoes[k] + " - " + k + "\n";
      algumaAlteracao = true;
    }
  }
  if (!algumaAlteracao) {
    resultado += "- Nenhuma alteração feita.\n";
  }

  var segundos = Number(dados.segundos) || 0;
  var minutos = Number(dados.minutos) || 0;
  var horas = Number(dados.horas) || 0;
  var horasAcumuladas = Number(dados.horasAcumuladas) || 0;

  resultado +=
    "\n------------ Tempo total estimado economizado ------------\n\n" +
    "- " +
    segundos.toFixed(0) +
    " segundos\n" +
    "- " +
    minutos.toFixed(2) +
    " minutos\n" +
    "- " +
    horas.toFixed(2) +
    " horas\n";

  resultado +=
    "\n-------------------- Tempo total acumulado --------------------\n\n" +
    "- " +
    horasAcumuladas.toFixed(2) +
    " horas\n\n";

  var caixaResultado = winTempoSalvo.add("edittext", undefined, resultado, {
    multiline: true,
    readonly: true,
  });
  caixaResultado.size = [320, 300];

  var fechar = winTempoSalvo.add("button", undefined, "Fechar");
  fechar.onClick = function () {
    winTempoSalvo.close();
  };

  winTempoSalvo.show();
  }

  function lerTempoAcumulado(caminhoArquivo) {
    var arquivo = new File(caminhoArquivo);
    if (arquivo.exists) {
      arquivo.open("r");
      var conteudo = arquivo.read();
      arquivo.close();

      // Extrai apenas o número do texto
      var match = conteudo.match(/([\d\.]+)/); // pega números e ponto decimal
      if (match) {
        return parseFloat(match[1]);
      } else {
        return 0;
      }
    } else {
      return 0;
    }
  }

  function salvarTempoAcumulado(caminhoArquivo, novoTempo) {
    var arquivo = new File(caminhoArquivo);
    arquivo.open("w"); // sobrescreve
    arquivo.write(
      "Aproximadamente " + novoTempo.toFixed(2) + " horas foram salvas ao todo."
    );
    arquivo.close();
  }

  //#region Aspas
  //*--------------------- FUNCOES DE CORRECAO DE ASPAS --------------------------------

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

  //#region Decimais
  //*---------------------------- FUNCOES TROCA DE DECIMAIS ----------------------------

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

    var alvosTexto = [];

    if (aplicarEmTodasAsPaginas) {
      var todasStories = doc.stories.everyItem().getElements();
      for (var s = 0; s < todasStories.length; s++) {
        var story = todasStories[s];
        for (var p = 0; p < story.paragraphs.length; p++) {
          alvosTexto.push(story.paragraphs[p]);
        }
      }
    } else {
      for (var i = 0; i < paginasArray.length; i++) {
        var pagina = doc.pages[paginasArray[i] - 1];
        if (!pagina) continue;

        var textFrames = pagina.textFrames;
        for (var j = 0; j < textFrames.length; j++) {
          var story = textFrames[j].parentStory;
          for (var p = 0; p < story.paragraphs.length; p++) {
            alvosTexto.push(story.paragraphs[p]);
          }
        }
      }
    }

    // ---- CRIA A JANELA DE PROGRESSO ----
    var w = new Window("palette", "Processando alterações...");
  
    // Texto superior
    var aviso = w.add(
      "statictext",
      undefined,
      " Deixe o InDesign aberto, isso pode levar um tempo."
    );
    aviso.alignment = "center";

    // Grupo para texto de status centralizado
    var groupTexto = w.add("group");
    groupTexto.alignment = "fill";
    groupTexto.alignChildren = "center";
    var texto = groupTexto.add("statictext", undefined, "Iniciando...");
    texto.characters = 25;
    w.show();

    // ---- LOOP DE PROCESSAMENTO ----
    for (var t = 0; t < alvosTexto.length; t++) {
      var textoAlvo = alvosTexto[t];
      var matches = textoAlvo.findGrep();

      for (var k = matches.length - 1; k >= 0; k--) {
        var ponto = matches[k];

        var idx = ponto.insertionPoints[0].index;
        var baseText = ponto.parentStory.contents;

        try {
          if (
            ponto.parent &&
            ponto.parent.constructor &&
            ponto.parent.constructor.name === "Cell"
          ) {
            baseText = ponto.parent.contents;
            idx = ponto.index;
          }
        } catch (e) {}

        if (deveIgnorarPorExcecaoComBase(baseText, idx)) {
          ignorados++;
          continue;
        }

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

        // Pequeno respiro a cada 500 matches
        if (k % 500 === 0) {
          $.sleep(10);
        }
      }

      texto.text = "             Processando bloco " + (t + 1) + " de " + alvosTexto.length;
      w.update();
    }

    w.close();

    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

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

    var alvosTexto = [];

    if (aplicarEmTodasAsPaginas) {
      var todasStories = doc.stories.everyItem().getElements();
      for (var s = 0; s < todasStories.length; s++) {
        var story = todasStories[s];
        for (var p = 0; p < story.paragraphs.length; p++) {
          alvosTexto.push(story.paragraphs[p]);
        }
      }
    } else {
      for (var i = 0; i < paginasArray.length; i++) {
        var pagina = doc.pages[paginasArray[i] - 1];
        if (!pagina) continue;

        var textFrames = pagina.textFrames;
        for (var j = 0; j < textFrames.length; j++) {
          var story = textFrames[j].parentStory;
          for (var p = 0; p < story.paragraphs.length; p++) {
            alvosTexto.push(story.paragraphs[p]);
          }
        }
      }
    }

    // ---- CRIA A JANELA DE PROGRESSO ----
    var w = new Window("palette", "Processando alterações...");
  
    // Texto superior
    var aviso = w.add(
      "statictext",
      undefined,
      " Deixe o InDesign aberto, isso pode levar um tempo."
    );
    aviso.alignment = "center";

    // Grupo para texto de status centralizado
    var groupTexto = w.add("group");
    groupTexto.alignment = "fill";
    groupTexto.alignChildren = "center";
    var texto = groupTexto.add("statictext", undefined, "Iniciando...");
    texto.characters = 25;
    w.show();

    // Percorre SOMENTE os textos-alvo
    for (var t = 0; t < alvosTexto.length; t++) {
      var textoAlvo = alvosTexto[t];
      var matches = textoAlvo.findGrep();

      for (var k = matches.length - 1; k >= 0; k--) {
        var virgula = matches[k];

        var idx = virgula.insertionPoints[0].index;
        var baseText = virgula.parentStory.contents;

        try {
          if (
            virgula.parent &&
            virgula.parent.constructor &&
            virgula.parent.constructor.name === "Cell"
          ) {
            baseText = virgula.parent.contents;
            idx = virgula.index;
          }
        } catch (e) {}

        if (deveIgnorarPorExcecaoComBase(baseText, idx)) {
          ignorados++;
          continue;
        }

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

        // ---- RESPIRO a cada 500 itens ----
        if (k % 500 === 0) {
          $.sleep(10);
        }
      }

      texto.text = "             Processando bloco " + (t + 1) + " de " + alvosTexto.length;
      w.update();
    }

    w.close();

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

  //#region Gerais
  //*-------------------------------- GERAIS -------------------------------------------
  


  //#region Vírgula
  //*------------------------- Espaçamento Vírgula ------------------------------------

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

  //#region Dois-Pontos
  //*------------------------- Espaçamento Dois-Pontos ----------------------------------

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

  //#region Ponto-Vírgula
  //*------------------------ Espaçamento Ponto-Vírgula ----------------------------------

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

  //#region Barra
  //*--------------------------------- Espaçamento Barra ---------------------------------

  function abrirJanelaAdicionarSiglaBarra() {
    var caminho =
      "Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/SiglasBarra.txt";

    function salvarSiglaBarra(sigla) {
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
    win.alignChildren = "center";

    win.add("statictext", undefined, "    Digite a sigla que deseja adicionar:    .");
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
      salvarSiglaBarra(sigla);
      win.close();
    };

    cancelar.onClick = function () {
      win.close();
    };

    win.center();
    win.show();
  }

  function carregarSiglasDoArquivoBarra() {
    try {
      var arquivo = File(
        "Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/SiglasBarra.txt"
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

  function detectarSiglaPresenteBarra(palavra, siglas) {
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

  function ajustarBarrEspaco(doc, ignorarTabelasBarra) {
  var totalAlteracoes = 0;
  var siglas = carregarSiglasDoArquivoBarra(); // carrega banco

  if (ignorarTabelasBarra) {
    var stories = doc.stories;
    for (var i = 0; i < stories.length; i++) {
      var story = stories[i];

      if (story.tables.length === 0) {
        totalAlteracoes += aplicarSeNecessario(
          story,
          "([^\\s])/([^\\s])",
          "$1 / $2",
          false,
          siglas
        );
      } else {
        for (var j = 0; j < story.paragraphs.length; j++) {
          var p = story.paragraphs[j];

          if (!estaDentroDeTabela(p)) {
            totalAlteracoes += aplicarSeNecessario(
              p,
              "([^\\s])/([^\\s])",
              "$1 / $2",
              false,
              siglas
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
      true,
      siglas
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

  function aplicarSeNecessario(alvo, procurar, substituir, permitirDentroTabela, siglas) {
  app.findGrepPreferences = app.changeGrepPreferences = null;
  app.findGrepPreferences.findWhat = procurar;
  app.changeGrepPreferences.changeTo = substituir;

  var encontrados = alvo.findGrep();
  var total = 0;

  for (var i = 0; i < encontrados.length; i++) {
    var item = encontrados[i];

    var dentroTabela = estaDentroDeTabela(item);
    if (!permitirDentroTabela && dentroTabela) continue;

    // 🔹 Aqui: pegar a palavra completa em volta da barra
    var palavra = item.contents;
    var siglaEncontrada = detectarSiglaPresenteBarra(palavra, siglas);

    if (siglaEncontrada) {
      // Ignora substituição se for sigla
      continue;
    }

    item.changeGrep();
    total++;
  }

  return total;
  }

  //#region Igual (=)
  //*------------------------------- Espaçamento Igual (=) -------------------------------
  
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

    return (changed1 ? changed1.length : 0);
  }

  //#region Ponto Final
  //*------------------------------ Espaçamento Ponto-Final ------------------------------

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

  //#region Percentual(%)
  //*----------------------------- Espaçamento Percentual (%) -----------------------------


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

    return (changed1 ? changed1.length : 0);
  }

  //#region Ponto Tabela
  //* ------------------------ PONTO FINAL EM TABELA ------------------------

  function fncAjustarPontoFinalTabela(onFinish) {
    alertaTemporario ("Buscando pontos finais dentro de tabelas...", 1500)
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
        win.zoomPercentage = 200; 

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
    info.characters = 25;
    info.justify = "center";

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

  //#region Capitalização
  //* ---------------------------- CAPITALIZAÇÃO ----------------------------

  function getAllParagraphStyles(doc) {
  var estilos = [];
  function coletar(estilosLista) {
    for (var i = 0; i < estilosLista.length; i++) {
      var estilo = estilosLista[i];
      // NÃO filtra mais "[Parágrafo padrão]" nem outros entre colchetes
      estilos.push(estilo.name);
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
  win.alignChildren = "center";

  win.add(
    "statictext",
    undefined,
    "    Digite a sigla que deseja adicionar:    ."
  );
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

  function carregarExcecoesDoArquivo(caminho, idiomaSelecionado) {
    try {
      var arquivo = File(caminho);
      if (!arquivo.exists) {
        alert("Arquivo de exceções não encontrado.");
        return [];
      }

      arquivo.encoding = "UTF-8";
      arquivo.open("r");
      var conteudo = arquivo.read();
      arquivo.close();

      conteudo = conteudo.replace(/\uFEFF/g, ""); // Remove BOM
      conteudo = conteudo.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

      var linhas = conteudo.split("\n");
      var excecoes = [];
      var idiomaAtual = null;

      for (var i = 0; i < linhas.length; i++) {
        var linha = linhas[i];

        if (typeof linha !== "string") continue;
        linha = linha.replace(/^\s+|\s+$/g, ""); // trim

        if (!linha || /^([#;]|\/\/)/.test(linha)) continue; // ignora comentários

        if (/^\[.+\]$/.test(linha)) {
          // Cabeçalho de idioma
          idiomaAtual = linha.replace(/^\s*\[|\]\s*$/g, "");
        } else if (idiomaAtual === idiomaSelecionado) {
          excecoes.push(linha);
        }
      }

      return excecoes;
    } catch (e) {
      return [];
    }
  }

  function getParagraphsIncludingTables(doc) {
    var out = [];
    var seen = {};

    var stories = doc.stories.everyItem().getElements();
    for (var s = 0; s < stories.length; s++) {
        var story = stories[s];

        // Parágrafos "normais" do story
        var pars = story.paragraphs.everyItem().getElements();
        for (var p = 0; p < pars.length; p++) {
            var spec = pars[p].toSpecifier();
            if (!seen[spec]) { seen[spec] = true; out.push(pars[p]); }
        }

        // Parágrafos dentro de TABELAS
        var tables = story.tables.everyItem().getElements();
        for (var t = 0; t < tables.length; t++) {
            var cells = tables[t].cells.everyItem().getElements();
            for (var c = 0; c < cells.length; c++) {
                var cpars = cells[c].paragraphs.everyItem().getElements();
                for (var cp = 0; cp < cpars.length; cp++) {
                    var spec2 = cpars[cp].toSpecifier();
                    if (!seen[spec2]) { seen[spec2] = true; out.push(cpars[cp]); }
                }
            }
        }
    }
    return out;
  }

  function aplicarCapitalizacaoComGREP(
    estilosSelecionados,
    idiomaSelecionado,
    capitalizacaoSelecionada
    
  ) {
    try {
      var doc = app.activeDocument;
      var paragrafos = getParagraphsIncludingTables(doc);


      if (!estilosSelecionados || !(estilosSelecionados instanceof Array)) {
        alert("estilosSelecionados não é um array válido.");
        return 0;
      }
      var siglasMaiusculas = carregarSiglasDoArquivo();


      // Carrega as exceções apenas do idioma selecionado
      var palavrasMin = carregarExcecoesDoArquivo(
        "Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/SiglasExcecoes.txt",
        idiomaSelecionado
      );

      if (palavrasMin.length === 0) {
        alert("Nenhuma exceção encontrada para o idioma: " + idiomaSelecionado);
      } 

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

  //* ---------------------------- APLICAÇÃO DE ESTILO ----------------------------

  function telaAplicacaoEstilos() {
    alertaTemporario("Carregando... ⌛", 2500);
    var win = new Window("dialog", "Aplicação de Estilos");
    win.orientation = "column";
    win.alignChildren = "fill";
    win.preferredSize = [500, 350];

    var grupoTexto = win.add("group");
    grupoTexto.alignment = "center";
    grupoTexto.add(
      "statictext",
      undefined,
      "Escolha os filtros a serem utilizados na aplicação de estilos."
    );

    var linha = win.add("panel");
    linha.alignment = "fill";

    // // Checkbox para ignorar Alteração de Caixa
    // var chkNaoAplicarEstilos = win.add(
    //   "checkbox",
    //   undefined,
    //   "Não aplicar aplicação de estilos"
    // );

    //*-----------------------------------------------------------------------------------------------------------------------------------------------------------

    var grupoPrexifo = win.add("panel", undefined, "");
    grupoPrexifo.orientation = "column";
    grupoPrexifo.alignChildren = "fill";

    var tituloGrupo = grupoPrexifo.add("group");
    tituloGrupo.alignment = "center";
    tituloGrupo.add("statictext", undefined, "    Escreva um prefixo:    .");
    var prefixInput = grupoPrexifo.add("edittext", undefined, "");
    prefixInput.characters = 20;
    prefixInput.alignment = "center";
    prefixInput.selection = 0;
    prefixInput.helpTip = "Selecione a cor para aplicar.";
    prefixInput.preferredSize = [150, 20];

    //*-----------------------------------------------------------------------------------------------------------------------------------------------------------

    var cores = getAllColors(doc);

    var grupoCores = win.add("panel", undefined, "");
    grupoCores.orientation = "column";
    grupoCores.alignChildren = "fill";

    var tituloGrupo = grupoCores.add("group");
    tituloGrupo.alignment = "center";
    tituloGrupo.add("statictext", undefined, "    Selecione a cor:    .");

    var ddCores = grupoCores.add(
      "dropdownlist",
      undefined,
      ["Nenhuma"].concat(cores)
    );
    ddCores.alignment = "center";
    ddCores.selection = 0;
    ddCores.helpTip = "Selecione a cor para aplicar.";
    ddCores.preferredSize = [150, 20];

    //*-----------------------------------------------------------------------------------------------------------------------------------------------------------

    var familias = getAllFonts(doc);

    var grupoFE = win.add("panel", undefined, "");
    grupoFE.orientation = "column";
    grupoFE.alignChildren = "fill";

    var tituloGrupo = grupoFE.add("group");
    tituloGrupo.alignment = "center";
    tituloGrupo.add(
      "statictext",
      undefined,
      "    Selecione a fonte e estilo:    ."
    );

    var ddFamilia = grupoFE.add(
      "dropdownlist",
      undefined,
      ["Nenhuma"].concat(familias)
    );
    
    ddFamilia.alignment = "center";
    ddFamilia.selection = 0;
    ddFamilia.helpTip = "Selecione a família da fonte.";
    ddFamilia.preferredSize = [150, 20];

    var ddEstilo = grupoFE.add("dropdownlist", undefined, ["Qualquer"]);
    ddEstilo.alignment = "center";
    ddEstilo.selection = 0;
    ddEstilo.helpTip = "Selecione o estilo da fonte.";
    ddEstilo.preferredSize = [150, 20];
    ddEstilo.enabled = false;

    var grupoTamanho = grupoFE.add("group");
    grupoTamanho.alignment = "center";
    grupoTamanho.add("statictext", undefined, "Tamanho da fonte:");
    var inputTamanho = grupoTamanho.add("edittext", undefined, "");
    inputTamanho.characters = 5;
    inputTamanho.helpTip = "Digite o tamanho da fonte em pontos (ex: 12)";
    inputTamanho.preferredSize = [35, 20];

    ddFamilia.onChange = function () {

      ddEstilo.removeAll();

      if (ddFamilia.selection && ddFamilia.selection.index > 0) {
        var fam = ddFamilia.selection.text;
        var estilos = getFontStylesForFamily(doc, fam);
        ddEstilo.add("item", "Qualquer");
        for (var i = 0; i < estilos.length; i++)
          ddEstilo.add("item", estilos[i]);
        ddEstilo.selection = 0;
        ddEstilo.enabled = true; 
      } else {
        ddEstilo.add("item", "Qualquer");
        ddEstilo.selection = 0;
        ddEstilo.enabled = false;
      }
    };

    //*-----------------------------------------------------------------------------------------------------------------------------------------------------------

    var grupoEstilos = win.add("panel", undefined, "");
    grupoEstilos.orientation = "column";
    grupoEstilos.alignChildren = "center";

    grupoEstilos.add("statictext", undefined, "Aplicar estilo:");

    var todosEstilos = getAllParagraphStyles2(doc); // pega todos os estilos, inclusive dentro de grupos
    var estiloDropdown = grupoEstilos.add("dropdownlist", undefined, []);
    estiloDropdown.alignment = "center";
    for (var j = 0; j < todosEstilos.length; j++) {
      estiloDropdown.add("item", todosEstilos[j].name);
    }
    estiloDropdown.selection = 0;
    grupoEstilos.preferredSize = [150, 20];
    //*-----------------------------------------------------------------------------------------------------------------------------------------------------------

    var grupoManual = win.add("panel", undefined, "");
    grupoManual.orientation = "column";
    grupoManual.alignChildren = "center";

    var chkManual = grupoManual.add(
      "checkbox",
      undefined,
      "Fazer revisão manual"
    );
    chkManual.helpTip = "Permite revisão manual antes de aplicar os estilos.";

    //*-----------------------------------------------------------------------------------------------------------------------------------------------------------

    // Botões
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
      win.close();
      prefixoEscolhido = prefixInput.text !== "" ? prefixInput.text : null;

      corEscolhida =
        ddCores.selection && ddCores.selection.index > 0
          ? ddCores.selection.text
          : null;

      fonteEscolhida =
        ddFamilia.selection && ddFamilia.selection.index > 0
          ? ddFamilia.selection.text
          : null;

      estiloEscolhido =
        ddEstilo.selection && ddEstilo.selection.index > 0
          ? ddEstilo.selection.text
          : null;

      tamanhoEscolhido =
        inputTamanho.text !== "" ? parseFloat(inputTamanho.text) : null;

      var estiloParaAplicar = estiloDropdown.selection
        ? todosEstilos[estiloDropdown.selection.index]
        : null;

      // Guarda em variável global
      EstiloParagrafoEscolhido =
        estiloParaAplicar && estiloParaAplicar.isValid
          ? estiloParaAplicar
          : null;

      var modoManual = !!chkManual.value;

      aplicarMudancas(
        doc,
        prefixoEscolhido,
        corEscolhida,
        fonteEscolhida,
        estiloEscolhido,
        tamanhoEscolhido,
        EstiloParagrafoEscolhido,
        modoManual
      );

      win.close();
    };

    win.center();
    win.show();
  }

  function numberingTextOf(par) {
    try {
      var s = par.bulletsAndNumberingResultText || "";
      if (s && s.charAt(s.length - 1) === "\t") s = s.substr(0, s.length - 1);
      return s;
    } catch (e) {
      return "";
    }
  }

  function buildRegexFromUserPrefix(pfx) {
    if (!pfx) return null;
    var esc = pfx.replace(/([\\^$.*+?()[\\]{}|])/g, "\\$1"); // escapa metacaracteres
    esc = esc.replace(/X+/gi, "\\\\d+"); // XX → \d+
    return new RegExp("^" + esc, "i"); // ancorado ao início, case-insensitive
  }

  function getAllParagraphStyles2(doc) {
    var estilos = [];

    function percorrerGrupo(grupo) {
      for (var i = 0; i < grupo.paragraphStyles.length; i++) {
        estilos.push(grupo.paragraphStyles[i]);
      }
      for (var j = 0; j < grupo.paragraphStyleGroups.length; j++) {
        percorrerGrupo(grupo.paragraphStyleGroups[j]);
      }
    }

    percorrerGrupo(doc);
    return estilos;
  }

  function getAllFonts(doc) {
    var out = [];
    var seen = {};
    var fnts = doc.fonts;

    for (var i = 0; i < fnts.length; i++) {
      var fam = fnts[i].fontFamily;
      if (fam && !seen[fam]) {
        out.push(fam);
        seen[fam] = true;
      }
    }
    out.sort();
    return out;
  }

  function getFontStylesForFamily(doc, family) {
    var out = [];
    var seen = {};
    var fnts = doc.fonts;

    for (var i = 0; i < fnts.length; i++) {
      var f = fnts[i];
      if (f.fontFamily === family) {
        // Algumas versões expõem fontStyleName; outras, styleName/fontStyle
        var sty = f.fontStyleName || f.styleName || f.fontStyle;
        if (sty && !seen[sty]) {
          out.push(sty);
          seen[sty] = true;
        }
      }
    }
    out.sort();
    return out;
  }

  function getAllColors(doc) {
    var cores = [];
    try {
      for (var i = 0; i < doc.swatches.length; i++) {
        var sw = doc.swatches[i];

        // Evita pegar [None], [Registration], etc.
        if (
          sw.name !== "[None]" &&
          sw.name !== "[Registration]" &&
          sw.name !== "[Paper]"
        ) {
          cores.push(sw.name);
        }
      }
    } catch (e) {
      alert("Erro ao obter cores: " + e);
    }
    return cores;
  }

  function safeStartsWith(s, prefix) {
    if (!s || !prefix) return false;
    return s.substr(0, prefix.length) === prefix;
  }

  function firstCharOfParagraph(par) {
    try {
      return par.characters && par.characters.length > 0
        ? par.characters[0]
        : null;
    } catch (e) {
      return null;
    }
  }

  function getFontFamilyAndStyleFromChar(ch) {
    if (!ch) return { family: null, style: null };
    var f = ch.appliedFont;
    var family = null,
      style = null;
    try {
      family = f.fontFamily;
    } catch (e) {}
    try {
      style = f.fontStyleName || f.styleName || f.fontStyle;
    } catch (e) {}
    return { family: family, style: style };
  }

  function revisarManualmente(lista, estiloParagrafo) {
    var idx = 0,
      alterados = 0,
      pulados = 0;
    var aplicarRestantes = false;

    var win = new Window("dialog", "Revisão manual – Aplicação de Estilos");
    win.orientation = "column";
    win.alignment = "center";
    win.preferredSize = [120, 150];

    // Cabeçalho / status
    var grpHeader = win.add("group");
    grpHeader.orientation = "row";
    grpHeader.alignChildren = "center";
    var lblStatus = grpHeader.add("statictext", undefined, "");
    lblStatus.characters = 35;
    lblStatus.justify = "center";
    
    // Info de numeração e estilo alvo
    var grpInfo = win.add("group");
    grpInfo.orientation = "column";
    grpInfo.alignChildren = "center";
    var lblNumero = grpInfo.add("statictext", undefined, "");
    lblNumero.characters = 35;
    lblNumero.justify = "center";

    var lblEstilo = grpInfo.add("statictext", undefined, "");
    lblEstilo.characters = 35;
    lblEstilo.justify = "center";

    // Botões
    var grpBtns = win.add("group");
    grpBtns.alignment = "center";
    var btnAplicar = grpBtns.add("button", undefined, "Aplicar");
    var btnPular = grpBtns.add("button", undefined, "Pular");
    var btnAplicarRest = grpBtns.add("button", undefined, "Aplicar restantes");
    var btnCancelar = grpBtns.add("button", undefined, "Cancelar", {
      name: "cancel",
    });

    function focarParagrafo(p, zoomPct) {
      try {
        if (!p) return;

        // Janela: prefere layoutWindows (mais previsível)
        var win =
          app.layoutWindows && app.layoutWindows.length
            ? app.layoutWindows[0]
            : app.activeWindow;

        // Determina o TextFrame "principal" onde o parágrafo realmente começa/está visível
        var ip =
          p.insertionPoints && p.insertionPoints.length
            ? p.insertionPoints[0]
            : null;
        var mainTF =
          ip && ip.parentTextFrames && ip.parentTextFrames.length
            ? ip.parentTextFrames[0]
            : p.parentTextFrames && p.parentTextFrames.length
            ? p.parentTextFrames[0]
            : null;

        if (!mainTF) return;

        // Define a página/folha antes de qualquer seleção (evita o piscar)
        if (mainTF.parentPage) {
          win.activePage = mainTF.parentPage;
        } else if (mainTF.parent && mainTF.parent.typename === "Spread") {
          win.activeSpread = mainTF.parent;
        }

        try {
          win.zoomPercentage = zoomPct || 125;

          // centraliza no parágrafo
          app.select(p);
          app.selection[0].showText();
        } catch (e) {}

        // Calcula o centro onde devemos rolar
        var cx, cy;
        var frames =
          p.parentTextFrames && p.parentTextFrames.length
            ? p.parentTextFrames
            : [mainTF];

        // Se todos os frames estão na mesma página, usamos a união das bounds para centralizar melhor
        var samePage = true;
        for (var i = 0; i < frames.length; i++) {
          if (
            !frames[i].parentPage ||
            frames[i].parentPage != mainTF.parentPage
          ) {
            samePage = false;
            break;
          }
        }

        var bounds;
        if (frames.length > 1 && samePage) {
          bounds = frames[0].visibleBounds.slice(); // [top, left, bottom, right]
          for (var j = 1; j < frames.length; j++) {
            var vb = frames[j].visibleBounds;
            bounds[0] = Math.min(bounds[0], vb[0]); // top
            bounds[1] = Math.min(bounds[1], vb[1]); // left
            bounds[2] = Math.max(bounds[2], vb[2]); // bottom
            bounds[3] = Math.max(bounds[3], vb[3]); // right
          }
        } else {
          // Apenas o frame principal
          bounds = mainTF.visibleBounds;
        }

        cx = (bounds[1] + bounds[3]) / 2; // (left + right) / 2
        cy = (bounds[0] + bounds[2]) / 2; // (top + bottom) / 2

        try {
          win.scrollCoordinates = [cx, cy];
        } catch (e) {}

        try {
          app.select(p);
        } catch (e) {}
      } catch (err) {
        $.writeln("Erro em focarParagrafo: " + err);
      }
    }

    function carregar(i) {
      var p = lista[i];
      focarParagrafo(p);

      var num = numberingTextOf(p);
      var cont = "";
      try {
        cont = p.contents || "";
      } catch (e) {}

      lblStatus.text = "Ocorrência " + (i + 1) + " de " + lista.length;
      lblNumero.text = num ? "Numeração: " + num : "Sem numeração automática";
      lblEstilo.text = "Aplicar estilo: " + estiloParagrafo.name;
    }

    btnAplicar.onClick = function () {
      try {
        lista[idx].applyParagraphStyle(estiloParagrafo, true);
        alterados++;
      } catch (e) {
        pulados++;
      }
      idx++;
      if (idx >= lista.length) {
        win.close();
      } else {
        carregar(idx);
      }
    };

    btnPular.onClick = function () {
      pulados++;
      idx++;
      if (idx >= lista.length) {
        win.close();
      } else {
        carregar(idx);
      }
    };

    btnAplicarRest.onClick = function () {
      aplicarRestantes = true;
      win.close();
    };

    btnCancelar.onClick = function () {
      win.close();
    };

    carregar(0);
    win.center();
    win.show();

    if (aplicarRestantes && idx < lista.length) {
      for (var k = idx; k < lista.length; k++) {
        try {
          lista[k].applyParagraphStyle(estiloParagrafo, true);
          alterados++;
        } catch (e) {
          pulados++;
        }
      }
    } else {
      if (idx < lista.length) {
        pulados += lista.length - idx;
      }
    }

    return { alterados: alterados, ignorados: pulados };
  }

  function aplicarMudancas(
    doc,
    prefixoEscolhido,
    corEscolhida,
    fonteEscolhida,
    estiloEscolhido,
    tamanhoEscolhido,
    EstiloParagrafoEscolhido,
    modoManual
  ) {
    app.doScript(
      function () {
        var total = 0,
          alterados = 0,
          ignorados = 0;
        var candidatos = [];

        function processarParagrafo(paragrafo) {
          total++;
          var texto = "";
          try {
            texto = paragrafo.contents || "";
          } catch (e) {}

          var atende = true;

          // --- filtro por prefixo
          if (prefixoEscolhido && atende) {
            var numTxt = numberingTextOf(paragrafo);
            var re = buildRegexFromUserPrefix(prefixoEscolhido);
            if (numTxt) {
              if (re && !re.test(numTxt)) atende = false;
            } else {
              if (!safeStartsWith(texto, prefixoEscolhido)) atende = false;
            }
          }

          var ch = firstCharOfParagraph(paragrafo);

          // --- filtro por fonte/estilo
          if (atende && (fonteEscolhida || estiloEscolhido)) {
            var fs = getFontFamilyAndStyleFromChar(ch);
            if (fonteEscolhida && fs.family !== fonteEscolhida) atende = false;
            if (estiloEscolhido && fs.style !== estiloEscolhido) atende = false;
          }

          // --- filtro por tamanho
          if (atende && tamanhoEscolhido != null) {
            var pt = null;
            try {
              pt = ch ? ch.pointSize : null;
            } catch (e) {}
            if (pt !== tamanhoEscolhido) atende = false;
          }

          // --- filtro por cor
          if (atende && corEscolhida) {
            var colorName = null;
            try {
              colorName =
                ch && ch.fillColor && ch.fillColor.name
                  ? ch.fillColor.name
                  : null;
            } catch (e) {}
            if (colorName !== corEscolhida) atende = false;
          }

          // --- aplicar estilo
          if (
            atende &&
            EstiloParagrafoEscolhido &&
            EstiloParagrafoEscolhido.isValid
          ) {
            if (modoManual) {
              candidatos.push(paragrafo);
            } else {
              try {
                paragrafo.applyParagraphStyle(EstiloParagrafoEscolhido, true);
                alterados++;
              } catch (e) {
                ignorados++;
              }
            }
          } else {
            ignorados++;
          }
        }

        // ----------------- LOOP PRINCIPAL -----------------
        for (var i = 0; i < doc.stories.length; i++) {
          var story = doc.stories[i];

          // Parágrafos normais
          for (var j = 0; j < story.paragraphs.length; j++) {
            processarParagrafo(story.paragraphs[j]);
          }

          // Parágrafos dentro de tabelas
          for (var t = 0; t < story.tables.length; t++) {
            var tabela = story.tables[t];
            for (var l = 0; l < tabela.cells.length; l++) {
              var celula = tabela.cells[l];
              for (var p = 0; p < celula.paragraphs.length; p++) {
                processarParagrafo(celula.paragraphs[p]);
              }
            }
          }
        }

        // Se for manual, abre a revisão
        if (modoManual && candidatos.length > 0) {
          var res = revisarManualmente(candidatos, EstiloParagrafoEscolhido);
          alterados += res.alterados;
          ignorados += res.ignorados;
        }

        // alert(
        //   "Aplicação de estilos concluída!\n\n" +
        //     "Total de parágrafos analisados: " +
        //     total +
        //     "\n" +
        //     "Parágrafos alterados: " +
        //     alterados +
        //     "\n" +
        //     "Parágrafos ignorados: " +
        //     ignorados
        // );
      },
      ScriptLanguage.JAVASCRIPT,
      undefined,
      UndoModes.ENTIRE_SCRIPT,
      "Aplicação de Estilos"
    );
  }

  //#region End
  //* ---------------------------- End ----------------------------
  

  TelaInicial();
})();
