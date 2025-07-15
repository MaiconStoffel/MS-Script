(function () {
    if (app.documents.length === 0) {
        alert("Abra um documento antes de rodar o script.");
        return;
    }

    var doc = app.activeDocument;

    // Variáveis para armazenar as escolhas do usuário
    var aspasAbertura = "";
    var aspasFechamento = "";
    var aplicarAspas = false;
    var trocarPontoPorVirgula = false;
    var trocarVirgulaPorPonto = false;
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
    var contadorIgualEspaco = 0;
    var contadorEspacamentoDoisPontos = 0;
    var contadorPontoFinalEspaco = 0;
    var contadorEspacoAntesPercent = 0;
    var contadorEspacoDepoisPontoVirgula = 0;


    // Tela 1 – Escolha do idioma das aspas
    function telaAspas() {
        var win = new Window("dialog", "Passo 1: Escolha de Aspas");
        win.orientation = "column";
        win.alignChildren = "left";

        win.add("statictext", undefined, "Escolha o idioma para as aspas:");
        var idiomaGrupo = win.add("group");
        idiomaGrupo.orientation = "column";
        idiomaGrupo.alignChildren = "left";
        var idiomaPolones = idiomaGrupo.add("radiobutton", undefined, "Polonês („ ”)");
        var idiomaAlemao = idiomaGrupo.add("radiobutton", undefined, "Alemão („ “)");
        var idiomaFrances = idiomaGrupo.add("radiobutton", undefined, "Francês (« »)");
        var idiomaRusso = idiomaGrupo.add("radiobutton", undefined, "Russo (« »)");
        var idiomaPadrao = idiomaGrupo.add("radiobutton", undefined, "Padrão (“ ”)");
        var idiomaOutro  = idiomaGrupo.add("radiobutton", undefined, "Não Efetuar Alteração");
        idiomaOutro.value = true;

        var botoes = win.add("group");
        botoes.alignment = "center";
        var proximo = botoes.add("button", undefined, "Próximo", { name: "ok" });
        var cancelar = botoes.add("button", undefined, "Cancelar", { name: "cancel" });

        cancelar.onClick = function () { win.close(); };

        proximo.onClick = function () {
            if (!idiomaOutro.value) {
                aplicarAspas = true;
                if (idiomaPolones.value)  { aspasAbertura = "„"; aspasFechamento = "”"; idiomaEscolhido = "Polonês („ ”)"; }
                if (idiomaAlemao.value)   { aspasAbertura = "„"; aspasFechamento = "“"; idiomaEscolhido = "Alemão („ “)"; }
                if (idiomaFrances.value)  { aspasAbertura = "«"; aspasFechamento = "»"; idiomaEscolhido = "Francês (« »)"; }
                if (idiomaRusso.value)    { aspasAbertura = "«"; aspasFechamento = "»"; idiomaEscolhido = "Russo (« »)"; }
                if (idiomaPadrao.value)   { aspasAbertura = "“"; aspasFechamento = "”"; idiomaEscolhido = "Padrão (“ ”)"; }
            }
            win.close();
            telaDecimais();
        };

        win.center();
        win.show();
    }

    // Tela 2 – Escolha das trocas decimais
 function telaDecimais() {
    var win = new Window("dialog", "Passo 2: Troca de Decimais");
    win.orientation = "column";
    win.alignChildren = "left";

    win.add("statictext", undefined, "Escolha o tipo de troca decimal:");

    // Grupo de radiobuttons
    var grupoDecimal = win.add("group");
    grupoDecimal.orientation = "column";
    grupoDecimal.alignChildren = "left";
    var rbPontoVirgula = grupoDecimal.add("radiobutton", undefined, "Trocar ponto por vírgula decimal (3.14 → 3,14)");
    var rbVirgulaPonto = grupoDecimal.add("radiobutton", undefined, "Trocar vírgula por ponto decimal (3,14 → 3.14)");
    var rbNenhum = grupoDecimal.add("radiobutton", undefined, "Não Efetuar Alteração");
    rbNenhum.value = true; 

    var botoes = win.add("group");
    botoes.alignment = "center";
    var voltar = botoes.add("button", undefined, "Voltar");
    var proximo = botoes.add("button", undefined, "Próximo", { name: "ok" });
    var cancelar = botoes.add("button", undefined, "Cancelar", { name: "cancel" });

    cancelar.onClick = function () { win.close(); };

    voltar.onClick = function () {
        win.close();
        telaAspas();
    };

    proximo.onClick = function () {
        trocarPontoPorVirgula = rbPontoVirgula.value;
        trocarVirgulaPorPonto = rbVirgulaPonto.value;
        win.close();
        telaEspacos();  
    };

    win.center();
    win.show();
}


    function telaEspacos() {
        var win = new Window("dialog", "Passo 3: Ajuste de Espaços");
        win.orientation = "column";
        win.alignChildren = "left";

        win.add("statictext", undefined, "Escolha os ajustes de espaços para pontuação e símbolos:");

        var chkPontoFinalEspaco = win.add("checkbox", undefined, "Padronizar espaço após ponto final '. '");
        var chkVirgulaEspaco = win.add("checkbox", undefined, "Padronizar espaço após vírgula ', '");
        var chkIgualEspaco = win.add("checkbox", undefined, "Padronizar espaço antes e depois do sinal ' = '");
        var chkEspacoAntesPercent = win.add("checkbox", undefined, "Padronizar espaço antes do símbolo ' %'");
        var chkDoisPontosEspaco = win.add("checkbox", undefined, "Padronizar espaçamento dos dois-pontos ':'");
        var chkEspacoDepoisPontoVirgula = win.add("checkbox", undefined, "Padronizar espaçamento do ponto e vírgula ';'");


        var botoes = win.add("group");
        botoes.alignment = "center";
        var voltar = botoes.add("button", undefined, "Voltar");
        var proximo = botoes.add("button", undefined, "Próximo", { name: "ok" });
        var cancelar = botoes.add("button", undefined, "Cancelar", { name: "cancel" });

        cancelar.onClick = function () { win.close(); };

        voltar.onClick = function () {
            win.close();
            telaDecimais();
        };

        proximo.onClick = function () {
            ajustarVirgulaEspaco = chkVirgulaEspaco.value;
            ajustarIgualEspaco = chkIgualEspaco.value;
            ajustarEspacamentoDoisPontos = chkDoisPontosEspaco.value;
            ajustarPontoFinalEspaco = chkPontoFinalEspaco.value;
            ajustarEspacoAntesPercent = chkEspacoAntesPercent.value;
            ajustarEspacamentoPontoVirgula = chkEspacoDepoisPontoVirgula.value;
            win.close();
            telaCapitalizacaoTitulos(); 
        };

        win.center();
        win.show();
    }

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
    

   function telaCapitalizacaoTitulos() {
    var win = new Window("dialog", "Passo 4: Capitalização de Títulos");
    win.orientation = "column";
    win.alignChildren = "center";

    win.add("statictext", undefined, "Escolha o idioma e os estilos de parágrafo para aplicar a capitalização:");

    // Checkbox para ignorar capitalização
    var chkNaoAplicar = win.add("checkbox", undefined, "Não aplicar capitalização de títulos");
    chkNaoAplicar.alignment = "left";
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
    grupoCapitalizacao.add("statictext", undefined, "Tipo de capitalização:");
    var dropdownCapitalizacao = grupoCapitalizacao.add("dropdownlist", undefined, ["Maiúsculas", "Minúsculas", "Title Case"]);
    dropdownCapitalizacao.selection = 2; // default em Title Case


    // Estilos de parágrafo
    var listaEstilos = getAllParagraphStyles(doc);

    var grupoEstilos = win.add("panel", undefined, "Estilos de Parágrafo:");
    grupoEstilos.orientation = "column";
    grupoEstilos.alignChildren = "fill";

    var scrollGroup = grupoEstilos.add("group");
    scrollGroup.orientation = "column";
    scrollGroup.alignChildren = "fill";
    scrollGroup.preferredSize = [300, 150];

    var listaEstilosBox = scrollGroup.add("listbox", undefined, [], { multiselect: true });
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
    var cancelar = botoes.add("button", undefined, "Cancelar", { name: "cancel" });

    cancelar.onClick = function () { win.close(); };
    voltar.onClick = function () { win.close(); telaEspacos();};

    adicionarSigla.onClick = function () {
    abrirJanelaAdicionarSigla(); };


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

    if (!listaEstilosBox.selection || listaEstilosBox.selection.length === 0) {
        alert("Selecione ao menos um estilo ou marque 'Não aplicar capitalização de títulos'.");
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
        var win = new Window("dialog", "Passo 4: Confirmação Final");
        win.orientation = "column";
        win.alignChildren = "left";

        var resumo = "As seguintes ações serão aplicadas:\n\n";
        if (aplicarAspas) resumo += "- Correção das aspas para o idioma: " + idiomaEscolhido + "\n";
        if (trocarPontoPorVirgula) resumo += "- Trocar ponto por vírgula decimal (3.14 → 3,14)\n";
        if (trocarVirgulaPorPonto) resumo += "- Trocar vírgula por ponto decimal (3,14 → 3.14)\n";
        if (ajustarPontoFinalEspaco) resumo += "- Padronizar espaço após ponto final '. '\n";
        if (ajustarVirgulaEspaco) resumo += "- Padronizar espaço após vírgula ', '\n";
        if (ajustarIgualEspaco) resumo += "- Padronizar espaço antes e depois do sinal ' = '\n";
        if (ajustarEspacoAntesPercent) resumo += "- Padronizar espaço antes do símbolo ' %'\n";
        if (ajustarEspacamentoDoisPontos) resumo += "- Padronizar espaçamento do dois-pontos ':'\n";
        if (ajustarEspacamentoPontoVirgula) resumo += "- Padronizar espaçamento do ponto e vírgula ';'\n";
        if (idiomaSelecionado && estilosSelecionados.length > 0) {
        resumo += "- Capitalização de Títulos: \n    • Modelo: " + capitalizacaoSelecionada + "\n    • Idioma: " + idiomaSelecionado + "\n";
        resumo += "\n- Estilos selecionados:\n";
        for (var i = 0; i < estilosSelecionados.length; i++) {
        resumo += "    • " + estilosSelecionados[i] + "\n";}}



        if (
        !aplicarAspas &&
        !trocarPontoPorVirgula &&
        !trocarVirgulaPorPonto &&
        !ajustarVirgulaEspaco &&
        !ajustarIgualEspaco &&
        !ajustarEspacamentoDoisPontos &&
        !ajustarPontoFinalEspaco &&
        !ajustarEspacoAntesPercent &&
        !ajustarEspacamentoPontoVirgula &&
        !(idiomaSelecionado && estilosSelecionados.length > 0)
        ) {resumo += "- Nenhuma ação selecionada.\n";}

        var textoResumo = win.add("edittext", undefined, resumo, {multiline: true, readonly: true });
        textoResumo.preferredSize.width = 300;
        textoResumo.preferredSize.height = 180;

        var botoes = win.add("group");
        botoes.alignment = "center";
        var voltar = botoes.add("button", undefined, "Voltar");
        var executar = botoes.add("button", undefined, "Executar", { name: "ok" });
        var cancelar = botoes.add("button", undefined, "Cancelar", { name: "cancel" });

        cancelar.onClick = function () { win.close(); };

        voltar.onClick = function () {
            win.close();
            telaEspacos();
        };

        
        executar.onClick = function () {
            win.close();
            executarCorrecao();
  
        };


        win.center();
        win.show();
    }

    function executarCorrecao() {

    if (aplicarAspas) {contadorAspasAberto = corrigirAspasAbertura(aspasAbertura);
                       contadorAspasFechado = corrigirAspasFechamento(aspasFechamento);
        contadorAspas = contadorAspasAberto + contadorAspasFechado;}


    if (trocarPontoPorVirgula) contadorPontoVirgula = substituirPontoPorVirgula();
    if (trocarVirgulaPorPonto) contadorVirgulaPonto = substituirVirgulaPorPonto();

    if (ajustarVirgulaEspaco) {var contagemV1 = removerEspacoAntesVirgula();
                               var contagemV2 = corrigirEspacoDepoisVirgula();
        contadorVirgulaEspaco = contagemV1 + contagemV2;}

    if (ajustarIgualEspaco) contadorIgualEspaco = ajustarEspacoIgual();
        
    if (ajustarPontoFinalEspaco) {var contagemPF1 = removerEspacoAntesPontoFinal();   
                                  var contagemPF2 = corrigirEspacoDepoisPontoFinal();
        contadorPontoFinalEspaco = contagemPF1 + contagemPF2; }

    if (ajustarEspacoAntesPercent) contadorEspacoAntesPercent = ajustarEspacoAntesPercentual();
    
    if (ajustarEspacamentoDoisPontos) { var contagemDP1 = removerEspacoAntesDoisPontos();   
                                           var contagemDP2 = corrigirEspacoDepoisDoisPontos();
        contadorEspacamentoDoisPontos = contagemDP1 + contagemDP2; }

    if (ajustarEspacamentoPontoVirgula) { var contagemPV1 = removerEspacoAntesPontoVirgula();   
                                           var contagemPV2 = corrigirEspacoDepoisPontoVirgula();
        contadorEspacoDepoisPontoVirgula = contagemPV1 + contagemPV2; }

    if (estilosSelecionados && estilosSelecionados.length > 0) {
        contadorCapitalizacao = aplicarCapitalizacaoComGREP(estilosSelecionados, idiomaSelecionado, capitalizacaoSelecionada);
    }


    var mensagem = "Correções concluídas!\n\n";
    if (aplicarAspas) mensagem += "✅ " + contadorAspas + " - Aspas alteradas (" + contadorAspasAberto + " de abertura e " + contadorAspasFechado + " de fechamento).\n";
    if (trocarPontoPorVirgula) mensagem += "✅ " + contadorPontoVirgula + " - Pontos trocados por vírgula.\n";
    if (trocarVirgulaPorPonto) mensagem += "✅ " + contadorVirgulaPonto + " - Vírgulas trocadas por ponto.\n";
    if (ajustarVirgulaEspaco) mensagem += "✅ " + contadorVirgulaEspaco + " - Espaços após vírgula ajustados.\n";
    if (ajustarIgualEspaco) mensagem += "✅ " + contadorIgualEspaco + " - Espaços em torno do sinal '=' ajustados.\n";
    if (ajustarEspacamentoDoisPontos) mensagem += "✅ " + contadorEspacamentoDoisPontos + " - Espaçamentos dos dois-pontos ajustados.\n";
    if (ajustarPontoFinalEspaco) mensagem += "✅ " + contadorPontoFinalEspaco + " - Espaços após ponto final ajustados.\n";
    if (ajustarEspacoAntesPercent) mensagem += "✅ " + contadorEspacoAntesPercent + " - Espaços antes do '%' ajustados.\n";
    if (ajustarEspacamentoPontoVirgula) mensagem += "✅ " + contadorEspacoDepoisPontoVirgula + " - Espaçamentos do ';' ajustados. \n";
    if (contadorCapitalizacao > 0) {mensagem += "✅ " + contadorCapitalizacao + " - Títulos capitalizados: \n       -  Idioma: " + idiomaSelecionado + ", \n       -  Modelo: " + capitalizacaoSelecionada + ", \n       -  Estilos: " + estilosSelecionados.join(", ") + "\n";}

if (
    !aplicarAspas &&
    !trocarPontoPorVirgula &&
    !trocarVirgulaPorPonto &&
    !ajustarVirgulaEspaco &&
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
    var win = new Window("dialog", "Resultado Final!");
    win.orientation = "column";
    win.alignChildren = "left";

    win.add("statictext", undefined, "Execução concluída com sucesso!");

    var textoResumo = win.add("edittext", undefined, resumoExecucao, { multiline: true, readonly: true });
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

function corrigirAspasAbertura(aspasAbertura) {
    app.findGrepPreferences = app.changeGrepPreferences = null;

    var aspasAberturaEscapada = aspasAbertura.replace(/([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])/g, '\\$1');

    app.findGrepPreferences.findWhat = '(?<=[\\s\\(\\[\\{—])["“”„«]';
    app.changeGrepPreferences.changeTo = aspasAberturaEscapada;

    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;

    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
}

function corrigirAspasFechamento(aspasFechamento) {
    app.findGrepPreferences = app.changeGrepPreferences = null;

    var aspasFechamentoEscapada = aspasFechamento.replace(/([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])/g, '\\$1');

    app.findGrepPreferences.findWhat = '(?<![\\s\\(\\[\\{—])["“”„»]';
    app.changeGrepPreferences.changeTo = aspasFechamentoEscapada;

    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;

    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
}


//FUNCOES DE EXECUCAO, FAZEM AS ALTERACOES NO ARQUIVO UTILIZANDO GREP


function substituirPontoPorVirgula() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "(?<=\\d)\\.(?=\\d)";
    app.changeGrepPreferences.changeTo = ",";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
}

function substituirVirgulaPorPonto() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "(?<=\\d)\\,(?=\\d)";
    app.changeGrepPreferences.changeTo = ".";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
}

function removerEspacoAntesVirgula() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\s+,";
    app.changeGrepPreferences.changeTo = ",";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
}

function corrigirEspacoDepoisVirgula() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = ",(?!\\s|\\d)(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])";
    app.changeGrepPreferences.changeTo = ", ";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
}


function ajustarEspacoIgual() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\s*=\\s*";
    app.changeGrepPreferences.changeTo = " = ";
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
    app.findGrepPreferences.findWhat = ":(?! )(?=[A-Za-zÀ-ÖØ-öø-ÿ0-9])"; // : sem espaço depois
    app.changeGrepPreferences.changeTo = ": ";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
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



    function ajustarEspacoAntesPercentual() {
    app.findGrepPreferences = app.changeGrepPreferences = null;
    app.findGrepPreferences.findWhat = "\\s*(%)";
    app.changeGrepPreferences.changeTo = " $1";
    var changed = doc.changeGrep();
    var count = changed ? changed.length : 0;
    app.findGrepPreferences = app.changeGrepPreferences = null;
    return count;
    }

    function abrirJanelaAdicionarSigla() {
    var caminho = "Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/SiglasPadrao.txt";

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
    var cancelar = botoes.add("button", undefined, "Cancelar", { name: "cancel" });

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
    };


    function carregarSiglasDoArquivo() {
    try {
        var arquivo = File("Q:/GROUPS/BR_SC_JGS_WAU_DESENVOLVIMENTO_PRODUTOS/Documentos dos Produtos/Manuais dos Produtos/MS-SCRIPT/SiglasPadrao.txt");

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



    function aplicarCapitalizacaoComGREP(estilosSelecionados, idiomaSelecionado, capitalizacaoSelecionada) {
    try {
        var doc = app.activeDocument;
        var paragrafos = doc.stories.everyItem().paragraphs.everyItem().getElements();

        if (!estilosSelecionados || !(estilosSelecionados instanceof Array)) {
            alert("estilosSelecionados não é um array válido.");
            return 0;
        }

        var excecoes = {
            "Português": ["de", "na", "pra", "da", "do", "das", "dos", "e", "em", "com", "por", "para", "a", "o", "as", "os", "nos"],
            "Inglês": ["a", "an", "and", "but", "or", "for", "nor", "as", "at", "by", "in", "of", "on", "per", "to", "the", "with"],
            "Polonês": ["i", "oraz", "a", "ale", "lecz", "lub", "czy", "nie", "w", "na", "do", "z", "od", "pod", "nad", "przed", "po", "bez", "dla", "o", "u"],
            "Francês": ["à", "au", "aux", "avec", "chez", "dans", "de", "des", "du","en", "et", "la", "le", "les", "ou", "par", "pour", "sur", "un", "une"]

        };

        var palavrasMin = excecoes[idiomaSelecionado] || [];
        var siglasMaiusculas = carregarSiglasDoArquivo();

        function estaNaLista(palavra, lista) {
        for (var i = 0; i < lista.length; i++) {
        if (palavra.toLowerCase() === lista[i].toLowerCase()) {
            return true;
        } } return false;}


        // Definir tipo de capitalização
        var modoChangecase = null;
        if (capitalizacaoSelecionada === "Maiúsculas") {
            modoChangecase = ChangecaseMode.UPPERCASE;
        } else if (capitalizacaoSelecionada === "Minúsculas") {
            modoChangecase = ChangecaseMode.SENTENCECASE;
        } else if (capitalizacaoSelecionada === "Title Case") {
            modoChangecase = "TITULO_MANUAL";
        } else {
            alert("Tipo de capitalização não reconhecido.");
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

                for (var k = 0; k < palavras.length; k++) {
                    var palavraOriginal = palavras[k].contents;
                    var palavraTexto = palavraOriginal.toLowerCase();

                    var siglaEncontrada = detectarSiglaPresente(palavraOriginal, siglasMaiusculas);

                if (siglaEncontrada) {
                var novaPalavra = palavraOriginal.replace(
                 new RegExp("^" + siglaEncontrada, "i"),
                 siglaEncontrada.toUpperCase()
                );
                palavras[k].contents = novaPalavra;
                } else if (k === 0 || !estaNaLista(palavraTexto, palavrasMin)) {
                 palavras[k].changecase(ChangecaseMode.TITLECASE);
                } else {
                palavras[k].changecase(ChangecaseMode.LOWERCASE);}

                }
            } else {
                p.changecase(modoChangecase);
            }

            contador++;
        }

        return contador;

    } catch (e) {
        alert("Erro ao aplicar capitalização:\n" + e.message);
        return 0;
    }
}





























    telaAspas();

})();
