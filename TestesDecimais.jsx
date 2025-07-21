function substituirDecimalIgnorandoEstilos() {
    if (app.documents.length === 0) {
        alert("Nenhum documento aberto.");
        return;
    }

    var doc = app.activeDocument;

    // Solicita ao usuário qual substituição deseja fazer
    var opcao = confirm("Clique em OK para substituir ponto (.) por vírgula (,).\nClique em Cancelar para substituir vírgula (,) por ponto (.)");

    var findPattern = opcao ? "(?<=\\d)\\.(?=\\d)" : "(?<=\\d),(?=\\d)";
    var substituirPor = opcao ? "," : ".";

    // Limpa preferências antes de iniciar
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = findPattern;

    var resultados = doc.findGrep();

    var alterados = 0;
    var ignorados = 0;

    for (var i = 0; i < resultados.length; i++) {
        var match = resultados[i];

        var fontStyle = "";
        var isBold = false;
        var isItalic = false;

        try {
            fontStyle = match.appliedFont.fontStyleName.toLowerCase();
            isBold = fontStyle.indexOf("bold") !== -1;
            isItalic = fontStyle.indexOf("italic") !== -1 || fontStyle.indexOf("oblique") !== -1;
        } catch (e) {
            // Se der erro, assume que não é bold nem italic
        }

        if (!isBold && !isItalic) {
            match.contents = substituirPor;
            alterados++;
        } else {
            ignorados++;
        }
    }

    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;

    alert(
        "Substituição concluída!\n\n" +
        "Total de ocorrências encontradas: " + resultados.length + "\n" +
        "Substituídos: " + alterados + "\n" +
        "Ignorados (negrito ou itálico): " + ignorados
    );
}

substituirDecimalIgnorandoEstilos();
