var doc = app.activeDocument;
var win = app.activeWindow;
var hyperlinks = doc.hyperlinks;

if (hyperlinks.length === 0) {
    alert("Nenhum hiperlink encontrado.");
} else {
    for (var i = 0; i < hyperlinks.length; i++) {
        var link = hyperlinks[i];
        var nome = link.name;
        var tipo = "Desconhecido";
        var destinoInfo = "";

        try {
            var destino = link.destination;

            if (destino instanceof HyperlinkURLDestination) {
                tipo = "URL";
                destinoInfo = destino.destinationURL;

            } else if (destino instanceof HyperlinkTextDestination) {
                tipo = "Texto";
                destinoInfo = destino.destinationText.contents;

                // Forçar navegação ao texto
                destino.destinationText.show();
                var pagina = destino.destinationText.parentTextFrames[0].parentPage;
                if (pagina) {
                    win.activePage = pagina;
                }

            } else if (destino instanceof HyperlinkPageDestination) {
                tipo = "Página";
                destinoInfo = "Página " + (destino.destinationPage.documentOffset + 1);
                win.activePage = destino.destinationPage;
            } else {
                destinoInfo = "Tipo de destino não reconhecido.";
            }

        } catch (e) {
            destinoInfo = "Destino inválido ou fora do escopo.";
        }

        alert(
            "🔗 Hiperlink " + (i + 1) + " de " + hyperlinks.length + ":\n\n" +
            "📌 Nome: " + nome + "\n" +
            "📎 Tipo: " + tipo + "\n" +
            "🎯 Destino: " + destinoInfo + "\n\n" +
            "Pressione OK para continuar."
        );
    }
}
