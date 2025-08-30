/**
 * Set-ZintPath.jsx
 * Utilitário para configurar/substituir o caminho do zint.exe salvo.
 */
(function () {
  var current = app.extractLabel("ZINT_EXE_PATH") || "(não definido)";
  if (!confirm("Caminho atual do zint.exe:\n\n" + current + "\n\nDeseja alterar?")) return;

  var exe = File.openDialog("Selecione o zint.exe", "Executável:*.exe");
  if (!exe) return;

  if (!/zint\.exe$/i.test(exe.fsName)) {
    alert("Selecione o arquivo zint.exe (não a pasta).");
    return;
  }

  app.insertLabel("ZINT_EXE_PATH", exe.fsName);
  alert("zint.exe atualizado para:\n" + exe.fsName);
})();
