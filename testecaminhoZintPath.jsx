// Zint-Check.jsx
(function () {
  var key = "ZINT_EXE_PATH";
  var path = app.extractLabel(key) || "(não definido)";
  var ok = (path && File(path).exists) ? "SIM" : "NÃO";
  alert(
    "Label: " + key +
    "\nCaminho salvo:\n" + path +
    "\n\nArquivo existe? " + ok
  );
})();
