(function () {
  var zint = app.extractLabel("ZINT_EXE_PATH");
  if (!zint) { alert("ZINT_EXE_PATH não definido."); return; }
  if (!File(zint).exists) { alert("zint.exe não encontrado:\n" + zint); return; }

  var outPng = Folder.temp.fsName + "/dm_smoketest.png";
  var logTxt = Folder.temp.fsName + "/dm_smoketest.log";
  // Monta a linha de comando com aspas corretas
  var cmd = '"' + zint + '" --barcode=datamatrix --data "M123456" --output "' + outPng + '" --scale 6 --notext';

  // VBScript: executa e grava ExitCode em LOG
  var vb = ''
    + 'Dim sh: Set sh = CreateObject("WScript.Shell")' + "\r\n"
    + 'Dim rc: rc = sh.Run("' + cmd.replace(/"/g, '""') + '", 0, True)' + "\r\n"
    + 'Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")' + "\r\n"
    + 'Dim f: Set f = fso.CreateTextFile("' + logTxt.replace(/\\/g,'\\\\') + '", True)' + "\r\n"
    + 'f.WriteLine "ExitCode=" & rc' + "\r\n"
    + 'f.Close' + "\r\n";

  app.doScript(vb, ScriptLanguage.VISUAL_BASIC);

  // Aguarda arquivo (8s)
  var f = new File(outPng), waited = 0, step = 200;
  while (!f.exists && waited < 8000) { $.sleep(step); waited += step; }

  var msg = "Via VBScript.\n\nPNG:\n" + outPng + "\nLOG:\n" + logTxt + "\n\n";
  if (f.exists) {
    alert(msg + "OK! Gerado.");
    f.execute();
  } else {
    var log = new File(logTxt), txt = "";
    if (log.exists) { log.open("r"); txt = log.read(); log.close(); }
    alert(msg + "Falhou.\n\nLOG:\n" + txt);
  }
})();
