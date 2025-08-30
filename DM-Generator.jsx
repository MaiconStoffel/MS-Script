/**
 * DM-Generator.jsx (VBScript edition)
 * Gera Data Matrix (PNG) via Zint e coloca no documento ativo (200x200 pt).
 * Conteúdo: 'M' + <numero_do_documento>.
 *
 * Requer: zint.exe (Zint CLI). Caminho salvo em app.insertLabel("ZINT_EXE_PATH", <fsName>).
 */

(function(){
  if (app.documents.length === 0) { alert("Abra e salve um documento antes de executar."); return; }
  var doc = app.activeDocument;

  // ---------- Utils ----------
  function getZintPath() {
    var p = app.extractLabel("ZINT_EXE_PATH");
    return (p && File(p).exists) ? p : null;
  }
  function ensureLinksFolder(doc) {
    if (!doc.saved) return null;
    var base = new Folder(doc.fullName.parent.fsName + "/Links");
    if (!base.exists) base.create();
    return base.exists ? base : null;
  }
  function waitFile(path, timeoutMs) {
    var f = new File(path), waited = 0, step = 200;
    while (!f.exists && waited < timeoutMs) { $.sleep(step); waited += step; }
    return f.exists;
  }
  function placeOnActivePage(doc, filePath, targetPx) {
    var page = app.layoutWindows.length ? app.layoutWindows[0].activePage : doc.pages[0];
    var gb = page.bounds; // [y1,x1,y2,x2]
    var x = gb[1] + 10, y = gb[0] + 10;
    var frame = page.rectangles.add();
    var sidePt = targetPx || 200; // 200 px ~ 200 pt
    frame.geometricBounds = [y, x, y + sidePt, x + sidePt];
    var placed = frame.place(new File(filePath));
    frame.fit(FitOptions.PROPORTIONALLY);
    frame.fit(FitOptions.FRAME_TO_CONTENT);
    // força moldura 200x200 mantendo proporção
    var cy = (frame.geometricBounds[0] + frame.geometricBounds[2]) / 2;
    var cx = (frame.geometricBounds[1] + frame.geometricBounds[3]) / 2;
    frame.geometricBounds = [cy - sidePt/2, cx - sidePt/2, cy + sidePt/2, cx + sidePt/2];
    frame.fit(FitOptions.PROPORTIONALLY);
    return placed && placed[0];
  }
  // Dobra aspas para string VBScript
  function vbEsc(s){ return String(s||"").replace(/"/g, '""'); }

  // ---------- Execução do Zint via VBScript ----------
  function runZintViaVB(zintPath, dataText, outPath, scale, logPath) {
    // Monta a linha de comando com aspas corretas
    var cmd =
      '"' + zintPath + '"' +
      ' --barcode=datamatrix' +
      ' --data "' + dataText + '"' +
      ' --output "' + outPath + '"' +
      ' --scale ' + (scale||6) +
      ' --notext';

    // VBScript: executa, espera concluir, grava ExitCode no LOG
    var vb = ''
      + 'Dim sh: Set sh = CreateObject("WScript.Shell")' + "\r\n"
      + 'Dim rc: rc = sh.Run("' + vbEsc(cmd) + '", 0, True)' + "\r\n"
      + 'Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")' + "\r\n"
      + 'Dim f: Set f = fso.CreateTextFile("' + vbEsc(logPath) + '", True)' + "\r\n"
      + 'f.WriteLine "Cmd=' + vbEsc(cmd) + '"' + "\r\n"
      + 'f.WriteLine "ExitCode=" & rc' + "\r\n"
      + 'f.Close' + "\r\n";

    app.doScript(vb, ScriptLanguage.VISUAL_BASIC);
  }

  function gerarEColocarDM(numeroDoc, alvoPx, escala) {
    var zint = getZintPath();
    if (!zint) { alert("Zint não configurado. Rode o Set-ZintPath.jsx e selecione o zint.exe."); return; }

    var conteudo = "M" + String(numeroDoc);
    var links = ensureLinksFolder(doc);
    var useTemp = !links;

    // Saídas
    var outDir = useTemp ? Folder.temp : links;
    var outPath = outDir.fsName + "/" + conteudo.replace(/[^0-9A-Za-z_-]/g, "_") + ".png";
    var logPath = Folder.temp.fsName + "/dm_gen.log";

    // Executa
    runZintViaVB(zint, conteudo, outPath, (escala||6), logPath);

    // Aguarda arquivo (até 12s)
    if (!waitFile(outPath, 12000)) {
      var txt = "";
      var log = new File(logPath);
      if (log.exists) { log.open("r"); txt = log.read(); log.close(); }
      alert("Falha ao gerar o Data Matrix.\n\nSaída:\n" + txt);
      return;
    }

    // Coloca no layout
    var item = placeOnActivePage(doc, outPath, (alvoPx||200));
    if (!item) {
      alert("Gerou o PNG, mas não consegui colocar a imagem:\n" + outPath);
      return;
    }

    // Info final
    var msg = "Data Matrix gerado com sucesso!\n\n" +
              "Conteúdo: " + conteudo + "\n" +
              "Arquivo:  " + outPath + (useTemp ? "  (TEMP – mova para Links se quiser manter o vínculo)" : "");
    alert(msg);
  }

  function testarZint() {
    var zint = getZintPath();
    if (!zint) { alert("ZINT_EXE_PATH não definido."); return; }
    var outTxt = Folder.temp.fsName + "/zint_version.txt";
    var cmd = '"' + zint + '" --version > "' + outTxt + '" 2>&1';

    var vb = ''
      + 'Dim sh: Set sh = CreateObject("WScript.Shell")' + "\r\n"
      + 'sh.Run "' + vbEsc(cmd) + '", 0, True' + "\r\n";
    app.doScript(vb, ScriptLanguage.VISUAL_BASIC);

    var f = new File(outTxt), waited = 0;
    while (!f.exists && waited < 4000) { $.sleep(200); waited += 200; }
    var ver = "";
    if (f.exists) { f.open("r"); ver = f.read(); f.close(); f.remove(); }
    alert("Zint: " + zint + "\n\n" + (ver || "(sem retorno de versão)"));
  }

  // ---------- UI ----------
  var w = new Window("dialog", "Gerar Data Matrix (PNG)");
  w.orientation = "column"; w.alignChildren = "fill";

  var g1 = w.add("group"); g1.add("statictext", undefined, "Número do documento (só dígitos):");
  var edt = g1.add("edittext", undefined, ""); edt.characters = 20;

  var g2 = w.add("group"); g2.add("statictext", undefined, "Tamanho alvo (px):");
  var edtSize = g2.add("edittext", undefined, "30"); edtSize.characters = 6;

  var g3 = w.add("group"); g3.add("statictext", undefined, "Escala (px por módulo):");
  var edtScale = g3.add("edittext", undefined, "6"); edtScale.characters = 6;

  var gbtn = w.add("group");
  var btnTest = gbtn.add("button", undefined, "Testar Zint");
  var ok = gbtn.add("button", undefined, "Gerar", {name:"ok"});
  var cancel = gbtn.add("button", undefined, "Cancelar", {name:"cancel"});

  btnTest.onClick = function(){ testarZint(); };
  ok.onClick = function(){
    var n = (edt.text||"").replace(/\s+/g, "");
    if (!/^\d+$/.test(n)) { alert("Informe apenas dígitos do número do documento."); return; }
    var tpx = parseInt(edtSize.text||"200", 10) || 200;
    var sc  = parseInt(edtScale.text||"6", 10) || 6;
    w.close();
    gerarEColocarDM(n, tpx, sc);
  };

  w.center(); w.show();
})();
