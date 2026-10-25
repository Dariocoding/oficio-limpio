<?php
declare(strict_types=1);

/**
 * Oficio Limpio — sube un PDF escaneado y descarga Word editable
 * sin firmas, sellos ni marcas de CamScanner.
 */

$storageUploads = dirname(__DIR__) . '/storage/uploads';
$storageOutput  = dirname(__DIR__) . '/storage/output';
$worker         = dirname(__DIR__) . '/worker/build_fiel.py';
$stirlingUrl    = rtrim(getenv('STIRLING_URL') ?: 'https://pdf.klowshop.com', '/');

foreach ([$storageUploads, $storageOutput] as $dir) {
    if (!is_dir($dir)) {
        mkdir($dir, 0775, true);
    }
}

$error = null;
$downloadName = null;
$downloadPath = null;
$meta = null;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    try {
        if (!isset($_FILES['pdf']) || $_FILES['pdf']['error'] !== UPLOAD_ERR_OK) {
            throw new RuntimeException('Selecciona un PDF válido para subir.');
        }

        $tmp = $_FILES['pdf']['tmp_name'];
        $orig = $_FILES['pdf']['name'] ?? 'documento.pdf';
        $ext = strtolower(pathinfo($orig, PATHINFO_EXTENSION));
        if ($ext !== 'pdf') {
            throw new RuntimeException('Solo se aceptan archivos PDF.');
        }
        if (!is_uploaded_file($tmp)) {
            throw new RuntimeException('La subida falló.');
        }

        $id = bin2hex(random_bytes(8));
        $pdfPath = $storageUploads . '/' . $id . '.pdf';
        $docxPath = $storageOutput . '/' . $id . '.docx';

        if (!move_uploaded_file($tmp, $pdfPath)) {
            throw new RuntimeException('No se pudo guardar el PDF.');
        }

        $python = getenv('PYTHON_BIN') ?: 'python3';
        $cmd = [
            $python,
            $worker,
            '--pdf', $pdfPath,
            '--out', $docxPath,
            '--stirling', $stirlingUrl,
        ];

        $descriptors = [
            0 => ['pipe', 'r'],
            1 => ['pipe', 'w'],
            2 => ['pipe', 'w'],
        ];
        $proc = proc_open($cmd, $descriptors, $pipes, dirname(__DIR__), [
            'STIRLING_URL' => $stirlingUrl,
            'PYTHONUNBUFFERED' => '1',
        ]);
        if (!is_resource($proc)) {
            throw new RuntimeException('No se pudo iniciar el procesador.');
        }
        fclose($pipes[0]);
        $stdout = stream_get_contents($pipes[1]);
        $stderr = stream_get_contents($pipes[2]);
        fclose($pipes[1]);
        fclose($pipes[2]);
        $code = proc_close($proc);

        if ($code !== 0 || !is_file($docxPath)) {
            $detail = trim($stderr ?: $stdout) ?: 'error desconocido';
            throw new RuntimeException('No se pudo generar el Word: ' . $detail);
        }

        $metaJson = null;
        if (preg_match('/META_JSON:(.+)$/m', $stdout, $m)) {
            $metaJson = json_decode(trim($m[1]), true);
        }
        $meta = is_array($metaJson) ? $metaJson : ['ok' => true];

        $base = pathinfo($orig, PATHINFO_FILENAME);
        $downloadName = preg_replace('/[^A-Za-z0-9_\-]+/', '_', $base) . '-LIMPIO.docx';
        $downloadPath = $docxPath;

        // Descarga directa
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Disposition: attachment; filename="' . $downloadName . '"');
        header('Content-Length: ' . (string) filesize($docxPath));
        header('Cache-Control: no-store');
        readfile($docxPath);
        @unlink($pdfPath);
        // dejar docx unas horas por si quieren re-descargar no — borrar también
        @unlink($docxPath);
        exit;
    } catch (Throwable $e) {
        $error = $e->getMessage();
    }
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Oficio Limpio</title>
  <style>
    :root {
      --bg: #0f1419;
      --card: #1a222c;
      --line: #2c3a4a;
      --text: #e8eef4;
      --muted: #9aa8b6;
      --accent: #3d8bfd;
      --accent2: #1f6feb;
      --danger: #f85149;
      --ok: #3fb950;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: "Segoe UI", system-ui, sans-serif;
      color: var(--text);
      background:
        radial-gradient(1200px 600px at 10% -10%, #1b3a5f 0%, transparent 55%),
        radial-gradient(900px 500px at 100% 0%, #243018 0%, transparent 50%),
        var(--bg);
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 24px;
    }
    .card {
      width: 100%;
      max-width: 560px;
      background: color-mix(in srgb, var(--card) 92%, black);
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 28px 28px 24px;
      box-shadow: 0 20px 50px rgba(0,0,0,.35);
    }
    h1 {
      margin: 0 0 6px;
      font-size: 1.55rem;
      letter-spacing: -.02em;
    }
    .sub {
      margin: 0 0 22px;
      color: var(--muted);
      line-height: 1.45;
      font-size: .95rem;
    }
    .drop {
      border: 1.5px dashed var(--line);
      border-radius: 12px;
      padding: 28px 18px;
      text-align: center;
      background: rgba(255,255,255,.02);
      transition: border-color .15s, background .15s;
      cursor: pointer;
    }
    .drop:hover, .drop.drag {
      border-color: var(--accent);
      background: rgba(61,139,253,.06);
    }
    .drop strong { display: block; margin-bottom: 6px; }
    .drop span { color: var(--muted); font-size: .88rem; }
    input[type=file] { display: none; }
    .file-name {
      margin-top: 12px;
      font-size: .9rem;
      color: var(--ok);
      min-height: 1.2em;
    }
    button {
      margin-top: 18px;
      width: 100%;
      border: 0;
      border-radius: 10px;
      padding: 13px 16px;
      font-size: 1rem;
      font-weight: 600;
      color: #fff;
      background: linear-gradient(180deg, var(--accent), var(--accent2));
      cursor: pointer;
    }
    button:disabled {
      opacity: .55;
      cursor: not-allowed;
    }
    .err {
      margin-top: 14px;
      padding: 10px 12px;
      border-radius: 8px;
      background: rgba(248,81,73,.12);
      border: 1px solid rgba(248,81,73,.35);
      color: #ffb4b0;
      font-size: .9rem;
    }
    .hint {
      margin-top: 18px;
      padding-top: 14px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: .82rem;
      line-height: 1.45;
    }
    .hint code {
      color: #c9d4e0;
      background: rgba(255,255,255,.05);
      padding: 1px 5px;
      border-radius: 4px;
    }
    .busy { display: none; margin-top: 12px; color: var(--muted); font-size: .9rem; }
    .busy.show { display: block; }
  </style>
</head>
<body>
  <main class="card">
    <h1>Oficio Limpio</h1>
    <p class="sub">
      Sube el PDF escaneado (oficio + solicitud + relación). Te devolvemos un Word editable,
      <strong>sin firmas, sellos ni CamScanner</strong>, con el formato militar.
    </p>

    <?php if ($error): ?>
      <div class="err"><?= htmlspecialchars($error, ENT_QUOTES, 'UTF-8') ?></div>
    <?php endif; ?>

    <form method="post" enctype="multipart/form-data" id="form">
      <label class="drop" id="drop" for="pdf">
        <strong>Arrastra el PDF aquí o haz clic</strong>
        <span>Máx. ~25 MB · solo .pdf</span>
      </label>
      <input type="file" name="pdf" id="pdf" accept="application/pdf,.pdf" required>
      <div class="file-name" id="fname"></div>
      <button type="submit" id="btn">Generar Word limpio</button>
      <div class="busy" id="busy">Procesando con OCR + plantilla… puede tardar 30–90 s.</div>
    </form>

    <div class="hint">
      Usa Stirling (<code><?= htmlspecialchars($stirlingUrl, ENT_QUOTES, 'UTF-8') ?></code>) para OCR
      y reconstruye el Word con la misma plantilla que te gustó.
      Optimizado para oficios CEFOA / solicitud de munición.
    </div>
  </main>
  <script>
    const drop = document.getElementById('drop');
    const input = document.getElementById('pdf');
    const fname = document.getElementById('fname');
    const form = document.getElementById('form');
    const btn = document.getElementById('btn');
    const busy = document.getElementById('busy');

    input.addEventListener('change', () => {
      fname.textContent = input.files[0] ? input.files[0].name : '';
    });
    ['dragenter','dragover'].forEach(ev => drop.addEventListener(ev, e => {
      e.preventDefault(); drop.classList.add('drag');
    }));
    ['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => {
      e.preventDefault(); drop.classList.remove('drag');
    }));
    drop.addEventListener('drop', e => {
      if (e.dataTransfer.files.length) {
        input.files = e.dataTransfer.files;
        fname.textContent = input.files[0].name;
      }
    });
    form.addEventListener('submit', () => {
      btn.disabled = true;
      busy.classList.add('show');
    });
  </script>
</body>
</html>
