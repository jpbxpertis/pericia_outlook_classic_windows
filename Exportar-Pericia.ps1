<#
.SYNOPSIS
  Exporta los correos de una carpeta de Outlook para una pericia.

.DESCRIPTION
  Por cada correo de la carpeta (por defecto "PERICIA") genera una subcarpeta con:
    - mensaje.msg       copia original (MSG Unicode)
    - encabezados.txt   datos del mensaje, destinatarios (incluye CCO), datos MAPI
                        (Message-ID, In-Reply-To, References, horas UTC) y
                        encabezados de Internet cuando existen
    - mensaje.htm       cuerpo en HTML (tal como lo guarda Outlook)
    - mensaje.pdf       encabezados + cuerpo del correo
    - adjuntos\         adjuntos sueltos
  En la carpeta de la corrida: inventario.csv, hashes_sha256.csv y registro.log

  Los correos ENVIADOS por el usuario no tienen encabezados de Internet (esos
  encabezados solo existen en la copia del destinatario). Se marcan Tipo=ENVIADO.

  El PDF se genera en un proceso aparte con tiempo limite: si Word se cuelga,
  se corta, se marca el PDF como TIMEOUT y se sigue con el siguiente correo.
  Las imagenes remotas (http) se quitan de la copia usada para el PDF, para que
  Word no intente descargarlas.

  Se ejecuta con el usuario NORMAL (sin "Ejecutar como administrador").
  Outlook clasico debe estar abierto con ese mismo usuario.
  Salida: Documentos\Pericia_Export\AAAAMMDD_HHMMSS\

.EXAMPLE
  .\Exportar-Pericia.ps1
  .\Exportar-Pericia.ps1 -IncluirSubcarpetas
  .\Exportar-Pericia.ps1 -SinPDF
  .\Exportar-Pericia.ps1 -VerWord -TimeoutPDF 180
#>
[CmdletBinding()]
param(
    [string]$NombreCarpeta = "PERICIA",
    [string]$Destino = "",
    [switch]$IncluirSubcarpetas,
    [switch]$SinPDF,
    [switch]$VerWord,
    [int]$TimeoutPDF = 120
)

$ErrorActionPreference = "Stop"

# ---------------- Constantes ----------------
$olMSGUnicode           = 9
$olHTML                 = 5
$olExchangePublicFolder = 2
$MSGFLAG_FROMME         = 0x20

$PR_TRANSPORT_HEADERS    = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
$PR_INTERNET_MESSAGE_ID  = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$PR_IN_REPLY_TO_ID       = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
$PR_INTERNET_REFERENCES  = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"
$PR_CLIENT_SUBMIT_TIME   = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
$PR_MESSAGE_DELIVERY_TIME= "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
$PR_MESSAGE_FLAGS        = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
$PR_SMTP_ADDRESS         = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

if ([string]::IsNullOrWhiteSpace($Destino)) {
    $Destino = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "Pericia_Export"
}

$script:LogFile = $null
$Outlook = $null
$Namespace = $null

# Script que corre en un proceso aparte para convertir HTML -> PDF con Word.
# Recibe las rutas por variables de entorno (evita problemas con comillas en los nombres).
$ScriptConversor = @'
$ErrorActionPreference = 'Stop'
$codigo = 1
$w = $null; $d = $null
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = ($env:PERICIA_VISIBLE -eq '1')
    $w.DisplayAlerts = 0
    $d = $w.Documents.Open($env:PERICIA_SRC, $false, $true, $false)
    $d.ExportAsFixedFormat($env:PERICIA_DST, 17)
    $codigo = 0
}
catch {
    try { Set-Content -LiteralPath $env:PERICIA_ERR -Value $_.Exception.Message -Encoding UTF8 } catch {}
}
finally {
    if ($d) { try { $d.Close(0) } catch {} }
    if ($w) { try { $w.Quit() } catch {} }
}
exit $codigo
'@
$ConversorCodificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptConversor))

# ---------------- Funciones auxiliares ----------------
function Write-Log {
    param([string]$Mensaje, [string]$Nivel = "INFO")
    if (-not $script:LogFile) { return }
    $linea = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Nivel, $Mensaje
    Add-Content -LiteralPath $script:LogFile -Value $linea -Encoding UTF8
}

function Release-Com {
    param($Obj)
    if ($null -ne $Obj) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Get-Prop {
    param($Obj, [string]$Nombre)
    try {
        $v = $Obj.$Nombre
        if ($null -eq $v) { return "" }
        return [string]$v
    } catch { return "" }
}

function Get-Mapi {
    param($Obj, [string]$Tag)
    try { return $Obj.PropertyAccessor.GetProperty($Tag) } catch { return $null }
}

function Format-Utc {
    param($Valor)
    try {
        if ($null -eq $Valor) { return "" }
        $dt = [datetime]$Valor
        if ($dt.Year -ge 4500) { return "" }
        return $dt.ToString("yyyy-MM-dd HH:mm:ss") + " UTC"
    } catch { return "" }
}

function Get-NombreSeguro {
    param([string]$Texto, [int]$Max = 40, [string]$PorDefecto = "sin_asunto")
    if ([string]::IsNullOrWhiteSpace($Texto)) { return $PorDefecto }
    $invalidos = [IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Texto.ToCharArray()) {
        if ($invalidos -contains $c) { [void]$sb.Append('_') } else { [void]$sb.Append($c) }
    }
    $t = ($sb.ToString() -replace '\s+', ' ').Trim(' ', '.')
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max).Trim(' ', '.') }
    if ($t -eq '') { $t = $PorDefecto }
    return $t
}

function Get-FileHashOrEmpty {
    param([string]$Path)
    if ($Path -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
        try { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash } catch { return "" }
    }
    return ""
}

function Get-SenderSmtp {
    param($Item)
    try {
        if ((Get-Prop $Item "SenderEmailType") -eq "EX") {
            $eu = $Item.Sender.GetExchangeUser()
            if ($eu) { return [string]$eu.PrimarySmtpAddress }
        }
        return (Get-Prop $Item "SenderEmailAddress")
    } catch { return "" }
}

function Find-OutlookFolders {
    param($Folders, [string]$Nombre, [System.Collections.ArrayList]$Encontradas)
    foreach ($f in $Folders) {
        try {
            if ($f.Name -ieq $Nombre) { [void]$Encontradas.Add($f) }
            if ($f.Folders.Count -gt 0) {
                Find-OutlookFolders -Folders $f.Folders -Nombre $Nombre -Encontradas $Encontradas
            }
        } catch {}
    }
}

function Add-Subcarpetas {
    param($Folder, [System.Collections.ArrayList]$Lista)
    foreach ($sub in $Folder.Folders) {
        [void]$Lista.Add($sub)
        Add-Subcarpetas -Folder $sub -Lista $Lista
    }
}

# Crea una copia del HTML para el PDF: sin recursos remotos y con los encabezados al principio.
function New-HtmlParaPdf {
    param([string]$HtmlOrigen, [string]$HtmlDestino, [string]$TextoEncabezados)

    $bytes = [IO.File]::ReadAllBytes($HtmlOrigen)
    $muestra = [Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min(4096, $bytes.Length))
    $enc = [Text.Encoding]::Default
    $m = [regex]::Match($muestra, 'charset\s*=\s*["'']?([A-Za-z0-9_\-]+)', 'IgnoreCase')
    if ($m.Success) { try { $enc = [Text.Encoding]::GetEncoding($m.Groups[1].Value) } catch {} }

    $ms = New-Object IO.MemoryStream(,$bytes)
    $sr = New-Object IO.StreamReader($ms, $enc, $true)
    $html = $sr.ReadToEnd()
    $encReal = $sr.CurrentEncoding
    $sr.Close()

    # Quitar imagenes / fondos / hojas de estilo remotas (Word intentaria descargarlas)
    $html = [regex]::Replace($html, '(?i)\b(src|background)\s*=\s*"(https?:)?//[^"]*"', '$1=""')
    $html = [regex]::Replace($html, "(?i)\b(src|background)\s*=\s*'(https?:)?//[^']*'", '$1=""')
    $html = [regex]::Replace($html, '(?i)\b(src|background)\s*=\s*(https?:)?//[^\s>]+', '$1=""')
    $html = [regex]::Replace($html, '(?i)url\(\s*["'']?(https?:)?//[^)]*\)', 'none')
    $html = [regex]::Replace($html, '(?i)<link\b[^>]*\bhref\s*=\s*["'']?(https?:)?//[^>]*>', '')

    # Bloque de encabezados + salto de pagina
    $bloque = '<pre style="font-family:Consolas,monospace;font-size:8.0pt;white-space:pre-wrap">' +
              [System.Net.WebUtility]::HtmlEncode($TextoEncabezados) +
              '</pre>' + "<br clear=all style='page-break-before:always'>"

    $mb = [regex]::Match($html, '(?i)<body[^>]*>')
    if ($mb.Success) { $html = $html.Insert($mb.Index + $mb.Length, $bloque) }
    else { $html = $bloque + $html }

    $sw = New-Object IO.StreamWriter($HtmlDestino, $false, $encReal)
    $sw.Write($html)
    $sw.Close()
}

# Convierte HTML a PDF en un proceso aparte, con tiempo limite.
function Convert-HtmlAPdf {
    param([string]$Origen, [string]$PdfDestino, [int]$Timeout)

    $errFile = Join-Path $env:TEMP ("pericia_pdf_err_{0}.txt" -f [guid]::NewGuid())
    $wordAntes = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

    $env:PERICIA_SRC = $Origen
    $env:PERICIA_DST = $PdfDestino
    $env:PERICIA_ERR = $errFile
    $env:PERICIA_VISIBLE = $(if ($VerWord) { '1' } else { '0' })

    $p = Start-Process -FilePath "powershell.exe" `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $script:ConversorCodificado) `
        -WindowStyle Hidden -PassThru
    $null = $p.Handle   # necesario para poder leer ExitCode despues

    $resultado = @{ Estado = "ERROR"; Detalle = "" }

    if (-not $p.WaitForExit($Timeout * 1000)) {
        try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch {}
        Get-Process WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $wordAntes -notcontains $_.Id } |
            ForEach-Object { try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {} }
        $resultado.Estado = "TIMEOUT"
        $resultado.Detalle = "Word no respondio en $Timeout segundos"
    }
    elseif ($p.ExitCode -eq 0 -and (Test-Path -LiteralPath $PdfDestino)) {
        $resultado.Estado = "OK"
    }
    else {
        $resultado.Detalle = "Codigo de salida $($p.ExitCode)"
        if (Test-Path -LiteralPath $errFile) {
            $resultado.Detalle = (Get-Content -LiteralPath $errFile -Raw).Trim()
        }
    }

    Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    return $resultado
}

# ---------------- Proceso principal ----------------
try {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "PowerShell esta corriendo ELEVADO. Si Outlook esta abierto sin elevar, la conexion falla (error 0x80080005). Conviene ejecutarlo normal."
    }

    $RunFolder = Join-Path $Destino (Get-Date -Format "yyyyMMdd_HHmmss")
    New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null
    $script:LogFile = Join-Path $RunFolder "registro.log"

    Write-Log "Inicio. Usuario: $env:USERDOMAIN\$env:USERNAME - Equipo: $env:COMPUTERNAME"
    Write-Log "Zona horaria del equipo: $([TimeZoneInfo]::Local.DisplayName)"
    Write-Log "Carpeta buscada: $NombreCarpeta - Destino: $RunFolder"

    # 1) Conexion con Outlook
    try {
        $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        try {
            $Outlook = New-Object -ComObject Outlook.Application
        } catch {
            throw "No se pudo conectar con Outlook: $($_.Exception.Message). Verificar que sea Outlook clasico (el 'nuevo Outlook' no sirve) y que PowerShell NO este como administrador."
        }
    }
    $Namespace = $Outlook.GetNamespace("MAPI")

    # 2) Buscar la carpeta
    $encontradas = New-Object System.Collections.ArrayList
    foreach ($store in $Namespace.Stores) {
        try {
            if ($store.ExchangeStoreType -eq $olExchangePublicFolder) { continue }
            $root = $store.GetRootFolder()
            Find-OutlookFolders -Folders $root.Folders -Nombre $NombreCarpeta -Encontradas $encontradas
        } catch {
            Write-Log "No se pudo recorrer el almacen '$($store.DisplayName)': $($_.Exception.Message)" "WARN"
        }
    }

    if ($encontradas.Count -eq 0) {
        throw "No se encontro ninguna carpeta llamada '$NombreCarpeta' en Outlook."
    }
    if ($encontradas.Count -gt 1) {
        Write-Warning "Hay $($encontradas.Count) carpetas llamadas '$NombreCarpeta'. Se usa la primera:"
        foreach ($f in $encontradas) { Write-Host "   $($f.FolderPath)" }
    }
    $Raiz = $encontradas[0]
    Write-Log "Carpeta usada: $($Raiz.FolderPath)"
    Write-Host "Carpeta: $($Raiz.FolderPath)" -ForegroundColor Cyan

    $Carpetas = New-Object System.Collections.ArrayList
    [void]$Carpetas.Add($Raiz)
    if ($IncluirSubcarpetas) { Add-Subcarpetas -Folder $Raiz -Lista $Carpetas }

    # Direcciones propias del usuario (para identificar correos enviados)
    $misDirecciones = New-Object System.Collections.Generic.List[string]
    try { foreach ($acc in $Namespace.Accounts) { $s = Get-Prop $acc "SmtpAddress"; if ($s) { $misDirecciones.Add($s.ToLower()) } } } catch {}
    try {
        $eu = $Namespace.CurrentUser.AddressEntry.GetExchangeUser()
        if ($eu) { $misDirecciones.Add(([string]$eu.PrimarySmtpAddress).ToLower()) }
    } catch {}
    try { $sd = [string]$Raiz.Store.DisplayName; if ($sd -match '@') { $misDirecciones.Add($sd.ToLower()) } } catch {}
    Write-Log "Direcciones propias: $(($misDirecciones | Select-Object -Unique) -join ', ')"

    $Resultados = New-Object System.Collections.Generic.List[object]
    $n = 0

    foreach ($Carpeta in $Carpetas) {
        $Items = $Carpeta.Items
        try { $Items.Sort("[ReceivedTime]", $false) } catch {}
        $total = $Items.Count
        Write-Host "$($Carpeta.FolderPath): $total elementos" -ForegroundColor Cyan
        Write-Log "Procesando $($Carpeta.FolderPath): $total elementos"

        for ($k = 1; $k -le $total; $k++) {
            $n++
            $Item = $null
            $subject = ""; $sender = ""; $receivedText = ""; $sentText = ""; $tipo = ""; $messageId = ""
            $msgStatus = "PENDIENTE"; $headersStatus = "PENDIENTE"; $htmlStatus = "PENDIENTE"; $pdfStatus = "PENDIENTE"
            $adjStatus = "0"
            $msgPath = ""; $headersPath = ""; $htmlPath = ""; $pdfPath = ""
            $errores = New-Object System.Collections.Generic.List[string]

            try {
                $Item = $Items.Item($k)

                $subject    = Get-Prop $Item "Subject"
                $senderName = Get-Prop $Item "SenderName"
                $senderSmtp = Get-SenderSmtp $Item
                if ($senderSmtp) { $sender = "$senderName <$senderSmtp>" } else { $sender = $senderName }

                $received = $null
                try { $received = [datetime]$Item.ReceivedTime } catch {}
                if ($received) { $receivedText = $received.ToString("yyyy-MM-dd HH:mm:ss") }
                try { $sentText = ([datetime]$Item.SentOn).ToString("yyyy-MM-dd HH:mm:ss") } catch {}

                if ($received) { $fechaNombre = $received.ToString("yyyyMMdd_HHmm") } else { $fechaNombre = "sinfecha" }
                $dirMail = Join-Path $RunFolder ("{0:D4}_{1}_{2}" -f $n, $fechaNombre, (Get-NombreSeguro $subject 40))
                New-Item -ItemType Directory -Path $dirMail -Force | Out-Null

                Write-Host ("[{0}] {1} | {2}" -f $n, $receivedText, $subject)

                # --- MSG original
                $msgPath = Join-Path $dirMail "mensaje.msg"
                try {
                    $Item.SaveAs($msgPath, $olMSGUnicode)
                    $msgStatus = "OK"
                } catch {
                    $msgStatus = "ERROR"
                    $errores.Add("MSG: $($_.Exception.Message)")
                }

                # --- Adjuntos
                $adjNombres = New-Object System.Collections.Generic.List[string]
                $cantAdj = 0
                try { $cantAdj = $Item.Attachments.Count } catch {}
                if ($cantAdj -gt 0) {
                    $dirAdj = Join-Path $dirMail "adjuntos"
                    New-Item -ItemType Directory -Path $dirAdj -Force | Out-Null
                    $okAdj = 0
                    for ($a = 1; $a -le $cantAdj; $a++) {
                        $att = $null
                        try {
                            $att = $Item.Attachments.Item($a)
                            $original = Get-Prop $att "FileName"
                            if (-not $original) { $original = Get-Prop $att "DisplayName" }
                            $adjNombres.Add($original)
                            $fn   = Get-NombreSeguro $original 200 "adjunto"
                            $ext  = [IO.Path]::GetExtension($fn)
                            $base = [IO.Path]::GetFileNameWithoutExtension($fn)
                            if ($base.Length -gt 60) { $base = $base.Substring(0, 60) }
                            $att.SaveAsFile((Join-Path $dirAdj ("{0:D2}_{1}{2}" -f $a, $base, $ext)))
                            $okAdj++
                        } catch {
                            $errores.Add("Adjunto ${a}: $($_.Exception.Message)")
                        } finally {
                            Release-Com $att
                        }
                    }
                    $adjStatus = "$okAdj/$cantAdj"
                }

                # --- Destinatarios (Para / CC / CCO) con SMTP
                $destinatarios = New-Object System.Collections.Generic.List[string]
                try {
                    $cantRec = $Item.Recipients.Count
                    for ($r = 1; $r -le $cantRec; $r++) {
                        $rc = $null
                        try {
                            $rc = $Item.Recipients.Item($r)
                            $tipoRec = switch ($rc.Type) { 1 { "Para" } 2 { "CC" } 3 { "CCO" } default { "Otro" } }
                            $smtp = ""
                            try { $smtp = [string]$rc.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS) } catch {}
                            if (-not $smtp) { $smtp = Get-Prop $rc "Address" }
                            $destinatarios.Add(("{0,-5} {1} <{2}>" -f $tipoRec, (Get-Prop $rc "Name"), $smtp))
                        } catch {} finally { Release-Com $rc }
                    }
                } catch {}

                # --- Datos MAPI
                $internetHeaders = [string](Get-Mapi $Item $PR_TRANSPORT_HEADERS)
                $messageId       = [string](Get-Mapi $Item $PR_INTERNET_MESSAGE_ID)
                $inReplyTo       = [string](Get-Mapi $Item $PR_IN_REPLY_TO_ID)
                $references      = [string](Get-Mapi $Item $PR_INTERNET_REFERENCES)
                $submitUtc       = Format-Utc (Get-Mapi $Item $PR_CLIENT_SUBMIT_TIME)
                $deliveryUtc     = Format-Utc (Get-Mapi $Item $PR_MESSAGE_DELIVERY_TIME)
                $flags = 0
                try { $flags = [int](Get-Mapi $Item $PR_MESSAGE_FLAGS) } catch {}

                $tieneHeaders = -not [string]::IsNullOrWhiteSpace($internetHeaders)
                $esPropio = (($flags -band $MSGFLAG_FROMME) -ne 0) -or ($senderSmtp -and ($misDirecciones -contains $senderSmtp.ToLower()))

                if ($tieneHeaders) { $tipo = "RECIBIDO"; $headersStatus = "OK" }
                elseif ($esPropio) { $tipo = "ENVIADO";  $headersStatus = "NO_APLICA_ENVIADO" }
                else               { $tipo = "RECIBIDO"; $headersStatus = "NO_DISPONIBLES" }

                $b = New-Object System.Text.StringBuilder
                [void]$b.AppendLine("===== DATOS DEL MENSAJE =====")
                [void]$b.AppendLine("Tipo:            $tipo")
                [void]$b.AppendLine("Asunto:          $subject")
                [void]$b.AppendLine("De:              $sender")
                [void]$b.AppendLine("Enviado (local): $sentText")
                [void]$b.AppendLine("Recibido (local):$receivedText")
                [void]$b.AppendLine("Envio (UTC):     $submitUtc")
                [void]$b.AppendLine("Entrega (UTC):   $deliveryUtc")
                [void]$b.AppendLine("Zona horaria:    $([TimeZoneInfo]::Local.DisplayName)")
                [void]$b.AppendLine("Message-ID:      $messageId")
                [void]$b.AppendLine("In-Reply-To:     $inReplyTo")
                [void]$b.AppendLine("References:      $references")
                [void]$b.AppendLine("Conversacion:    $(Get-Prop $Item 'ConversationTopic') [$(Get-Prop $Item 'ConversationID')]")
                [void]$b.AppendLine("Carpeta Outlook: $($Carpeta.FolderPath)")
                [void]$b.AppendLine("Adjuntos:        $($adjNombres -join '; ')")
                [void]$b.AppendLine("")
                [void]$b.AppendLine("===== DESTINATARIOS =====")
                foreach ($d in $destinatarios) { [void]$b.AppendLine($d) }
                [void]$b.AppendLine("")
                [void]$b.AppendLine("===== ENCABEZADOS DE INTERNET =====")
                if ($tieneHeaders) {
                    [void]$b.AppendLine($internetHeaders.TrimEnd())
                } elseif ($tipo -eq "ENVIADO") {
                    [void]$b.AppendLine("(Mensaje ENVIADO por el titular del buzon: los encabezados de transporte")
                    [void]$b.AppendLine(" solo existen en la copia recibida por los destinatarios.)")
                } else {
                    [void]$b.AppendLine("(Sin encabezados de Internet: mensaje interno de Exchange o encabezados no conservados.)")
                }
                [void]$b.AppendLine("===== FIN DE ENCABEZADOS =====")
                $textoEncabezados = $b.ToString()

                $headersPath = Join-Path $dirMail "encabezados.txt"
                try {
                    [IO.File]::WriteAllText($headersPath, $textoEncabezados, (New-Object System.Text.UTF8Encoding($true)))
                } catch {
                    $headersStatus = "ERROR"
                    $errores.Add("Encabezados: $($_.Exception.Message)")
                }

                # --- HTML
                $htmlPath = Join-Path $dirMail "mensaje.htm"
                try {
                    $Item.SaveAs($htmlPath, $olHTML)
                    $htmlStatus = "OK"
                } catch {
                    $htmlStatus = "ERROR"
                    $errores.Add("HTML: $($_.Exception.Message)")
                }

                # --- PDF
                if ($SinPDF) {
                    $pdfStatus = "OMITIDO"
                } elseif ($htmlStatus -ne "OK") {
                    $pdfStatus = "HTML_FALLO"
                } else {
                    $pdfPath = Join-Path $dirMail "mensaje.pdf"
                    $htmlPdf = Join-Path $dirMail "_pdf_fuente.htm"
                    try {
                        New-HtmlParaPdf -HtmlOrigen $htmlPath -HtmlDestino $htmlPdf -TextoEncabezados $textoEncabezados
                        $res = Convert-HtmlAPdf -Origen $htmlPdf -PdfDestino $pdfPath -Timeout $TimeoutPDF
                        $pdfStatus = $res.Estado
                        if ($res.Detalle) { $errores.Add("PDF: $($res.Detalle)") }
                    } catch {
                        $pdfStatus = "ERROR"
                        $errores.Add("PDF: $($_.Exception.Message)")
                    } finally {
                        Remove-Item -LiteralPath $htmlPdf -Force -ErrorAction SilentlyContinue
                    }
                    if ($pdfStatus -ne "OK") { $pdfPath = "" }
                }

                foreach ($e in $errores) { Write-Host "  $e" -ForegroundColor Yellow }

                if (($msgStatus -ne "OK") -or ($htmlStatus -ne "OK") -or ($pdfStatus -notin @("OK", "OMITIDO")) -or ($headersStatus -eq "ERROR")) {
                    $overall = "REVISAR"
                } elseif ($headersStatus -eq "NO_DISPONIBLES") {
                    $overall = "REVISAR_ENCABEZADOS"
                } else {
                    $overall = "OK"
                }

                $Resultados.Add([pscustomobject]@{
                    Numero             = $n
                    Tipo               = $tipo
                    Carpeta            = $Carpeta.FolderPath
                    Asunto             = $subject
                    Remitente          = $sender
                    Enviado            = $sentText
                    Recibido           = $receivedText
                    MessageID          = $messageId
                    Estado             = $overall
                    EstadoMSG          = $msgStatus
                    EstadoEncabezados  = $headersStatus
                    EstadoHTML         = $htmlStatus
                    EstadoPDF          = $pdfStatus
                    Adjuntos           = $adjStatus
                    MSG                = $msgPath
                    EncabezadosTXT     = $headersPath
                    HTML               = $htmlPath
                    PDF                = $pdfPath
                    SHA256_MSG         = (Get-FileHashOrEmpty $msgPath)
                    SHA256_Encabezados = (Get-FileHashOrEmpty $headersPath)
                    SHA256_HTML        = (Get-FileHashOrEmpty $htmlPath)
                    SHA256_PDF         = (Get-FileHashOrEmpty $pdfPath)
                    Error              = ($errores -join " | ")
                })

                Write-Log "Correo ${n} ($tipo): MSG=$msgStatus; Headers=$headersStatus; HTML=$htmlStatus; PDF=$pdfStatus; Adjuntos=$adjStatus"
            }
            catch {
                $message = $_.Exception.Message
                Write-Host "  ERROR GENERAL: $message" -ForegroundColor Red
                Write-Log "Correo ${n}: error general: $message" "ERROR"
                $Resultados.Add([pscustomobject]@{
                    Numero = $n; Tipo = $tipo; Carpeta = $Carpeta.FolderPath; Asunto = $subject; Remitente = $sender
                    Enviado = $sentText; Recibido = $receivedText; MessageID = $messageId
                    Estado = "ERROR_GENERAL"; EstadoMSG = $msgStatus; EstadoEncabezados = $headersStatus
                    EstadoHTML = $htmlStatus; EstadoPDF = $pdfStatus; Adjuntos = $adjStatus
                    MSG = $msgPath; EncabezadosTXT = $headersPath; HTML = $htmlPath; PDF = $pdfPath
                    SHA256_MSG = (Get-FileHashOrEmpty $msgPath)
                    SHA256_Encabezados = (Get-FileHashOrEmpty $headersPath)
                    SHA256_HTML = (Get-FileHashOrEmpty $htmlPath)
                    SHA256_PDF = (Get-FileHashOrEmpty $pdfPath)
                    Error = $message
                })
            }
            finally {
                Release-Com $Item
                $Item = $null
            }
        }
        Release-Com $Items
    }

    # Inventario
    $CsvPath = Join-Path $RunFolder "inventario.csv"
    $Resultados | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log "Inventario generado: $CsvPath"

    $Ok = @($Resultados | Where-Object { $_.Estado -eq "OK" }).Count
    $Revisar = $Resultados.Count - $Ok
    $Enviados = @($Resultados | Where-Object { $_.Tipo -eq "ENVIADO" }).Count
    $Pdfs = @($Resultados | Where-Object { $_.EstadoPDF -eq "OK" }).Count
    Write-Log "Proceso finalizado. Elementos: $($Resultados.Count). Completos: $Ok. Para revisar: $Revisar. Enviados: $Enviados. PDF OK: $Pdfs."

    # Hash SHA-256 de todos los archivos generados (va al final; despues no se escribe nada mas)
    $HashCsv = Join-Path $RunFolder "hashes_sha256.csv"
    Get-ChildItem -LiteralPath $RunFolder -Recurse -File |
        Where-Object { $_.FullName -ne $HashCsv } |
        ForEach-Object {
            [pscustomobject]@{
                Archivo = $_.FullName.Substring($RunFolder.Length + 1)
                Bytes   = $_.Length
                SHA256  = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            }
        } | Export-Csv -LiteralPath $HashCsv -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "PROCESO FINALIZADO" -ForegroundColor Green
    Write-Host "Elementos revisados: $($Resultados.Count)  (enviados: $Enviados)"
    Write-Host "Completos:           $Ok"
    Write-Host "Para revisar:        $Revisar"
    Write-Host "PDF generados:       $Pdfs"
    Write-Host "Carpeta de salida:   $RunFolder" -ForegroundColor Cyan
    Write-Host "Inventario:          $CsvPath"
    Write-Host "Hashes:              $HashCsv"
    Write-Host "Registro:            $script:LogFile"
    Write-Host "============================================" -ForegroundColor Cyan
}
catch {
    Write-Host ""
    Write-Host "ERROR FATAL: $($_.Exception.Message)" -ForegroundColor Red
    try { Write-Log "ERROR FATAL: $($_.Exception.Message)" "ERROR" } catch {}
}
finally {
    Release-Com $Namespace
    Release-Com $Outlook
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
