# Exportar correos de Outlook para una pericia

Script de PowerShell que toma todos los correos de una carpeta de **Outlook clásico** (por defecto, una carpeta llamada `PERICIA`) y genera, por cada uno, la copia original, los encabezados, el cuerpo en PDF y los adjuntos. Además deja un inventario y los hashes SHA-256 de todo lo generado, pensado para presentar en una pericia informática o para la cadena de custodia.

No hace falta saber programar ni ser administrador: se copia en Documentos, se hace doble clic y listo.

---

## Qué genera

Cada ejecución crea una carpeta nueva con fecha y hora en:

```
Documentos\Pericia_Export\AAAAMMDD_HHMMSS\
```

Adentro hay **una subcarpeta por correo** (`0001_20220215_0954_Asunto del correo`, etc.) con:

| Archivo | Contenido |
|---|---|
| `mensaje.msg` | Copia original del correo (formato MSG de Outlook). |
| `encabezados.txt` | Datos del mensaje: remitente, destinatarios (Para, CC y **CCO**), fechas locales y UTC, Message-ID, In-Reply-To, References, conversación y, si existen, los **encabezados de Internet completos** (`Received:`, `DKIM`, etc.). |
| `mensaje.htm` | Cuerpo del correo en HTML, tal como lo guarda Outlook. |
| `mensaje.pdf` | Primera página con los encabezados y, a continuación, el cuerpo del correo. |
| `adjuntos\` | Los archivos adjuntos sueltos. |

Y en la raíz de la corrida:

| Archivo | Contenido |
|---|---|
| `inventario.csv` | Una fila por correo: tipo (enviado/recibido), asunto, remitente, fechas, Message-ID, estado de cada archivo y su SHA-256. |
| `hashes_sha256.csv` | Hash SHA-256 y tamaño de **todos** los archivos generados. |
| `registro.log` | Registro de la ejecución: usuario, equipo, zona horaria, carpeta usada y resultado de cada correo. |

Los CSV usan `;` como separador, así que abren directo en Excel con configuración regional en español.

---

## Requisitos

- Windows 10 u 11.
- **Outlook clásico** (Microsoft 365, 2016, 2019 o 2021) **abierto** con la cuenta que tiene la carpeta.
  El *nuevo Outlook* (el que se instala como app y se parece a Outlook web) **no sirve**, porque no permite automatización.
- **Word** instalado, solo para generar los PDF. Sin Word, se genera todo lo demás.
- No requiere permisos de administrador ni instalar nada.

---

## Uso paso a paso

1. **Preparar la carpeta en Outlook.** Crear una carpeta llamada `PERICIA` (en cualquier lugar del buzón; por ejemplo, dentro de la Bandeja de entrada) y **mover** ahí todos los correos que se van a presentar, tanto recibidos como enviados.
2. **Descargar** `Exportar-Pericia.ps1` y `Exportar-Pericia.bat` (botón verde **Code → Download ZIP**, y descomprimir).
3. **Copiar los dos archivos a la carpeta Documentos** del usuario.
4. Con Outlook abierto, **hacer doble clic en `Exportar-Pericia.bat`**.
   No usar "Ejecutar como administrador".
5. Esperar a que aparezca **PROCESO FINALIZADO**. La ventana muestra cuántos correos se exportaron y la carpeta de salida.

> Si Windows muestra "Windows protegió su PC", hacer clic en **Más información → Ejecutar de todas formas**. Si Outlook avisa que "un programa está intentando acceder a la información", elegir **Permitir**.

### Opciones

Se agregan después del nombre del `.bat`, desde una ventana de PowerShell o CMD abierta en Documentos:

| Opción | Para qué sirve |
|---|---|
| `-NombreCarpeta "OTRA"` | Usar otra carpeta de Outlook en vez de `PERICIA`. |
| `-IncluirSubcarpetas` | Incluir también las subcarpetas de `PERICIA`. |
| `-SinPDF` | No generar PDF (útil si Word da problemas; el resto se genera igual). |
| `-VerWord` | Mostrar Word mientras genera los PDF, para ver si se frena en algún cartel. |
| `-TimeoutPDF 180` | Segundos máximos por PDF antes de abandonarlo (por defecto, 120). |
| `-Destino "D:\Salida"` | Guardar en otra carpeta en vez de `Documentos\Pericia_Export`. |

Ejemplo:

```bat
Exportar-Pericia.bat -IncluirSubcarpetas
```

---

## Cómo interpretar los resultados

La columna **Estado** del inventario:

| Estado | Significado |
|---|---|
| `OK` | Todo se generó correctamente. |
| `REVISAR` | Falló algún archivo (MSG, HTML o PDF). El detalle está en la columna `Error`. |
| `REVISAR_ENCABEZADOS` | Correo recibido que no tiene encabezados de Internet guardados (ver abajo). |
| `ERROR_GENERAL` | No se pudo procesar el correo. Ver `Error` y `registro.log`. |

La columna **Tipo** indica si el correo es `ENVIADO` (por el titular del buzón) o `RECIBIDO`.

### ¿Por qué los correos enviados no tienen encabezados de Internet?

Los encabezados de transporte (`Received:`, `DKIM-Signature`, `Authentication-Results`, etc.) los van agregando los servidores **mientras el correo viaja**, y quedan guardados solo en la **copia que recibe el destinatario**. La copia en *Elementos enviados* nunca los tiene; no es un error ni una pérdida.

Para los enviados, el script igualmente registra lo que sí está disponible: hora de envío (local y UTC), Message-ID, In-Reply-To/References (que vinculan el correo con su hilo) y la lista completa de destinatarios, **incluida la CCO**, que solo figura en la copia del remitente.

Si se necesitan los encabezados de un correo enviado, hay que obtener la copia del destinatario (por ejemplo, la respuesta que contiene el hilo, o la copia en el buzón de la otra parte).

Un correo **recibido** sin encabezados suele ser un mensaje interno entre usuarios del mismo servidor Exchange/Microsoft 365, o un correo importado desde otro sistema o PST que no los conservó.

---

## Si algo falla

| Síntoma | Causa probable y solución |
|---|---|
| `No se pudo conectar con Outlook` | Outlook está cerrado, es el *nuevo Outlook*, o PowerShell se abrió como administrador. Abrir Outlook clásico y ejecutar el `.bat` normal. |
| `No se encontro ninguna carpeta llamada 'PERICIA'` | Revisar el nombre exacto de la carpeta o usar `-NombreCarpeta`. |
| Avisa que hay varias carpetas `PERICIA` | Usa la primera que encuentra y muestra la ruta de todas. Renombrar las que sobren o usar otro nombre. |
| PDF en `TIMEOUT` | Word quedó esperando un cartel oculto (activación, cuenta, privacidad, modo seguro). Abrir Word a mano una vez, cerrar el cartel, y volver a ejecutar. O usar `-VerWord` para verlo. |
| Se queda mucho tiempo en un correo | Buscar en la barra de tareas un aviso de Outlook ("un programa está intentando acceder…") y aceptarlo. |
| El `.bat` se cierra enseguida | Abrir PowerShell en Documentos y ejecutar `.\Exportar-Pericia.bat` para ver el mensaje. |

---

## Para entregar

Se recomienda comprimir la carpeta de la corrida y registrar el hash del ZIP en el acta:

```powershell
Compress-Archive "$env:USERPROFILE\Documents\Pericia_Export\AAAAMMDD_HHMMSS" "$env:USERPROFILE\Documents\Pericia_Export\Pericia.zip"
Get-FileHash "$env:USERPROFILE\Documents\Pericia_Export\Pericia.zip" -Algorithm SHA256
```

Borrar antes las corridas de prueba o incompletas, para que en `Pericia_Export` quede solo la definitiva.

---

## Qué hace y qué no hace

- **Solo lee.** No modifica, mueve ni borra correos en Outlook. El `.htm` usado para el PDF es una copia temporal que se elimina.
- Las **imágenes remotas** (enlazadas desde Internet) se quitan solo de la copia usada para el PDF, para evitar que Word quede esperando descargas. El `.msg` y el `.htm` conservan el contenido original.
- El `.bat` usa `-ExecutionPolicy Bypass` **solo para ese proceso**; no cambia la configuración del equipo.
- Cierra únicamente la instancia de Word que abre; Outlook queda abierto.
- Si hay varios usuarios involucrados, cada uno debe ejecutarlo **en su propia sesión**, con su Outlook.

> Esta herramienta facilita la extracción y el inventario de los correos, pero no reemplaza el criterio del perito. La validez de la prueba depende del procedimiento completo (acta, testigos, cadena de custodia, etc.).

---

## Licencia

[MIT](LICENSE). Uso libre, sin garantías.

---

Desarrollado por **[Xpertis](https://www.xpertis.com.ar)** · [www.xpertis.com.ar](https://www.xpertis.com.ar)
