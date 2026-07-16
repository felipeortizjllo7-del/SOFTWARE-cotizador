# Puente HTML → .exe (cotizaciones de clientes)

El **HTML** (clientes) envía cada cotización a una **Hoja de Google** mediante un
pequeño servicio gratis (Apps Script). El **.exe** (INNOBA) las importa al historial.

Necesito que hagas esto **una sola vez** y me pases la URL:

## Pasos

1. Entra a https://sheets.google.com y crea una hoja nueva, ponle nombre
   **"Cotizaciones INNOBA"**.
2. Menú **Extensiones → Apps Script**.
3. Borra lo que haya y **pega TODO este código**:

```javascript
function doPost(e){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Cotizaciones") || ss.insertSheet("Cotizaciones");
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch(err){ d = {}; }
  sh.appendRow([ new Date(), d.id||"", d.cliente||"", d.asesor||"", d.asesor_tel||"",
    d.email||"", d.fecha||"", (d.destinos||[]).join(", "), d.total||"",
    d.fechas_viaje||"", JSON.stringify(d) ]);
  return ContentService.createTextOutput("ok");
}
function doGet(e){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Cotizaciones");
  const out = [];
  if (sh){
    const rows = sh.getDataRange().getValues();
    for (let i=0;i<rows.length;i++){
      try { out.push(JSON.parse(rows[i][10])); } catch(err){}
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
```

4. Clic en **Implementar → Nueva implementación**.
5. En "Tipo", elige **Aplicación web**.
6. Configura:
   - **Ejecutar como:** Yo (tu cuenta)
   - **Quién tiene acceso:** **Cualquier persona**
7. Clic **Implementar** y **Autoriza** los permisos (te pedirá permiso una vez).
8. Copia la **URL de la aplicación web** (termina en **/exec**).
9. **Pásame esa URL** por el chat.

Con esa URL yo la configuro en el HTML (para que envíe) y en el .exe (para que
importe), y publico la nueva versión. A partir de ahí:
- Cuando un cliente genera una cotización en el HTML, se envía a esa hoja.
- Tu .exe, al abrir "Cotizaciones", importa las nuevas automáticamente para darles
  seguimiento.
