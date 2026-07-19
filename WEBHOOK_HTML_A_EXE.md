# Puente HTML → .exe (cotizaciones de clientes) + PRIVACIDAD

El **HTML** (clientes) envía cada cotización a una **Hoja de Google**; el **.exe**
(INNOBA) las importa. Para que **los clientes NO puedan leer** lo que cotizan los
demás, la lectura requiere una **clave** que solo tiene el .exe.

## Actualizar el código del Apps Script (para activar la privacidad)

1. Abre tu Hoja **"Cotizaciones INNOBA" → Extensiones → Apps Script**.
2. **Borra todo** y **pega este código** (ya trae la clave):

```javascript
var CLAVE = "inb_9f3Kx72Qp_seg2026";

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
  // Solo el .exe (que tiene la clave) puede LEER. Sin clave -> nada.
  if (!e || !e.parameter || e.parameter.key !== CLAVE) {
    return ContentService.createTextOutput("[]")
      .setMimeType(ContentService.MimeType.JSON);
  }
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

3. Guarda (💾).
4. **Implementar → Administrar implementaciones → ✏ Editar → Versión: "Nueva versión" → Implementar.**
   (La **URL no cambia**; solo se actualiza el código.)

Listo. A partir de ahí:
- El **HTML** solo **envía** (POST), nunca lee.
- Aunque un cliente vea el código del HTML y encuentre la URL, un GET **sin la clave**
  devuelve vacío → **no puede ver** las cotizaciones de nadie.
- El **.exe** sí lee (tiene la clave embebida) e importa todo para darle seguimiento,
  **editar** y **generar el PDF** de cada cotización, incluidas las hechas en el HTML.

> Nota: el .exe ya funciona aunque no actualices el código todavía; pero la
> privacidad de lectura queda activa solo cuando pegues este código y reimplementes.
