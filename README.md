# Rendimientos Hydraflora

App web instalable (PWA) para que el líder de campo registre los rendimientos de las labores del cultivo, sin señal, y los sincronice después con una hoja de Google. El mismo enlace abre el dashboard en celular o computador.

## Qué hace

- **Registrar:** empleado(s) con sugerencias al escribir parte del nombre o la cédula, labor con sugerencias, hora de inicio y fin, bloque, lote y camas (se tocan en una cuadrícula o se escriben como `1-10, 15`).
  - Labores en **Plantas**: bloque, lote, camas y plantas son obligatorios. Las plantas se calculan solas con el plano de producción; el líder puede corregirlas.
  - Labores en **Tallos** (Cortar flor): tallos obligatorios.
  - Labores en **Tiempo / Tiempo GR**: solo horas.
  - **Labor grupal:** se agregan todas las personas; la cantidad se reparte en partes iguales y cada persona queda con su fila.
  - Descuentos de desayuno, almuerzo y pausa se marcan solos si el horario cruza esa hora.
- **Registros:** lo anotado en el día, editar o borrar, horas por persona con alerta de jornada incompleta o horarios cruzados.
- **Dashboard:** rendimiento promedio y máximo por labor, horas por labor, rendimiento por persona, tendencia diaria o semanal, revisión de horas por día y avance por bloque frente al plano. Descarga a Excel con las columnas del formato actual.
- **Ajustes:** cargar el Excel de empleados (hoja DATA EMPLEADOS), el de labores y el plano de producción; horario de descuentos; conexión al servidor.

Fórmulas, iguales al Excel: `TIEMPO LAB = (fin − inicio) − desayuno − pausas − almuerzo` y `RENDIMIENTO = cantidad ÷ TIEMPO LAB`.

## Instalación

### 1. Servidor (Google Sheets)

1. Crea una Google Sheet vacía, por ejemplo *Rendimientos Hydraflora*.
2. Menú **Extensiones > Apps Script**. Borra lo que haya y pega `apps-script/Code.gs`.
3. Cambia `CLAVE_INICIAL` por una clave tuya. Guarda y ejecuta la función **configurar** (acepta los permisos). Se crean las pestañas RENDIMIENTOS, EMPLEADOS, LABORES y CAMAS.
4. **Implementar > Nueva implementación > Aplicación web.** Ejecutar como: *Yo*. Acceso: *Cualquier persona*. Copia la URL que termina en `/exec`.

### 2. App

La app es un sitio estático. Con GitHub Pages: sube esta carpeta a un repositorio, **Settings > Pages > Deploy from branch > main / (root)**. El enlace queda como `https://<usuario>.github.io/<repositorio>/`.

### 3. En cada celular

1. Abre el enlace en Chrome y elige **Instalar app** (o menú ⋮ > *Agregar a pantalla principal*).
2. **Ajustes > Sincronización:** pega la URL `/exec` y la clave, y toca *Guardar y probar*.
3. **Ajustes > Base de empleados:** carga el Excel de empleados (solo hace falta en un celular; los demás la bajan del servidor).

Cuando entre o salga personal, vuelve a cargar el Excel de empleados. Quien ya no aparezca queda como RETIRADO: deja de salir en las sugerencias y su historial se conserva.

## Publicar una versión nueva

Cambia `VERSION` en `sw.js` para que los celulares descarguen los archivos nuevos.

## Datos

- `data/labores.json` y `data/camas.json` vienen de *DATA LABORES CAMPO.xlsx* y del plano de producción. Se reemplazan desde Ajustes.
- La lista de empleados **no** va en el repositorio (tiene cédulas): llega del servidor o del Excel. La carpeta `local/` es solo para pruebas y está en `.gitignore`.
