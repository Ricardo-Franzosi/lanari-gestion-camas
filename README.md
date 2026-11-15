# Gestión de camas · Instituto Lanari

Panel web de **consulta** del estado de camas e internaciones del Instituto Lanari (IDIM-UBA). Muestra:

- **Mapa de camas:** 1º piso con UBE, 2º piso con Hemodinamia, UTI y Guardia. Compara la ocupación contra la matriz de 58 camas y marca aislamiento y pases.
- **Movimientos y agenda:** altas, pases, camas esperando, agenda semanal, próximos ingresos y traslados pendientes.
- **Espera y suspensiones:** ingresos de los próximos 14 días, OP pendientes, reprogramaciones y suspensiones por falta de cama.
- **Indicadores:** pedidos por mes, Sala/UTI, cápita/extracápita y calidad del registro.

> **Este repositorio no contiene datos de pacientes.** La página lee la planilla de Google *“SOLICITUD DE CAMAS PARA INTERNACION”* en el momento de cada consulta, a través de un Google Apps Script, y solo después de verificar el acceso. Nunca subas exports, Excel ni CSV con datos (el `.gitignore` los bloquea).

## Cómo funciona el acceso

1. La persona ingresa su mail **@lanari.uba.ar**.
2. El Apps Script le manda un **código de 6 dígitos** a ese mail. Vence en 10 minutos y admite 5 intentos.
3. Con el código correcto se abre una sesión de **6 horas**. Cada consulta y cada actualización (automática cada 5 minutos) trae el estado actual de la planilla.

Opcional:
- **Pestaña `ACCESO`** en la planilla, con mails en la columna A: si existe, solo entran esos mails del dominio. Si no existe, entra cualquier mail @lanari.uba.ar.
- **Pestaña `LOG ACCESOS`**: si existe, se registra fecha, mail y acción de cada ingreso y consulta.

## Puesta en marcha

### 1. Apps Script (una sola vez)

Lo tiene que hacer alguien con acceso de **edición** a la planilla. Los mails con el código salen desde la cuenta de esa persona.

1. Abrir la planilla → **Extensiones → Apps Script**.
2. Borrar el contenido de `Código.gs` y pegar el de [`apps-script/Code.gs`](apps-script/Code.gs). Guardar.
3. En el selector de funciones elegir **`autorizar`** → **Ejecutar** → aceptar los permisos (leer la planilla y enviar mails).
4. **Implementar → Nueva implementación** → tipo **Aplicación web**:
   - *Descripción:* `panel camas`
   - *Ejecutar como:* **Yo**
   - *Quién tiene acceso:* **Cualquier usuario**
5. **Implementar** y copiar la **URL de la aplicación web** (termina en `/exec`).

> Si después se modifica el script: **Implementar → Gestionar implementaciones → editar (lápiz) → Versión: nueva → Implementar**. Así la URL sigue siendo la misma.

### 2. Configurar la página

Editar [`config.js`](config.js) y pegar la URL:

```js
window.LANARI_CONFIG = {
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/XXXXXXXX/exec',
  DOMINIO: 'lanari.uba.ar',
};
```

### 3. GitHub Pages

**Settings → Pages → Source: Deploy from a branch → `main` / `(root)` → Save.** El panel queda en `https://ricardo-franzosi.github.io/lanari-gestion-camas/`.

## Sin conexión al servidor

El botón **Cargar export** permite abrir una descarga de la planilla (Archivo → Descargar → Microsoft Excel). El archivo se procesa solo en el navegador y no se sube a ningún lado.

## Seguridad: lo que conviene saber

- La URL del Apps Script es pública, pero sin un código válido no entrega datos. Los códigos y las sesiones se guardan con vencimiento en la caché del script.
- Cualquier persona con un mail @lanari.uba.ar puede pedir acceso, salvo que se use la pestaña `ACCESO`.
- La planilla en Google **no** debe estar compartida como “cualquier persona con el enlace”.
- Cuota de envío de mails de Apps Script: 100 por día con cuenta gratuita y 1500 con Google Workspace.
- La página tiene `noindex` para que no aparezca en buscadores.
