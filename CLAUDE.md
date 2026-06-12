# La Posta — Instrucciones del proyecto

App web de punto de venta (POS) para "La Posta", comercio de granja.
**Stack:** HTML/CSS/JS estático + Google Sheets (BD) + Google Apps Script (backend `backend.gs`). Sin servidor propio.
Se sirve desde **GitHub Pages**: `https://josedgriffo-gif.github.io/laposta/` · repo `josedgriffo-gif/laposta`, rama `main`.

## REGLA CRÍTICA — Cierre de cada cambio (4 capas)

El usuario pierde mucho tiempo cuando un cambio "queda a medias". Un cambio NO está completo hasta cerrar estas capas:

1. **Código local** — editar archivos + subir el badge de versión.
2. **Git push** — `git add + commit + push` a `main`. Publica el FRONTEND en GitHub Pages.
   - Hacerlo **de inmediato y SIN preguntar** al terminar. NO reportar "listo" sin haber pusheado.
   - (Un hook `Stop` pushea automático como red de seguridad — ver `.claude/settings.local.json`.)
3. **Apps Script** — SOLO si el cambio tocó `backend.gs`: el usuario debe copiar todo `backend.gs` y pegarlo en el editor de Apps Script + **Implementar**. Claude NO tiene acceso a esto.
   - Si tocaste `backend.gs`, avisarlo **destacado y al principio** del mensaje de cierre. Hasta que el usuario lo pegue, la función nueva da error aunque el badge muestre la versión nueva.
4. **Verificar** — GitHub Pages tarda 1-2 min en reconstruir; recién ahí Ctrl+F5 muestra la versión nueva.

Regla rápida: **solo HTML → con el push alcanza.** **Tocó `backend.gs` → además hace falta el pegado manual en Apps Script.**

## Versionado — REGLA CRÍTICA

**TODO cambio sube el número de versión, sin excepción** (hasta el fix más chico de una línea). NUNCA reutilizar el número anterior, aunque sea en la misma sesión.

**Por qué (no es prolijidad, es funcional para José):** el número de versión es su **semáforo de deploy**. Es la única forma que tiene de verificar, mirando la app, si lo que ve es la versión vieja o la nueva — si el número no cambia, no sabe si el cambio se subió o no. Olvidarse de subirlo le hace perder tiempo recargando sin saber qué está mirando.

UN SOLO número para todo el sistema. Los 4 archivos deben mostrar el MISMO número: `index.html`, `admin.html`, `dueno.html` (badge en el `<h1>`) y `backend.gs` (comentario de cabecera). NO versión por archivo (eso lo maneja git).

El hook `Stop` (ver `.claude/settings.local.json`) **no pushea** si detecta código `.html`/`.gs` commiteado sin bump de versión — red de seguridad, pero la responsabilidad de subirlo es de Claude en cada cambio.

## Arquitectura de pantallas

- `index.html` — Caja (facturación táctil). PIN_Caja.
- `admin.html` — Panel Admin. PIN_Admin (llave maestra). Tabs operativos: productos, compras, ctacte, gastos, stock, anular, cierre, exportar, config (`TABS_ADMIN`).
- `dueno.html` — Socios. PIN_Socios. Toda la **analítica vive acá**: tabs HOY / PERÍODO / Accionistas / Volumen / Afluencia.
- **Admin NO tiene tabs de análisis**: los alcanza por el menú "Informes ▾" que linkea a `dueno.html#evolucion`, `#accionistas`, `#afluencia`. Si agregás un análisis nuevo, va como tab en `dueno.html` + link en ese menú (no dupliques en admin).

## Cómo trabaja el usuario (José)

- Analista funcional, no programador. Define la funcionalidad; Claude escribe el código. Responder en español rioplatense.
- No tomar decisiones funcionales sin consultar. Dar opciones concretas (A/B/C) con recomendación, no preguntas abiertas.
- Antes de cambios de navegación/layout, describir cómo va a quedar y esperar confirmación.
- Dominio de precisión: no simplificar datos (cálculos, montos, costos) sin preguntar.

## Datos sensibles (no romper)

- Columnas de `BD_Ventas` y `Detalle_Ventas` mantienen compatibilidad con Excel existente — no renombrar sin consultar.
- El backend hay que pegarlo a mano en Apps Script; el usuario a veces lo tiene abierto en Bloc de notas (EBUSY) — pedir que lo cierre.
