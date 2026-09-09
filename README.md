# Valorización SIS · USPP Satipo

Dashboard de cuánto genera el hospital (y, cuando agregues más
establecimientos, los otros niveles) en valorización SIS: por servicio,
por profesional de salud, por componente (servicio/medicamento/insumo),
y cuánto se va a per cápita.

**No depende del proyecto de Producción del Digitador.** El cruce es
por N° de FUA directamente contra el RAtenDet, para traer el DNI y
nombre del **profesional de salud** (no del digitador).

## 1. Actualizar los datos

```bash
pip install pandas python-calamine
python consolidar_valorizacion.py
```

Antes de correrlo, abre `consolidar_valorizacion.py` y revisa las dos
rutas del inicio del archivo:

```python
VALORIZ_DIR  = Path(r"D:\ANALISIS DE DATOS\VALORIZACION BASE DE DATOS")
RATENDET_DIR = Path(r"C:\Users\USUARIO\Desktop\BASE DE DATOS\Escritorio\USPP\BASE DE DATOS 2")
```

- **VALORIZ_DIR**: carpeta con tus Excel de valorización (uno o varios —
  si hay varios que se repiten/acumulan, se deduplica por FUA y gana el
  más reciente).
- **RATENDET_DIR**: carpeta con los `RAtenDet_*.xlsx`. De ahí se saca
  el **DNI Resp. Aten.** y **Profesional** (el profesional de salud que
  atendió), no el digitador. También acepta `dig_AAAA_MM.json` si
  alguna vez quieres reusar esa fuente, pero ojo: esos archivos no
  traen el profesional de salud, solo sirven para lo que sea del
  digitador.

Al correrlo genera `data/val_AAAA_MM.json` (uno por mes) + `manifest.json`,
y en consola te muestra qué % de la valorización encontró cruce con el
profesional, y cuánta atención/valor se va a per cápita según el Anexo.

**Estado actual de los datos incluidos (enero–mayo 2026):** son reales,
generados desde tu Excel de valorización. El cruce con el profesional de
salud está en 0% porque el único RAtenDet que me pasaste es de agosto
(otro mes) — en cuanto corras el script con tus RAtenDet de enero-mayo,
la pestaña "Por profesional" se llena. Mientras tanto, "Resumen" y
"Per cápita" ya funcionan al 100% porque no necesitan ese cruce.

## 2. Diagnósticos con cobertura per cápita (Anexo 19)

El script trae embebidos los 354 códigos CIE10 del Anexo que compartiste
(`DIAGNOSTICOS_PERCAPITA` en `consolidar_valorizacion.py`). Cada atención
queda marcada con `diagnostico_percapita: true/false`, y la pestaña
**"Per cápita"** del dashboard muestra cuántas atenciones y cuánto Valor
Bruto SIS caen ahí — es decir, lo que "se pierde" para pago por
prestación porque ya está cubierto por la cápita.

Según la cláusula 22.4 del convenio SIS–Junín que compartiste, esta regla
(diagnóstico en el Anexo → per cápita) aplica tal cual para
establecimientos de **Nivel II con población adscrita** (tu EESS actual,
categoría II-1). Para establecimientos de **Nivel I**, la cláusula 22.3
agrega 4 excepciones que sí se pagan por prestación aunque el
diagnóstico esté en el Anexo (especialistas con RNE, consejería
nutricional en HTA/DM, tamizaje de cáncer, cuidados paliativos a
domicilio) — **esa parte todavía no está en el script**, porque
necesito confirmar si tu Excel trae las columnas de RNE/CPMS antes de
programarla. Avísame cuando tengas datos de un EESS de Nivel I y lo
armamos.

Un dato para revisar: en la muestra actual hay atenciones marcadas
`Capita: "Cap"` cuyo diagnóstico **no** está en el Anexo 19 (debería ser
al revés según la 22.4). Puede ser normal — quizás hay más motivos para
`Capita: "Cap"` en tu EESS más allá del diagnóstico — pero conviene que
lo revises con lo que ya conoces del proceso, antes de tomar el % de
"Per cápita" del dashboard como cifra oficial de conciliación de pago.

## 3. Abrir el dashboard

**En GitHub Pages (o cualquier hosting estático):** solo sube la
carpeta tal cual. `index.html` funciona directo, sin configurar nada
más.

**Para probarlo en tu computadora antes de subirlo:** los navegadores
bloquean que `index.html` lea los `.json` si lo abres con doble clic
(protocolo `file://`). Corre esto en la carpeta del proyecto:
```bash
python -m http.server 8765
```
y abre `http://localhost:8765/index.html`.

## 4. Qué trae cada pestaña

- **Resumen**: servicios con mayor valorización, valor por componente
  (servicio/medicamento/insumo/apoyo al diagnóstico), diagnósticos más
  frecuentes, % Observado.
- **Por profesional**: tarjetas por profesional de salud (DNI + nombre
  vía el cruce) — Valor Bruto SIS, N° de atenciones, promedio, y % de
  lo suyo que se va a per cápita.
- **Matriz por día**: Valor Bruto SIS por profesional × día de atención.
- **Ranking**: ordenable por total, promedio o N° de atenciones — nota
  aclaratoria de que es para evaluación estadística, no mide calidad
  clínica ni carga real de trabajo.
- **Per cápita**: cuántas atenciones y cuánto valor caen en el Anexo,
  y qué diagnósticos/servicios concentran esa "pérdida".

## 5. Datos de pacientes

Nombre, DNI y fecha de nacimiento del **paciente** nunca se guardan en
los JSON. Sí se guarda el DNI y nombre del **profesional de salud**
(es justo lo que pediste cruzar).

## Estructura

```
valorizacion_satipo/
├── index.html
├── consolidar_valorizacion.py
├── logo.png
└── data/
    ├── manifest.json
    └── val_2026_01.json ... val_2026_05.json
```
