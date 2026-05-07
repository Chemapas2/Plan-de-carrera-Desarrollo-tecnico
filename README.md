# TechTeam · Assessment + Plan de Carrera Técnico

Aplicación en **Streamlit** para evaluar el perfil técnico de un profesional, construir un **plan de carrera** y acceder desde la misma app a un **buscador corporativo de cursos, congresos, seminarios y formación en producción animal**.

La solución integra dos piezas:

1. **App principal de plan de carrera**
   - carga del assessment individual en Excel;
   - análisis de resultados;
   - fortalezas, debilidades y prioridades de desarrollo;
   - recomendaciones formativas internas y externas;
   - preguntas guiadas editables;
   - exportación de informes.

2. **Buscador corporativo de formación**
   - acceso directo desde la app principal;
   - filtros por especie y área;
   - enlaces web visibles y clicables;
   - descarga de informe con plantilla corporativa.

---

## Funcionalidades principales

### 1. Acceso restringido
La app solicita contraseña al abrirse.

### 2. Branding corporativo
La interfaz incorpora los logos de:
- Cátedra Nanta de Ganadería de Precisión
- Nutreco
- TechTeam

y mantiene una plantilla visual corporativa en la interfaz y en los informes exportados.

### 3. Assessment técnico individual
La app permite cargar un assessment individual en formato:
- `.xlsm`
- `.xlsx`

A partir del Excel:
- lee los **25 indicadores**;
- agrupa los resultados por **4 áreas / troncos**;
- calcula score global y comparativas;
- clasifica el perfil del técnico.

### 4. Clasificación del perfil
La clasificación visible contempla estas **6 categorías**:
- **Básico**
- **Controla**
- **Supera**
- **Certificado**
- **Excelente**
- **Máster**

Los umbrales pueden ajustarse desde la barra lateral.

### 5. Diagnóstico del perfil
La app muestra:
- score global;
- comparación frente a objetivo;
- comparación frente a benchmark;
- Top 5 fortalezas;
- Top 5 debilidades;
- áreas prioritarias de desarrollo;
- áreas de especialización.

### 6. Formación recomendada
La app combina:
- **formación interna** a partir de catálogos Excel;
- **formación externa** mediante búsqueda web reforzada.

Los cursos se muestran con:
- área;
- programa;
- institución;
- tipo;
- duración;
- modalidad;
- ubicación;
- fuente;
- **enlace clicable**.

### 7. Preguntas guiadas
La app incluye **20 preguntas** agrupadas en 4 bloques:
- Potencial y proyección
- Fortalezas y especialización
- Brechas y riesgos
- Plan de acción y formación

Cada pregunta:
- se añade desde un desplegable;
- genera una respuesta inicial automática;
- puede editarse;
- puede incluirse o no en el informe final.

### 8. Gráficos de araña
La app genera:
- **1 radar global** con las 4 áreas del assessment;
- **4 radares específicos**, uno por cada área:
  - Alimentación
  - Sanidad
  - Manejo
  - Herramientas

Estos gráficos se muestran en pantalla y se incorporan a DOCX y PDF cuando la generación de imagen está disponible en el entorno. La app base ya incluye este bloque de radares para el informe. fileciteturn13file6

### 9. Informe final editable
La aplicación genera un informe editable con:
- resumen ejecutivo;
- perfil del técnico;
- resultados por áreas;
- fortalezas y debilidades;
- plan formativo;
- preguntas incorporadas;
- gráficos del assessment.

### 10. Exportaciones
La app integrada permite descargar:
- **TXT**
- **HTML**
- **DOCX**
- **PDF**
- **Excel**

La versión HTML se genera con plantilla corporativa y branding visual.

---

## Buscador corporativo integrado

La app principal incorpora acceso directo al archivo HTML de búsqueda avanzada.

El buscador incluye:
- filtro por **especie**:
  - Avicultura
  - Porcino
  - Conejos
  - Vacuno leche
  - Vacuno carne
  - Ovino y caprino
  - Caballos
- filtro por **área del assessment**:
  - Nutrición
  - Sanidad
  - Manejo
  - Herramientas
- filtro por tipo de oportunidad;
- prioridad geográfica;
- número de resultados.

La versión HTML del buscador ya trabaja con identidad visual corporativa y permite descarga de informe en **HTML, Word y PDF**. fileciteturn13file1 fileciteturn13file2

---

## Archivos necesarios en el repositorio

Deja en la raíz del repositorio, como mínimo:

- `main.py`
- `requirements.txt`
- `README.md`
- `buscador_formacion_produccion_animal_v2.html`

Y para mantener el branding corporativo:

- `Catedra Nanta Zaragoza.jpg`
- `Logo Nutreco.jpg`
- `Logo TechTeam 2.jpg`
- `Solapa rosa.jpg`

Además, puedes dejar catálogos Excel en el repositorio para que la app los detecte automáticamente.

---

## Archivos Excel que puede usar la app

### Assessment
Archivo individual con estructura compatible con la herramienta interna y, como mínimo:
- hoja `REFERENCIAS`
- hoja `EVALUACION`

### Catálogos de formación
La app admite:
- subida manual de uno o varios catálogos;
- autodescubrimiento de catálogos ya presentes en el repositorio.

En la app base del plan de carrera ya se contemplaba la carga de múltiples catálogos y el uso opcional de búsqueda web externa. fileciteturn13file7

---

## Estructura recomendada del repositorio

```text
/
├── main.py
├── requirements.txt
├── README.md
├── buscador_formacion_produccion_animal_v2.html
├── Catedra Nanta Zaragoza.jpg
├── Logo Nutreco.jpg
├── Logo TechTeam 2.jpg
├── Solapa rosa.jpg
├── assessment_ejemplo.xlsm
├── catalogo_1.xlsx
├── catalogo_2.xlsx
└── ...
```

---

## Despliegue en Streamlit Community Cloud

### 1. Subir a GitHub
Sube al repositorio:
- `main.py`
- `requirements.txt`
- `README.md`
- `buscador_formacion_produccion_animal_v2.html`
- logos corporativos
- catálogos Excel opcionales

### 2. Crear la app
En Streamlit Community Cloud:
1. conecta tu cuenta de GitHub;
2. pulsa **Create app**;
3. selecciona:
   - repositorio
   - rama
   - archivo principal: `main.py`
4. pulsa **Deploy**.

---

## Dependencias

El archivo debe llamarse exactamente:

```text
requirements.txt
```

Contenido recomendado:

```txt
streamlit>=1.44.0
pandas>=2.2.0
openpyxl>=3.1.0
xlsxwriter>=3.2.0
plotly>=5.20.0
requests>=2.31.0
beautifulsoup4>=4.12.0
python-docx>=1.1.0
reportlab>=4.0.0
kaleido>=0.2.1
```

---

## Cómo usar la app

1. Abre la app.
2. Introduce la contraseña de acceso.
3. Sube el assessment del técnico.
4. Sube uno o varios catálogos o deja que la app use los del repositorio.
5. Selecciona especie y subespecie.
6. Ajusta los umbrales si lo deseas.
7. Revisa:
   - resumen ejecutivo;
   - resultados por áreas;
   - fortalezas y debilidades;
   - plan formativo;
   - gráficos de araña.
8. Añade preguntas desde el desplegable.
9. Edita el informe final.
10. Descárgalo en el formato deseado.
11. Accede al buscador de formación integrado si quieres ampliar opciones.

---

## Qué muestra la app en formación

La app presenta la formación en cuatro bloques:
- Formación base interna
- Formación base externa
- Formación especializada interna
- Formación especializada externa

Los enlaces de cursos se presentan como enlaces clicables en tabla, mediante una columna de apertura directa. La app base ya lo hacía con `LinkColumn` en Streamlit. fileciteturn13file8

---

## Limitaciones conocidas

- La contraseña integrada en el código es una barrera básica, no un sistema de seguridad corporativa avanzada.
- La estructura del assessment debe ser compatible con la plantilla esperada.
- Los catálogos Excel pueden requerir cierta homogeneidad para maximizar resultados.
- La búsqueda externa depende de la disponibilidad de fuentes web.
- La exportación gráfica a DOCX/PDF depende del entorno y de las librerías instaladas.
- El buscador HTML integrado funciona como recurso complementario dentro del ecosistema de la app principal.

---

## Recomendaciones

- Mantén los catálogos en la raíz del repositorio para mejorar la detección automática.
- Revisa manualmente la vigencia de las propuestas externas.
- Usa el informe como herramienta de soporte a la decisión, no como único criterio.
- Si quieres mayor seguridad, mueve la contraseña a `st.secrets` en una siguiente versión.

---

## Objetivo de negocio

Esta app ayuda a:
- evaluar con rigor el perfil técnico de una persona;
- identificar gaps y fortalezas;
- construir un plan de carrera accionable;
- y localizar formación útil y accesible dentro del mismo entorno de trabajo.

