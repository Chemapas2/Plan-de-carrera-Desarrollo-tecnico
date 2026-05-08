# TechTeam · Assessment + Plan de Carrera Técnico

Aplicación en **Streamlit** para evaluar el perfil técnico de un profesional, generar un **informe integrado de assessment**, recomendar un **plan de carrera por módulos** y acceder desde la misma app a un **buscador global de formación** en producción animal.

La versión actual integra **tres pestañas** dentro de una misma app:

1. **Assessment**
2. **Plan de carrera**
3. **Buscador global de formación**

La identidad visual de la app y de los informes utiliza solo estos activos de marca:
- **Nutreco**
- **TechTeam**
- **franja rosa corporativa**

---

## Qué hace la app

## 1. Acceso restringido
La app pide contraseña al abrirse.

Contraseña actual en esta versión:
```text
TechTeam2026+
```

Es una barrera básica de acceso en la interfaz.  
Si más adelante quieres mayor seguridad, conviene moverla a `st.secrets`.

---

## 2. Pestaña Assessment
Permite:

- cargar un assessment individual en Excel (`.xlsm` o `.xlsx`);
- leer los **25 indicadores**;
- agruparlos en **4 áreas / troncos**:
  - Alimentación
  - Sanidad
  - Manejo
  - Herramientas
- calcular:
  - score global medio;
  - comparación frente a objetivo;
  - comparación frente a benchmark;
- mostrar:
  - fortalezas;
  - debilidades;
  - áreas prioritarias;
  - áreas de especialización;
  - recomendaciones formativas internas y externas;
  - preguntas guiadas editables;
  - gráficos de radar.

### Clasificación del perfil
La clasificación **no se basa en la media simple del score**.  
La app se ciñe a la lógica del propio Excel de assessment y toma la categoría desde la hoja correspondiente del archivo, para que coincida con la clasificación del modelo interno.

Las categorías visibles son:
- **Básico**
- **Controla**
- **Supera**
- **Certificado**
- **Excelente**
- **Máster**

---

## 3. Pestaña Plan de carrera
Esta pestaña añade una capa nueva sobre el assessment.

Su función es traducir el resultado obtenido en cada tronco a una **recomendación de módulos concretos** del plan de formación corporativo.

### Qué hace
- analiza el resultado por tronco;
- interpreta si el nivel recomendado debe ser:
  - **Certificado**
  - **Máster**
- identifica el módulo concreto recomendado por:
  - especie
  - tronco
  - nivel
- muestra:
  - prioridad;
  - código del módulo;
  - objetivo;
  - perfil destinatario;
  - contenidos clave;
  - formato recomendado;
  - evidencia de aprendizaje;
  - carga estimada;
  - fuente curricular base;
  - justificación directiva.

### Regla de asignación
La lógica utilizada es:

- **Básico / Controla / Supera** → módulo **Certificado**
- **Certificado** → módulo **Certificado** o **Máster** si el tronco es estratégico
- **Excelente** → módulo **Máster**
- **Máster** → orientación a **mentoring / docencia / casos complejos**

### Troncos estratégicos
La app permite marcar troncos como **estratégicos para el rol**, de forma que un técnico en nivel Certificado pueda escalar a recomendación Máster en esos bloques.

---

## 4. Pestaña Buscador global de formación
La app integra un buscador HTML corporativo independiente, accesible desde la tercera pestaña.

Ese buscador mantiene su funcionamiento propio y su informe propio.

### Qué permite hacer
- filtrar por especie;
- filtrar por área del assessment;
- filtrar por tipo de oportunidad;
- priorizar geografía;
- abrir enlaces web visibles y clicables;
- generar su propio informe.

### Especies del buscador
- Avicultura
- Porcino
- Conejos
- Vacuno leche
- Vacuno carne
- Ovino y caprino
- Caballos

### Áreas del buscador
- Nutrición
- Sanidad
- Manejo
- Herramientas

---

## Informe integrado de la app principal

La app principal genera un **único informe integrado** que unifica:

- resumen ejecutivo;
- perfil del técnico;
- resultados por troncos;
- fortalezas y debilidades;
- formación base y especializada;
- preguntas y respuestas incorporadas;
- gráficos de radar;
- y ahora también:
  - **recomendación de módulos del plan de carrera**.

El buscador global mantiene su informe aparte.

---

## Exportaciones disponibles

La app principal permite descargar:

- **TXT**
- **HTML**
- **DOCX**
- **PDF**
- **Excel**

### HTML
El informe HTML sale con plantilla corporativa.

### DOCX y PDF
Incluyen el contenido consolidado del informe y la identidad visual corporativa.

### Excel
Incluye el detalle tabular del análisis y de las recomendaciones.

---

## Gráficos incluidos

La app genera:

- **1 radar global** por áreas;
- **4 radares específicos**:
  - Alimentación
  - Sanidad
  - Manejo
  - Herramientas

Estos gráficos se usan en pantalla y se incorporan al informe exportado cuando la generación de imagen está disponible en el entorno.

---

## Preguntas guiadas

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

---

## Archivos necesarios en el repositorio

Debes dejar en la raíz del repositorio, como mínimo:

- `main.py`
- `requirements.txt`
- `README.md`

Y además, para que la app funcione completa:

- `Matriz_modulos_plan_formacion_tecnica.xlsx`
- `buscador_formacion_produccion_animal_v2.html`
- `Logo Nutreco.jpg`
- `Logo TechTeam 2.jpg`
- `Solapa rosa.jpg`

Opcionalmente, puedes añadir también:
- assessments de ejemplo;
- catálogos Excel de formación;
- programas máster por especie.

---

## Archivos Excel que usa la app

## 1. Assessment
Archivo individual compatible con la herramienta interna, con al menos:
- hoja `REFERENCIAS`
- hoja `EVALUACION`

## 2. Catálogos de formación
La app permite:
- subir catálogos manualmente;
- o detectar catálogos Excel ya presentes en el repositorio.

## 3. Matriz de módulos
Archivo obligatorio para la pestaña **Plan de carrera**:
- `Matriz_modulos_plan_formacion_tecnica.xlsx`

Debe contener, al menos, la hoja:
- `Matriz módulos`

Y opcionalmente:
- `Reglas de asignación`

---

## Estructura recomendada del repositorio

```text
/
├── main.py
├── requirements.txt
├── README.md
├── Matriz_modulos_plan_formacion_tecnica.xlsx
├── buscador_formacion_produccion_animal_v2.html
├── Logo Nutreco.jpg
├── Logo TechTeam 2.jpg
├── Solapa rosa.jpg
├── CTC 3.xlsm
├── CTC 4.xlsm
├── Resumewn formaciones especificas.xlsx
├── PROGRAMA FORMACIÓN CUNICULTURA NANTA 2025.xlsx
├── PROGRAMA MASTER AVICULTURA NANTA 2025.xlsx
├── PROGRAMA MASTER PEQUEÑOS RUMIANTES 2025.xlsx
├── PROGRAMA MASTER PORCINO NANTA 2025.xlsx
├── PROGRAMA MASTER RUMIANTES VC2025.xlsx
└── PROGRAMA MASTER VACUNO DE LECHE 2025.xlsx
```

---

## Despliegue en Streamlit Community Cloud

### Paso 1. Subir a GitHub
Sube al repositorio:
- `main.py`
- `requirements.txt`
- `README.md`
- la matriz de módulos;
- el buscador HTML;
- los logos corporativos;
- y, si quieres, los catálogos y programas Excel.

### Paso 2. Crear la app
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
2. Introduce la contraseña.
3. Sube el assessment del técnico.
4. Sube catálogos de formación o deja que la app detecte los del repositorio.
5. Selecciona especie y subespecie.
6. Revisa la pestaña **Assessment**.
7. Revisa la pestaña **Plan de carrera** para ver los módulos recomendados.
8. Marca, si procede, los troncos estratégicos.
9. Añade preguntas desde el desplegable.
10. Edita el informe final.
11. Descárgalo en el formato deseado.
12. Si quieres ampliar opciones, usa la pestaña **Buscador global de formación**.

---

## Qué muestra la pestaña Plan de carrera

Para cada tronco, la app propone:
- nivel recomendado;
- prioridad;
- módulo;
- objetivo;
- contenidos;
- evidencia;
- justificación.

Esto permite que el director técnico traduzca el assessment en una decisión concreta de itinerario formativo.

---

## Limitaciones conocidas

- La contraseña está escrita en el código; es una barrera básica.
- La estructura del assessment debe ser compatible con la plantilla esperada.
- La matriz de módulos debe existir y estar bien formada para que funcione la pestaña Plan de carrera.
- Los catálogos Excel pueden requerir cierta homogeneidad para maximizar resultados.
- La búsqueda externa depende de la disponibilidad de fuentes web.
- La generación de imágenes para DOCX/PDF depende del entorno y de las librerías instaladas.
- El buscador HTML integrado mantiene su informe separado del informe integrado principal.

---

## Recomendaciones

- Mantén la matriz de módulos en la raíz del repositorio.
- Mantén los logos corporativos en la raíz del proyecto.
- Revisa manualmente la vigencia de la formación externa.
- Usa la pestaña Plan de carrera como capa prescriptiva del assessment.
- Si quieres mayor seguridad, migra la contraseña a `st.secrets`.

---

## Objetivo de negocio

Esta app ayuda a:

- evaluar con rigor el perfil técnico de una persona;
- identificar gaps y fortalezas;
- construir un plan de carrera accionable;
- recomendar módulos concretos por especie, tronco y nivel;
- y mantener, dentro del mismo entorno, acceso a un buscador global de formación.

