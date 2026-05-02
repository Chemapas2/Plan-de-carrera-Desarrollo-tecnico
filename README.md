# Assessment + Plan de Carrera Técnico

Aplicación en **Streamlit** para analizar el assessment técnico individual de un veterinario o técnico de campo, interpretar sus resultados y generar un **plan de carrera** con comparativas, prioridades de desarrollo y recomendaciones formativas **internas y externas**.

Esta app está orientada a un **técnico evaluado individualmente**.  
No está pensada para elegir un senior, sino para responder de forma objetiva:

- cuál es el nivel actual del técnico,
- en qué áreas destaca,
- qué brechas frenan más su desarrollo,
- dónde conviene reforzar formación base,
- y en qué áreas tiene más potencial de especialización.

---

## Qué hace la app

La aplicación permite:

- cargar un **assessment individual** en Excel (`.xlsm` o `.xlsx`);
- leer los **25 indicadores** y su agrupación por **4 áreas / troncos**;
- calcular:
  - score global medio,
  - comparación frente a objetivo,
  - comparación frente a benchmark / media BBDD;
- clasificar el perfil del técnico en una de estas **6 categorías**:
  - **Básico**
  - **Controla**
  - **Supera**
  - **Certificado**
  - **Excelente**
  - **Máster**
- identificar:
  - **Top 5 fortalezas**
  - **Top 5 debilidades**
- detectar:
  - **áreas prioritarias de formación**
  - **áreas de especialización**
- cargar **uno o varios catálogos de formación** manualmente;
- **autodescubrir catálogos Excel** ya presentes en el repositorio;
- buscar cursos internos por:
  - especie,
  - subespecie,
  - área/tronco,
  - contenido del curso;
- reforzar la propuesta con **búsqueda web externa**;
- añadir **20 preguntas agrupadas** para enriquecer el análisis;
- editar las respuestas;
- decidir si esas respuestas se incorporan o no al informe final;
- mostrar **enlaces web clicables** en los cursos encontrados;
- descargar el informe en:
  - **TXT**
  - **DOCX**
  - **PDF**
  - **Excel**

---

## Estructura del análisis

La app genera un informe con estos bloques:

1. **Resumen ejecutivo del caso**
2. **Perfil del técnico**
3. **Resultados por áreas / troncos**
4. **Top fortalezas y Top debilidades**
5. **Áreas prioritarias de formación**
6. **Áreas de especialización**
7. **Plan de formación**
   - formación base
   - formación especializada
   - propuestas internas
   - propuestas externas
8. **Preguntas y respuestas incorporables al informe**
9. **Informe final editable**

---

## Categorías de perfil

La app clasifica al técnico en 6 niveles visibles:

- **Básico**
- **Controla**
- **Supera**
- **Certificado**
- **Excelente**
- **Máster**

Los umbrales se pueden ajustar en la barra lateral.

### Umbrales por defecto
- **Básico**: < 1.5
- **Controla**: 1.5 – < 2.5
- **Supera**: 2.5 – < 3.5
- **Certificado**: 3.5 – < 4.5
- **Excelente**: 4.5 – < 5.3
- **Máster**: ≥ 5.3

> Estos valores pueden cambiarse en la app para adaptarlos a tu criterio.

---

## Gráficos incluidos

La app genera gráficos de radar para visualizar el perfil técnico.

### 1. Radar global
Comparación de las **4 áreas del assessment**:
- Alimentación
- Sanidad
- Manejo
- Herramientas

### 2. Radars específicos por área
Además del radar global, la app genera **4 radares adicionales**, uno por cada área:
- **Alimentación** → con los indicadores de esa área
- **Sanidad** → con los indicadores de esa área
- **Manejo** → con los indicadores de esa área
- **Herramientas** → con los indicadores de esa área

Estos gráficos:
- se muestran en pantalla,
- y se incorporan también a las exportaciones en **DOCX** y **PDF**.

---

## Entradas necesarias

### 1. Assessment del técnico
Archivo Excel individual en formato:

- `.xlsm`
- `.xlsx`

La app espera una estructura compatible con la herramienta interna y, como mínimo, busca estas hojas:

- `REFERENCIAS`
- `EVALUACION`

### 2. Catálogos de formación
La app admite dos mecanismos:

#### Opción A: carga manual
El usuario sube uno o varios catálogos de formación desde la propia app.

#### Opción B: autodescubrimiento
La app intenta localizar automáticamente catálogos Excel que ya estén en el proyecto.

Ejemplos de archivos compatibles:
- `Resumewn formaciones especificas.xlsx`
- `Formacion_Tecnicos_Vacuno_Carne_Intensivo.xlsx`
- `Especializaciones_Tecnico_Avicultura_Puesta.xlsx`
- `Formacion_Tecnicos_Porcino_Industrial.xlsx`
- `Formacion_Tecnicos_Ovino_Caprino.xlsx`
- `Formacion Vacuno Lechero.xlsx`
- `Formacion_Tecnicos_Vacuno_Lechero_Intensivo.xlsx`

---

## Búsqueda de cursos

La app busca cursos de dos formas:

### Formación interna
A partir de los catálogos Excel:
- especie,
- subespecie,
- contenidos,
- palabras clave,
- área del assessment.

### Formación externa
Si activas la búsqueda externa, la app intenta complementar la propuesta con opciones web y mostrar:
- nombre del programa,
- institución/proveedor,
- enlace,
- y área orientativa.

> Los enlaces se muestran de forma **clicable** dentro de la interfaz.

---

## Estructura recomendada del repositorio

```text
/
├── main.py
├── requirements.txt
├── README.md
├── CTC 4.xlsm
├── Resumewn formaciones especificas.xlsx
├── Formacion_Tecnicos_Vacuno_Carne_Intensivo.xlsx
├── Especializaciones_Tecnico_Avicultura_Puesta.xlsx
├── Formacion_Tecnicos_Porcino_Industrial.xlsx
├── Formacion_Tecnicos_Ovino_Caprino.xlsx
├── Formacion Vacuno Lechero.xlsx
└── Formacion_Tecnicos_Vacuno_Lechero_Intensivo.xlsx
```

No es obligatorio tener todos esos catálogos dentro del repo, pero ayuda a que la app encuentre más cursos automáticamente.

---

## Archivos mínimos para desplegar

Debes dejar en GitHub, como mínimo:

- `main.py`
- `requirements.txt`
- `README.md`

Si además quieres que la app encuentre cursos sin subirlos manualmente, añade también los Excel de formación al repositorio.

---

## Despliegue en GitHub + Streamlit Community Cloud

### Paso 1. Subir a GitHub
Sube al repositorio:
- `main.py`
- `requirements.txt`
- `README.md`
- assessment de ejemplo opcional
- catálogos Excel opcionales

### Paso 2. Crear la app en Streamlit
En Streamlit Community Cloud:

1. Conecta tu cuenta de GitHub.
2. Pulsa **Create app**.
3. Selecciona:
   - repositorio
   - rama
   - archivo principal: `main.py`
4. Pulsa **Deploy**.

---

## Muy importante: dependencias

La app necesita que el archivo de dependencias se llame exactamente:

```text
requirements.txt
```

No sirve:
- `requirements_github.txt`
- `requirements_plan_carrera.txt`
- `requirements_v2.txt`
- `requirements_v5.txt`

Si el nombre no es exactamente `requirements.txt`, Streamlit puede no instalar librerías y aparecerán errores como:

```text
ModuleNotFoundError: No module named 'openpyxl'
```

---

## Contenido recomendado de `requirements.txt`

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

1. Sube el **assessment del técnico**.
2. Sube **uno o varios catálogos** o deja que la app detecte los del repositorio.
3. Ajusta, si quieres, los **umbrales de perfil** en la barra lateral.
4. Revisa:
   - el resumen ejecutivo,
   - la lectura por áreas,
   - fortalezas,
   - debilidades,
   - plan de formación.
5. Añade preguntas desde el desplegable.
6. Edita las respuestas.
7. Marca cuáles se incorporan al informe.
8. Edita el informe final si lo deseas.
9. Descárgalo en:
   - TXT
   - DOCX
   - PDF
   - Excel

---

## Botones disponibles

### Nuevo / borrar evaluación cargada
Limpia la sesión y permite empezar con otro técnico.

### Añadir pregunta al informe
Añade una pregunta del desplegable al bloque de respuestas editables.

### Eliminar esta pregunta
Borra una pregunta ya añadida.

### Refrescar informe con el contenido actual
Reconstruye el informe final incorporando:
- respuestas editadas,
- preguntas seleccionadas,
- y datos actualizados del análisis.

---

## Limitaciones conocidas

- La app depende de una estructura de Excel razonablemente compatible con la herramienta interna.
- Los catálogos de formación no siempre tienen el mismo formato; la app intenta adaptarse, pero puede necesitar ajustes si cambia mucho la plantilla.
- El benchmark se interpreta a partir de los datos del archivo; si el Excel no está actualizado o no se recalculó antes de guardarlo, la lectura puede ser incorrecta.
- La búsqueda web externa depende de la disponibilidad de las fuentes y puede requerir validación manual.
- Algunas propuestas formativas pueden requerir revisión humana si el catálogo tiene información parcial o inconsistente.
- La exportación de gráficos a DOCX/PDF depende de que `kaleido` esté correctamente instalado.
- La app no sustituye el criterio del director técnico; es una herramienta de apoyo a la decisión.

---

## Recomendaciones de uso

- Usa la app como **soporte objetivo**, no como único criterio.
- Revisa siempre:
  - el tronco más débil,
  - los gaps frente a benchmark,
  - la coherencia entre fortalezas y plan de especialización.
- Antes de presentar un informe final, edita el texto generado para ajustarlo al contexto del técnico y del equipo.
- Si la app no encuentra suficientes cursos, añade más catálogos Excel al repositorio o súbelos manualmente en la propia interfaz.
- Revisa las propuestas externas antes de incorporarlas como recomendación definitiva.

---

## Objetivo de negocio

Esta app ayuda a construir una evaluación **rigurosa, defendible y útil** del desarrollo técnico de una persona, y a traducir esa evaluación en un **plan de carrera accionable**.

Sirve para responder con más objetividad:
- dónde está hoy el técnico,
- qué debe reforzar,
- dónde puede especializarse,
- y qué formación concreta tiene más sentido para su evolución.
