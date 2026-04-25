# Assessment + Plan de Carrera Técnico

Aplicación en **Streamlit** para analizar el assessment técnico individual de un veterinario o técnico de campo, interpretar sus resultados y generar un **plan de carrera** con comparativas, prioridades de desarrollo y recomendaciones formativas **internas y externas**.

Esta es la versión correcta de la app actual:  
**no está orientada a elegir un senior**, sino a construir un **itinerario de desarrollo del técnico evaluado**.

---

## Qué hace la app

La aplicación permite:

- cargar un **assessment individual** en Excel (`.xlsm` o `.xlsx`);
- leer los **25 indicadores** y su agrupación por **troncos**;
- calcular:
  - score global medio,
  - comparación frente a objetivo,
  - comparación frente a benchmark / media BBDD;
- proponer una **categoría de perfil**:
  - Básico
  - Supera
  - Certificado
  - Excelente
  - Máster
- identificar:
  - **Top 5 fortalezas**
  - **Top 5 debilidades**
- detectar:
  - **áreas prioritarias de formación**
  - **áreas de especialización**
- cargar **uno o varios catálogos de formación** manualmente;
- **autodescubrir catálogos** de formación que ya estén en el repositorio;
- buscar cursos internos por:
  - especie,
  - subespecie,
  - tronco/área,
  - contenido del curso,
  - archivo fuente;
- reforzar la propuesta con **búsqueda web externa**;
- añadir **20 preguntas agrupadas** para enriquecer el análisis;
- editar las respuestas;
- decidir si esas respuestas se incorporan o no al informe final;
- descargar el resultado en **TXT** y **Excel**.

---

## Enfoque de la app

La lógica de esta aplicación está pensada para construir un **plan de carrera técnico individual**.

No responde a la pregunta:

> “¿Quién debe ser senior?”

Sino a estas otras:

- ¿Cuál es el nivel actual del técnico?
- ¿En qué destaca realmente?
- ¿Qué brechas frenan más su desarrollo?
- ¿Qué áreas conviene reforzar primero?
- ¿En qué ámbitos tiene más potencial de especialización?
- ¿Qué formación interna y externa tiene más sentido en su caso?

---

## Entradas necesarias

### 1. Assessment del técnico
Archivo Excel individual en formato:

- `.xlsm`
- `.xlsx`

La app espera una estructura compatible con la herramienta interna de assessment y, como mínimo, busca estas hojas:

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

> Recomendación: si quieres maximizar el número de cursos detectados, deja los catálogos en la raíz del repositorio y además podrás subir otros adicionales manualmente.

---

## Qué devuelve la app

### 1. Resumen ejecutivo del caso
Un informe breve en 10 líneas con:
- perfil propuesto,
- lectura general frente a objetivo y benchmark,
- fortalezas,
- debilidades,
- prioridad de desarrollo,
- recomendación general de enfoque.

### 2. Lectura por troncos
Incluye:
- score técnico por tronco,
- referencia aproximada,
- benchmark aproximado,
- gráfico de barras,
- gráfico de araña.

### 3. Indicadores clave
- Top 5 fortalezas
- Top 5 debilidades

### 4. Plan de carrera y formación
- áreas prioritarias de desarrollo;
- áreas recomendadas de especialización;
- formación base interna;
- formación especializada interna;
- formación base externa;
- formación especializada externa.

### 5. Preguntas-respuestas incorporables al informe
La app incluye un bloque con **20 preguntas agrupadas** en 4 categorías:
- Potencial y proyección
- Fortalezas y especialización
- Brechas y riesgos
- Plan de acción y formación

Cada pregunta:
- puede añadirse al análisis,
- genera una respuesta inicial automática,
- puede editarse manualmente,
- y puede incluirse o no en el informe final.

### 6. Informe final editable
El usuario puede modificar libremente el informe antes de descargarlo.

---

## Criterio de clasificación del perfil

La app propone una categoría de perfil a partir del **score global medio** del assessment.

Por defecto, usa estos umbrales:

- **Básico**: < 2.0
- **Supera**: 2.0 – < 3.0
- **Certificado**: 3.0 – < 4.0
- **Excelente**: 4.0 – < 5.0
- **Máster**: ≥ 5.0

Estos umbrales se pueden ajustar en la **barra lateral**.

---

## Cómo interpreta fortalezas y debilidades

### Fortalezas
Se priorizan indicadores que combinan:
- buen resultado frente a objetivo,
- buen resultado frente a benchmark,
- y buena relación frente al máximo.

### Debilidades
Se priorizan indicadores con:
- mayor gap frente a objetivo,
- mayor gap frente a benchmark,
- y score bruto relativamente bajo.

---

## Cómo propone formación

La app intenta asignar cada curso a una de estas áreas:

- Nutrición
- Patología
- Manejo
- Herramientas
- General

Después filtra y prioriza:
- por especie,
- por subespecie,
- por área,
- por contenido,
- por coincidencia con el tronco a desarrollar,
- y por rol dentro del plan:
  - **formación base** para cerrar brechas,
  - **formación de especialización** para reforzar fortalezas.

Además, puede activar una casilla de la interfaz:

**“Refuerza con búsqueda web externa”**

Con esa opción activa, la app intenta añadir propuestas externas con:
- título,
- enlace,
- fuente,
- y una clasificación orientativa por área.

---

## Búsqueda de cursos: cómo funciona esta versión

Esta versión mejora la detección de cursos respecto a las anteriores porque:

- lee **varios catálogos a la vez**;
- detecta catálogos automáticamente desde el repositorio;
- soporta **estructuras de Excel heterogéneas**;
- intenta inferir especie/subespecie incluso cuando no están perfectamente normalizadas;
- clasifica el contenido en áreas mediante palabras clave;
- permite complementar la propuesta con búsqueda externa.

La interfaz muestra además un bloque de **resumen de catálogos cargados**, para revisar si la app realmente ha detectado cursos.

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

Si el nombre no es exactamente `requirements.txt`, Streamlit puede no instalar librerías como `openpyxl`, y aparecerá un error como:

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
```

---

## Cómo usar la app

1. Sube el **assessment del técnico**.
2. Sube **uno o varios catálogos** o deja que la app detecte los del repositorio.
3. Ajusta, si quieres, los **umbrales de perfil** en la barra lateral.
4. Decide cuántos cursos por área quieres proponer.
5. Decide cuántas propuestas externas por área quieres proponer.
6. Activa o desactiva la casilla **“Refuerza con búsqueda web externa”**.
7. Revisa:
   - el resumen ejecutivo,
   - la lectura por troncos,
   - fortalezas,
   - debilidades,
   - plan de formación.
8. Añade preguntas desde el desplegable.
9. Edita las respuestas.
10. Marca cuáles se incorporan al informe.
11. Edita el informe final si lo deseas.
12. Descárgalo en:
   - TXT
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
