# Senior Technical Consultant Selector

Aplicación en **Streamlit** para analizar evaluaciones técnicas individuales, construir un **ranking comparativo de candidatos** y proponer, con criterios transparentes, qué técnico está más preparado para asumir un rol de **consultor senior**.

La app está pensada para equipos técnico-comerciales veterinarios y trabaja a partir de archivos de assessment en **Excel** (`.xlsm` o `.xlsx`).

---

## Qué hace la app

La aplicación permite:

- subir entre **2 y 10 evaluaciones** individuales;
- leer los **25 indicadores** y los **4 troncos** del assessment;
- comparar cada técnico contra:
  - **objetivo**,
  - **máximo**,
  - **media BBDD / benchmark**;
- calcular un **Senior Score** con lógica multicriterio;
- ordenar a todos los candidatos en un **ranking de mejor a peor aptitud**;
- proponer un candidato para **consultor senior**;
- justificar técnica y objetivamente la elección;
- explicar también por qué los demás candidatos **no han sido elegidos todavía**;
- mostrar un **gráfico de barras** y un **gráfico de araña** comparativo;
- responder preguntas frecuentes del director técnico mediante un bloque de **preguntas-respuestas editables**;
- generar un **informe final editable** y exportable.

---

## Lógica de evaluación

La elección del senior **no se basa solo en la nota global**.  
La app utiliza un ranking compuesto por cinco dimensiones:

- **35%** Rendimiento técnico global
- **20%** Equilibrio entre los 4 troncos
- **20%** Indicadores críticos para senior
- **15%** Capacidad de transferencia y formación
- **10%** Ventaja respecto a la media del equipo

Además, la app trabaja con **3 niveles de exigencia**:

- **Nivel 3**: exigencia alta
- **Nivel 2**: exigencia media
- **Nivel 1**: exigencia flexible

Si con el nivel elegido no aparece ningún candidato apto, la app puede proponer automáticamente el mejor perfil disponible en el siguiente nivel inferior.

---

## Entradas necesarias

La app espera archivos de evaluación individuales con una estructura equivalente a la herramienta interna de assessment.

### Archivos admitidos
- `.xlsm`
- `.xlsx`

### Contenido esperado del Excel
- Hoja `REFERENCIAS`
- Hoja `EVALUACION`

La lógica del parser espera, como mínimo:

- 25 indicadores técnicos
- 4 troncos / áreas
- pesos
- objetivo / referencia
- máximo
- benchmark / BBDD
- score del técnico por indicador

> Recomendación: abre y guarda previamente cada Excel en Microsoft Excel antes de subirlo, para asegurarte de que las fórmulas estén recalculadas.

---

## Salidas de la app

La aplicación genera:

### 1. Conclusión ejecutiva
- candidato propuesto para senior, o
- mensaje indicando que **todavía no hay un senior claro**

### 2. Ranking completo
Lista de todos los candidatos ordenados de:
**mejor aptitud senior → menor aptitud senior**

### 3. Argumentos técnicos
Para cada candidato:
- fortalezas objetivas,
- gaps,
- motivos por los que sí o no es prioritario como senior.

### 4. Gráficos
- gráfico de barras comparativo;
- gráfico de araña con los 5 componentes del ranking.

### 5. Informe final
- informe breve global;
- bloque de preguntas y respuestas editables;
- exportación a Excel y JSON.

---

## Preguntas frecuentes integradas

La app incluye un desplegable con **20 preguntas posibles** que un director técnico puede querer resolver, por ejemplo:

- ¿El evaluado podría ser formado en áreas muy específicas o exigentes?
- ¿Tiene potencial real para convertirse en consultor senior?
- ¿Qué gap le separa con más claridad del perfil senior?
- ¿Qué combinación de formación interna y externa parece más eficiente?
- ¿Qué argumentos objetivos justifican que no sea el elegido ahora mismo?

Cada pregunta:
- se puede añadir al análisis;
- genera una respuesta inicial automática;
- puede ser **editada manualmente**;
- y puede **incluirse o no** en el informe final.

---

## Estructura recomendada del repositorio

```text
/
├── main.py
├── requirements.txt
└── README.md
```

---

## Despliegue en GitHub + Streamlit Community Cloud

### 1. Subir archivos a GitHub
En el repositorio deja, al menos:

- `main.py`
- `requirements.txt`
- `README.md`

### 2. Crear la app en Streamlit
En Streamlit Community Cloud:

1. Conecta tu cuenta de GitHub.
2. Pulsa **Create app**.
3. Selecciona:
   - repositorio
   - rama
   - archivo principal: `main.py`
4. Pulsa **Deploy**.

---

## Ejecución local

Si quieres probarla en local:

```bash
pip install -r requirements.txt
streamlit run main.py
```

---

## Cómo usar la app

1. Sube entre **2 y 10 archivos** de evaluación.
2. Selecciona el **nivel de exigencia**.
3. Revisa:
   - la conclusión ejecutiva,
   - el ranking,
   - los argumentos por candidato,
   - los gráficos.
4. Añade preguntas desde el desplegable.
5. Edita las respuestas si lo consideras necesario.
6. Decide qué preguntas incorporar al informe.
7. Descarga el resultado final.

---

## Botones disponibles

### Refrescar / recalcular
Recalcula el análisis con los parámetros actuales.

### Borrar evaluación cargada
Limpia el estado de la app y permite empezar desde cero.

---

## Limitaciones conocidas

- La app depende de que el Excel tenga una estructura coherente con la herramienta interna.
- Si las fórmulas del archivo no están actualizadas, los resultados pueden ser incorrectos.
- El parser actual está orientado a una plantilla concreta; si cambia la estructura del Excel, puede requerir ajustes.
- La app genera argumentos automáticos, pero el usuario puede y debe revisarlos antes de usar el informe como documento definitivo.

---

## Recomendaciones de uso

- Utilizar la app como herramienta de **apoyo objetivo a la decisión**, no como único criterio.
- Revisar especialmente:
  - el tronco más débil,
  - los indicadores críticos,
  - la capacidad de transferencia/formación.
- Cuando no haya un senior claro, usar el informe para definir un **plan de desarrollo** y no forzar una elección.

---

## Requisitos

Contenido típico de `requirements.txt`:

```txt
streamlit>=1.44.0
pandas>=2.2.0
openpyxl>=3.1.0
xlsxwriter>=3.2.0
plotly>=5.20.0
```

---

## Objetivo de negocio

La finalidad de la app es ayudar a un director técnico a tomar una decisión **rigurosa, defendible y justa** sobre:

- quién puede asumir un rol de **consultor senior**,
- quién todavía no está preparado,
- y qué necesita cada técnico para evolucionar dentro del plan de carrera.

