# 📋 Matriz de Gestión de Comunicaciones — AMPA 2026

Herramienta en Excel para planificar, seguir y reportar las actividades comunicacionales del equipo AMPA. Generada con Python + openpyxl y lista para editar directamente en Excel.

---

## 📁 Estructura del repositorio

```
ampa-matriz-comunicaciones/
│
├── Matriz_Comunicaciones_AMPA_2026_v2.xlsx   ← Archivo Excel principal (úsalo directamente)
│
├── generar_matriz.py       ← Script para regenerar el Excel desde cero
├── corregir_formulas.py    ← Script para corregir fórmulas en una versión existente
├── requirements.txt        ← Dependencias Python (solo openpyxl)
├── .gitignore
└── README.md               ← Esta guía
```

---

## 🚀 Uso rápido (sin instalar nada)

Si solo quieres usar el Excel:

1. Ve a la sección **Code** del repositorio en GitHub
2. Haz clic en `Matriz_Comunicaciones_AMPA_2026_v2.xlsx`
3. Haz clic en el botón **Download** (ícono de descarga ↓)
4. Abre el archivo en Excel o LibreOffice Calc

---

## 🗂️ Hojas del Excel y cómo usarlas

| Hoja | Para qué sirve |
|------|---------------|
| **Resumen General** | Vista ejecutiva de todos los proyectos. % Avance, N° Actividades e Indicadores se calculan solos. |
| **Actividades por Proyecto** | Lista operativa de todas las actividades. El % Avance se actualiza desde TAREAS. |
| **TAREAS** | Detalle de tareas por actividad con cronograma mensual. Aquí se actualiza el avance. |
| **Indicadores Comunicación** | Seguimiento de indicadores por trimestre para reportes a financiadores. |
| **Cronograma** | Vista mensual de todas las actividades del año. |
| **Comunicación Institucional** | Acciones de comunicación propias de AMPA (redes, boletines, web, eventos). |

### ➕ Cómo agregar un nuevo proyecto
1. Abre la hoja **Resumen General**
2. En la tabla de proyectos, escribe en la siguiente fila vacía disponible
3. Completa: Código (`PRY-XX`), Nombre, Financiador, Estado, Fechas, Responsable
4. El % Avance, N° Actividades y N° Indicadores se calcularán solos al agregar datos en las otras hojas

### ➕ Cómo agregar actividades al nuevo proyecto
1. Ve a la hoja **Actividades por Proyecto**
2. Escribe en la siguiente fila vacía de la tabla
3. Completa todos los campos. El % Avance se llenará automáticamente desde TAREAS

### ➕ Cómo agregar tareas
1. Ve a la hoja **TAREAS**
2. Escribe en la siguiente fila vacía
3. El campo **ID Actividad** es clave: usa el formato `PRY-XX_N` donde `N` es el número de fila de esa actividad en "Actividades por Proyecto"
   - Ejemplo: si "Producción documental" está en la fila 1 de datos → `PRY-01_1`
   - Si "Campaña RRSS" está en la fila 2 → `PRY-01_2`
4. Marca con **X** los meses de ejecución en el cronograma
5. Cuando cambies el Estado a **"Completada"**, el % Avance de la actividad se actualizará solo

---

## 🐍 Uso con Python (para regenerar o modificar el Excel)

### 1. Requisitos previos

- Python 3.8 o superior
- pip (gestor de paquetes de Python)

### 2. Clonar el repositorio

```bash
git clone https://github.com/TU_USUARIO/ampa-matriz-comunicaciones.git
cd ampa-matriz-comunicaciones
```

> Reemplaza `TU_USUARIO` con tu nombre de usuario de GitHub

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Regenerar el Excel desde cero

Si quieres crear un Excel limpio con los datos de ejemplo:

```bash
python generar_matriz.py
```

Esto creará (o sobreescribirá) `Matriz_Comunicaciones_AMPA_2026_v2.xlsx` en la misma carpeta.

### 5. Corregir fórmulas en un Excel existente

Si modificaste el Excel y las fórmulas dejaron de funcionar:

```bash
python corregir_formulas.py
```

> ⚠️ Este script lee `Matriz_Comunicaciones_AMPA_2026_v2.xlsx` de la misma carpeta y lo sobreescribe con las fórmulas corregidas.

---

## 🔄 Flujo de trabajo recomendado para el equipo

```
Cada semana:
  1. Abre el Excel descargado desde GitHub
  2. Actualiza el estado de las TAREAS (En proceso / Completada)
  3. Actualiza los avances en Indicadores (trimestral)
  4. Guarda y sube el Excel actualizado a GitHub

Para subir cambios a GitHub:
  1. Ve al repositorio en github.com
  2. Haz clic en el archivo .xlsx → Edit (ícono de lápiz) → no funciona para Excel
  3. Usa en cambio: Add file → Upload files → arrastra el archivo actualizado
  4. Escribe un mensaje de commit, ej: "Actualización semana 15 - avance PRY-09"
  5. Haz clic en Commit changes
```

---

## 📌 Notas importantes

- **No renombres el archivo** si usas los scripts Python, ya que estos buscan el nombre exacto `Matriz_Comunicaciones_AMPA_2026_v2.xlsx`
- **El ID de actividad en TAREAS** debe coincidir exactamente con el formato `PRY-XX_N` (mayúsculas, guion bajo) para que el vínculo funcione
- **Las tablas de Excel se expanden automáticamente** al escribir en la fila vacía justo debajo de la última fila de datos
- Si compartes el archivo por WhatsApp o correo, los destinatarios verán siempre los datos actuales pero no podrán subir cambios a GitHub

---

## 🛠️ Tecnologías usadas

- **Python 3** + [openpyxl](https://openpyxl.readthedocs.io/) para generación del Excel
- **Microsoft Excel** / **LibreOffice Calc** para edición
- **GitHub** para versionado y almacenamiento compartido

---

## 📬 Contacto

Equipo de Comunicaciones — AMPA  
Para dudas sobre el archivo, contactar a la Coordinación de Comunicaciones.
