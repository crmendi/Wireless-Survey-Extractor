# Wireless Survey Extractor - Desktop App

Una aplicación de escritorio rápida y nativa construida en Python (Tkinter) para un mejor manejo y análisis de los archivos de diseños inalámbricos `.esx` (generados por software de levantamiento como Ekahau).

> **⚠️ AVISO LEGAL DE MARCA:** 
> "Ekahau" y "Ekahau Pro" son marcas comerciales registradas de Ekahau, Inc. 
> Este proyecto **NO está afiliado, respaldado, soportado ni patrocinado por Ekahau.** Esta es una herramienta independiente y de código abierto desarrollada por terceros para habilitar flujos de trabajo personalizados y extracciones analíticas leyendo la estructura de archivos `.esx`.

> **OS SOPORTADO:** Esta versión en código abierto y el binario ejecutable están construidos **ÚNICAMENTE para sistemas operativos Windows**.

---

## 🎯 ¿Qué hace esta herramienta?

Los archivos `.esx` son esencialmente contenedores ZIP complejos que guardan planos, coordenadas JSON, metadatos y modelos de Access Points (APs). Esta herramienta interactúa con esa estructura sin necesidad de abrir programas pesados ni contar con licencias activas para:
1. **Contar y Cuantificar:** Listado tabular agrupado por modelo, archivos y pisos.
2. **Post-procesamiento Visual:** Colocación automática de burbujas/marcas precisas (APs) sobre los planos en crudo.
3. **Generación de Reportes Dinámicos:** Exportación a un documento Word (DOCX) limpio con todos los planos demarcados, tablas de resúmenes totales y gráficos de barras analíticos.
4. **Exportar a CSV y PDFs:** Exportación de tablas planas (CSV) o unión masiva de PDFs.

---

## 📖 Instrucciones de Uso (Paso a Paso)

### 1. Requisitos Previos y Preparación
- Asegúrate de tener tu archivo `.esx` finalizado.
- **🚨 ADVERTENCIA CLAVE - SÓLO DISEÑOS PREDICTIVOS:** Esta herramienta lee las ubicaciones estáticas de los APs. Por lo tanto, tu archivo `.esx` debe contener **ÚNICAMENTE APs Simulados (Levantamientos Predictivos)**. Si tu archivo incluye "APs Detectados" capturados automáticamente en un barrido físico de *Site Survey Activo*, el algoritmo generará superposiciones, lecturas erráticas o simplemente no funcionará correctamente, ya que los APs detectados tienen otra estructura de datos en el JSON de base. 

### 2. Carga de Archivos
- Abre el programa. Encontrarás la interfaz principal dividida en Panel de Filtros y Tabla de Datos.
- Ve a **"Archivo" > "Cargar archivo(s) .esx"**. 
- Puedes seleccionar *Múltiples* archivos simultáneamente. La aplicación los abrirá silenciosamente, leerá sus bases de datos internas JSON e iterará sobre cada iteración inalámbrica mapeando Coordenadas (X, Y) vs Pisos.
- Tras 1-2 segundos, verás en la parte inferior un texto confirmando cuáles archivos están montados en memoria y la tabla se llenará.

### 3. Filtros y Búsquedas Dinámicas
En la parte media, verás campos desplegables (Comboboxes). Úsalos para segmentar masivamente tus proyectos si subiste varios a la vez:
- **Archivo:** Muestra el listado de archivos individuales. Útil para aislar conteos de un solo edificio si subiste un Campus entero.
- **Modelo AP:** Muestra únicamente las frecuencias o familias de equipos (Ej. *Cisco Catalyst 9120AXI*). Selecciona un modelo específico para ver en qué pisos exactos de qué archivos fue usado.
- **Piso:** Aislar los datos para un nivel arquitectónico específico (Ej. "Piso 2", "Mezanine").

### 4. Personalización del Graficado (¡Vital para los Planos!)
Antes de generar entregables de Word o Imágenes, debes definir **cómo se verán visualmente** los Access Points en los mapas.
- Ve a **"Herramientas" > "Configuración..."**
- Se abrirá un panel con un **lienzo de prueba interactivo**. En la izquierda podrás modificar:
    - **Radio y Color del círculo:** Es el tamaño de la "burbuja" del AP en el mapa. Ajustalo dependiendo de qué tan grande o pequeño sea tu plano original.
    - **Fuentes, Textos y Contornos:** Configura el número textual del AP (Ej. 'AP-01'). Agregar contornos (Outline) suele hacer que el texto sea legible sin importar si el fondo del plano es oscuro o claro.
    - **Estilos de Notas:** Lo mismo para las notas de texto que hayas dejado en tu proyecto.
- Presiona **Aplicar** y fíjate en la vista previa derecha. Cierra la ventana cuando estés a gusto.

### 5. Configurar Portada Corporativa (Opcional)
Si vas a generar un Reporte en Word, quizás quieras el logo de tu compañía en la primera página:
- Ve a **"Herramientas" > "Importar Imagen para Informe..."** y selecciona tu logo (PNG o JPG).

### 6. Exportación de Entregables
Ahora el sistema está listo, ve al menú **Exportar**:
- **Generar Reporte Word:** La joya de la corona. Te pedirá dónde quieres guardar el `.docx`. La aplicación empezará a abrir todos los planos internamente, calculará algoritmos de Anti-Colisión (`bboxes_overlap`) para que ninguna burbuja de AP tape un texto existente de Ekahau, armará gráficos estadísticos usando Matplotlib, cruzará modelos vs tablas, y te escupirá un Documento de estilo ejecutivo. (*Nota: Puede tardar unos segundos dependiendo del peso volumétrico de las imágenes*).
- **Exportar Imágenes con APs:** Si no requieres Word, y solo quieres las fotos de los mapas puras (`.png`) para anexarlas a tu propio PDF o plantilla. Requerirá una carpeta de destino.
- **Exportar a CSV:** Exportará exactamente lo que estás viendo en tu tabla frontal activa a un excel, con los filtros aplicados.

### 7. Trabajo Incompleto (Guardar / Cargar Proyectos)
Si pasaste media hora perfeccionando el color exacto, el tamaño de la burbuja y los filtros aplicados y debes irte a almorzar, puedes guardar la "Sesión":
- Usa **Archivo > Guardar Proyecto**. 
- Esto generará un archivo ligero extensión `.aproj` que recuerda los archivos base `.esx` que subiste y cómo los dejaste parametrizados. 
- Puedes retomarlo después desde **"Archivo > Cargar Proyecto"**.

### 8. Utilidad Extra: Unir PDFs
- La aplicación incluye una utilidad ligera en **Herramientas > Unir PDFs** para consolidar rápidamente actas de entrega finales u otros informes externos sin necesidad de herramientas web de terceros.

---

## 🚀 Uso Directo en Windows (Recomendado)

Para usuarios que no son programadores y no desean lidiar con entornos virtuales de Python:

1. Ve a la ruta `Desktop_App/dist/wireless_survey_extractor.exe`. 
2. Haz doble clic sobre él (no requiere la terminal de comandos de CMD, iniciará la interfaz visual de una vez).
3. (Windows Defender podría mostrar una alerta de seguridad ("Pantalla Azul/SmartScreen") dado que es un ejecutable open-source sin firma digital comercial. Selecciona *"Más Información"* y luego *"Ejecutar de todas formas"*).

---

## 💻 Para Desarrolladores

Si deseas compilar la app tú mismo, ver el código o hacer "fork".

### Prerequisitos
* Python 3.9+
* Pip (Gestor de dependencias de Python)

### Instalación en Entorno
```bash
# Clona el repositorio y entra al mismo
cd Desktop_App/

# Instala todas las dependencias vitales (PIL, Docx, Matplotlib, Numpy)
pip install -r ../requirements.txt 
```

### Ejecutar Local
```bash
python wireless_survey_extractor.py
```

### Compilar un nuevo `.exe`
Si le haces modificaciones al código fuente y necesitas generar otro ejecutable:
```bash
pyinstaller --onefile --windowed --add-data "icon.ico;." --add-data "icon.png;." --icon="icon.ico" -n "wireless_survey_extractor" wireless_survey_extractor.py
```

---

## ⚖️ Licencia 
Este proyecto está bajo la Licencia **Apache 2.0**. Consulta el archivo `LICENSE` en la raíz del repositorio para más detalles legales. Se fomenta el uso libre, la modificación arquitectónica y el intercambio comercial o privado, **siempre y cuando se den creditos al autor (Christian Mendivelso)** o se incluyan enlaces a repeticiones de la licencia Apache en las distribuciones dadas.
