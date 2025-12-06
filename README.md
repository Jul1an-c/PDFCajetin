# PDFCajetín – Cajetín automático

![Python](https://img.shields.io/badge/Python-3.14%2B-3776AB?style=flat&logo=python&logoColor=white)
![Flet](https://img.shields.io/badge/Flet-Interfaz-FF6F00?style=flat&logo=flet&logoColor=white)
![Estado](https://img.shields.io/badge/Estado-Listo%20para%20usar-success)
![Uso](https://img.shields.io/badge/Uso-Tareas%20%20·%20Mate-9cf)
![Facilidad](https://img.shields.io/badge/Facilidad-100%25-brightgreen)

---

## El problema de siempre

En tareas nos piden que **la primera hoja** lleve sí o un cajetín con:

- Nombre  
- Nombre completo  
- Curso 
- Fecha  

Y claro, se hacen los ejercicios, pero después se pierden **15 minutos** ajustado el cajetín en Word, alineando textos, exportando el PDF, o reimprimiendo porque quedó mal…

---

## La solución 

**PDFCajetín**: una app tan simple como útil, hecha para colocar el cajetín oficial sobre el PDF **en 5 segundos**, sin pensar y sin sufrir.

---

## 📝 ¿Cómo queda tu hoja?

- Cajetín arriba, centrado y con tus datos (actualízalos en el archivo Word antes de usar la app, ya que el programa no los modifica).  
- Tus ejercicios justo debajo, sin recortes ni deformaciones.  
- El resto de páginas del PDF quedan intactas.  

---

# 🔧 Instalación  
Elegí el modo que más te convenga:

---

## Opción A — Instalación rápida (PDFCajetin.exe)

**La forma más fácil y rápida**

### 1. Ir a la sección de Releases  
👉 [https://github.com/Jul1an-c/PDFCajetin/releases]

### 2. Descargar el archivo  
- `PDFCajetin.exe` 

### 3. Ejecutar  
Doble clic → **se abrirá una terminal + la ventana de la app**.

**La terminal es 100 % normal y necesaria**  
Es la consola de Python que se usa en segundo plano para convertir el Word a PDF.  
No la cierres, se cerrará sola cuando termines o cierres la app.

¡Listo! Ya podés usar el programa sin instalar nada más.


---

## 🟩 **Opción B — Instalación con código (para usuarios curiosos o programadores)**

### **1. Clonar el repositorio**
```bash
git clone https://github.com/Jul1an-c/PDFCajetin
```

### **2. Entrar a la carpeta**

```bash
cd pdfcajetin
```

### **3. (Opcional, pero recomendado) Crear un entorno virtual**

```bash
python -m venv venv
```

### **4. Instalar dependencias**
```bash
pip install -r requirements.txt
```

### **5. Ejecutar la app**
```bash
python main.py
```

### **Espero y te sirva :)**
