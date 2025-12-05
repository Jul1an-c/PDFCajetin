import flet as ft
import fitz
import os
import time
import subprocess
import sys
import shutil
from docx2pdf import convert
from pathlib import Path
import psutil
import tempfile


def log(msg: str):
    print(f"[LOG] {msg}")


# === CARPETA DE SALIDA EN DOCUMENTOS ===
def obtener_carpeta_salida():
    documentos = Path.home() / "Documentos"
    carpeta_salida = documentos / "PDFCajetin"
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    return str(carpeta_salida)


CARPETA_SALIDA = obtener_carpeta_salida()


# === CARPETA TEMPORAL INTELIGENTE ===
def obtener_carpeta_temp():
    if getattr(sys, 'frozen', False):  # Estamos en .exe
        temp_dir = os.path.join(tempfile.gettempdir(), "PDFCajetin_temp")
    else:
        temp_dir = os.path.join(os.path.dirname(__file__), "storage", "temp_cajetin")
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir


CARPETA_TEMP = obtener_carpeta_temp()
log(f"Carpeta temporal en uso: {CARPETA_TEMP}")  # Para que veas dónde está trabajando


# === LIMPIEZA TOTAL AL SALIR (funciona siempre, incluso cerrando con la X) ===
def limpiar_al_salir():
    try:
        # Mata procesos de Word colgados
        if sys.platform == "win32":
            os.system("taskkill /f /im WINWORD.EXE >nul 2>&1")

        # Borra la carpeta temporal (sea la del proyecto o la del sistema)
        if os.path.exists(CARPETA_TEMP):
            shutil.rmtree(CARPETA_TEMP, ignore_errors=True)
            log(f"Carpeta temporal eliminada: {CARPETA_TEMP}")
    except Exception as e:
        log(f"Error al limpiar: {e}")
    finally:
        # Salida forzada limpia
        try:
            os._exit(0)
        except:
            pass


# Registramos la limpieza de dos formas (una siempre gana)
import atexit
atexit.register(limpiar_al_salir)


# === FUNCIONES DE ABRIR ===
def esperar_liberacion(ruta: str, timeout: int = 10) -> bool:
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            with open(ruta, "rb"):
                return True
        except PermissionError:
            time.sleep(0.2)
    return False


def abrir_carpeta(ruta: str):
    if os.path.exists(ruta):
        if sys.platform == "win32":
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.run(["open", ruta])
        else:
            subprocess.run(["xdg-open", ruta])


def abrir_archivo(ruta: str):
    if os.path.exists(ruta):
        if sys.platform == "win32":
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.run(["open", ruta])
        else:
            subprocess.run(["xdg-open", ruta])


# === MAIN ===
def main(page: ft.Page):
    page.title = "PDFCajetín"
    page.window_width = 750
    page.window_height = 620
    page.padding = 30
    page.scroll = "adaptive"
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    pdf_path = ft.Text("Ningún PDF seleccionado", size=13, italic=True)
    docx_path = ft.Text("Ningún cajetín seleccionado", size=13, italic=True)
    status = ft.Text("", size=15, weight="bold", text_align="center")
    progress = ft.ProgressBar(visible=False, width=500)

    btn_abrir_carpeta = ft.TextButton(
        "Abrir carpeta de salida",
        icon=ft.Icons.FOLDER_OPEN,
        visible=False,
        on_click=lambda _: abrir_carpeta(CARPETA_SALIDA)
    )
    btn_abrir_pdf = ft.TextButton(
        "Abrir PDF generado",
        icon=ft.Icons.PICTURE_AS_PDF,
        visible=False
    )

    ultimo_pdf = [""]

    picker_pdf = ft.FilePicker(on_result=lambda e: sel(e, pdf_path))
    picker_docx = ft.FilePicker(on_result=lambda e: sel(e, docx_path))
    page.overlay.extend([picker_pdf, picker_docx])

    def sel(e: ft.FilePickerResultEvent, txt: ft.Text):
        if e.files:
            txt.value = e.files[0].path
            txt.italic = False
            page.update()

    def procesar(e):
        if "Ningún" in pdf_path.value or "Ningún" in docx_path.value:
            status.value = "Selecciona ambos archivos"
            status.color = ft.Colors.RED
            page.update()
            return

        progress.visible = True
        status.value = "Procesando..."
        status.color = None
        btn_abrir_carpeta.visible = False
        btn_abrir_pdf.visible = False
        page.update()

        doc_esc = doc_caj = nuevo = None
        cajetin_pdf = os.path.join(CARPETA_TEMP, "cajetin.pdf")

        try:
            if os.path.exists(cajetin_pdf):
                try: os.remove(cajetin_pdf)
                except: pass

            convert(docx_path.value, cajetin_pdf)
            if not esperar_liberacion(cajetin_pdf):
                raise Exception("El cajetín no se liberó a tiempo")

            doc_esc = fitz.open(pdf_path.value)
            doc_caj = fitz.open(cajetin_pdf)
            nuevo = fitz.open()

            ANCHO = 600
            ALTO_OFICIO = 1008
            ALTO_CARTA = 755

            pag1 = nuevo.new_page(width=ANCHO, height=ALTO_OFICIO)
            pag1.show_pdf_page(fitz.Rect(0, 0, ANCHO, ALTO_OFICIO), doc_caj, 0)
            pag1.show_pdf_page(fitz.Rect(0, ALTO_OFICIO - ALTO_CARTA, ANCHO, ALTO_OFICIO), doc_esc, 0)

            for i in range(1, len(doc_esc)):
                nuevo.insert_pdf(doc_esc, from_page=i, to_page=i)

            nombre_archivo = f"{os.path.splitext(os.path.basename(pdf_path.value))[0]}_CON_CAJETIN.pdf"
            ultimo_pdf[0] = os.path.join(CARPETA_SALIDA, nombre_archivo)
            nuevo.save(ultimo_pdf[0], garbage=4, deflate=True, clean=True)

            status.value = "¡Perfecto! PDF generado en Documentos/PDFCajetin"
            status.color = ft.Colors.GREEN_800

            btn_abrir_carpeta.visible = True
            btn_abrir_pdf.visible = True
            btn_abrir_pdf.on_click = lambda _: abrir_archivo(ultimo_pdf[0])

        except Exception as ex:
            status.value = f"Error: {ex}"
            status.color = ft.Colors.RED
            log(f"ERROR: {ex}")

        finally:
            for doc in (doc_esc, doc_caj, nuevo):
                if doc:
                    try: doc.close()
                    except: pass
            progress.visible = False
            page.update()

    # === INTERFAZ ===
    page.add(
        ft.Column([
            ft.Text("PDFCajetín", size=26, weight="bold", text_align="center"),
            ft.Text("Agregar cajetín en la primera página del PDF", size=14, text_align="center"),
            ft.Divider(height=30),

            ft.Container(
                content=ft.Column([
                    ft.Row([ft.Icon(ft.Icons.PICTURE_AS_PDF_OUTLINED, size=28),
                            ft.Text("PDF Escaneado", weight="bold", size=15)], alignment="center"),
                    ft.Container(pdf_path, padding=ft.padding.only(left=35, right=35), alignment=ft.alignment.center),
                    ft.ElevatedButton("Seleccionar PDF", icon=ft.Icons.FOLDER_OPEN,
                                      on_click=lambda _: picker_pdf.pick_files(allowed_extensions=["pdf"]), width=260)
                ], horizontal_alignment="center", spacing=10),
                padding=20, border=ft.border.all(1, ft.Colors.OUTLINE), border_radius=10, width=600
            ),

            ft.Container(height=16),

            ft.Container(
                content=ft.Column([
                    ft.Row([ft.Icon(ft.Icons.DESCRIPTION_OUTLINED, size=28),
                            ft.Text("Cajetín (Word)", weight="bold", size=15)], alignment="center"),
                    ft.Container(docx_path, padding=ft.padding.only(left=35, right=35), alignment=ft.alignment.center),
                    ft.ElevatedButton("Seleccionar cajetín", icon=ft.Icons.FOLDER_OPEN,
                                      on_click=lambda _: picker_docx.pick_files(allowed_extensions=["docx","doc"]), width=260)
                ], horizontal_alignment="center", spacing=10),
                padding=20, border=ft.border.all(1, ft.Colors.OUTLINE), border_radius=10, width=600
            ),

            ft.Container(height=10),
            progress,
            status,
            ft.ElevatedButton("GENERAR PDF FINAL", icon=ft.Icons.DONE_ALL, on_click=procesar, height=52, width=320),
            ft.Row([btn_abrir_pdf, btn_abrir_carpeta], alignment=ft.MainAxisAlignment.CENTER, spacing=20)

        ], spacing=12, horizontal_alignment="center", scroll=ft.ScrollMode.AUTO)
    )


# === ARRANQUE SEGURO CON LIMPIEZA GARANTIZADA ===
if __name__ == "__main__":
    try:
        ft.app(target=main)
    finally:
        # Esto se ejecuta SIEMPRE, incluso si cierras con la X
        limpiar_al_salir()