import sys
import os

# PARCHE PARA PYINSTALLER SIN CONSOLA: Evita que la app crashee al inicio
if sys.stdout is None or sys.stderr is None:
    sys.stdout = open(os.devnull, 'w')
    sys.stderr = open(os.devnull, 'w')

import customtkinter as ctk
from tkinter import messagebox
import threading
from PIL import Image
import smtplib
from email.message import EmailMessage
import multiprocessing
import Main

# colores, use como paleta los colores correspondientes al logo de ing. biomédica PUCP
COLOR_FONDO = "#8FB73F"
COLOR_ACENTO_BLUE = "#011A3E"
COLOR_TEXTO = COLOR_ACENTO_BLUE
COLOR_INPUT_TEXTO = "#4CA6FF"

ctk.set_appearance_mode("dark")

# Esta parte redirige la consola o cmd a la app, para que se pueda ver el proceso
# sinceramente queda muy bonito xd
class RedirigirConsola:
    def __init__(self, textbox_func):
        self.textbox_func = textbox_func

    def write(self, texto):
        if texto:
            self.textbox_func(texto)

    def flush(self):
        pass

# MAIN DASHBOARD
class AppPrincipal(ctk.CTk):
    def __init__(self, rol_usuario):
        super().__init__(fg_color=COLOR_FONDO)
        #Para implementar el servicio de correo, se uso el modo app - password, el cuál permite acceder a un gmail vía codigo
        self.correo_remitente = "reportesmantenimiento96@gmail.com"
        #La contraseña esta hardcodeada, confiare en que no accederás al gmail ....
        self.password_app = "vrofliuaxzycvvkd"

        # Rutas dinámicas separadas para evitar conflictos con el excel (honestamente fue sugerencia de claude)
        if getattr(sys, 'frozen', False):
            directorio_base_externo = os.path.dirname(sys.executable)
            directorio_recursos_internos = sys._MEIPASS
        else:
            directorio_base_externo = os.path.dirname(os.path.abspath(__file__))
            directorio_recursos_internos = directorio_base_externo

        self.ruta_excel = os.path.join(directorio_base_externo, "ArchivosGenerados", "INFORME_COMPLETO_ESA615.xlsx")
        self.ruta_logo = os.path.join(directorio_recursos_internos, "Captura de pantalla 2026-03-07 203532.png")

        self.title("Sistema de seguridad eléctrica automatizado")
        self.geometry("650x850")
        self.resizable(False, False)

        # Encabezado
        self.label_bienvenida = ctk.CTkLabel(
            self, text=f"Panel de Control ({rol_usuario})",
            font=("Roboto", 22, "bold"), text_color=COLOR_TEXTO
        )
        self.label_bienvenida.pack(pady=(20, 10))

        # Botón Principal
        self.boton_iniciar = ctk.CTkButton(
            self, text="INICIAR REVISIÓN", font=("Roboto", 16, "bold"),
            fg_color=COLOR_ACENTO_BLUE, hover_color="#002244", text_color=COLOR_FONDO,
            height=50, width=300, command=self.iniciar_hilo_automatizacion
        )
        self.boton_iniciar.pack(pady=10)

        # Barra de progreso
        self.progressbar = ctk.CTkProgressBar(
            self, width=500, mode="indeterminado", progress_color=COLOR_ACENTO_BLUE
        )
        self.progressbar.pack(pady=10)
        self.progressbar.set(0)

        # Consola Integrada
        self.label_consola = ctk.CTkLabel(
            self, text="Monitor de Actividad (En vivo):",
            font=("Roboto", 12, "bold"), text_color=COLOR_TEXTO
        )
        self.label_consola.pack(anchor="w", padx=75, pady=(5, 0))

        # Cuadro de texto de la consola
        self.consola_texto = ctk.CTkTextbox(
            self, width=500, height=200, fg_color="black", text_color="#4CA6FF",
            font=("Consolas", 12)
        )
        self.consola_texto.pack(pady=5)
        self.consola_texto.configure(state="disabled")

        # Botones de Acción
        self.boton_abrir_excel = ctk.CTkButton(
            self, text="ABRIR REPORTE EXCEL", font=("Roboto", 14, "bold"),
            fg_color=COLOR_ACENTO_BLUE, hover_color="#002244", text_color=COLOR_FONDO,
            height=40, width=250, command=self.abrir_archivo_excel, state="disabled"
        )
        self.boton_abrir_excel.pack(pady=(15, 5))

        self.boton_enviar_correo = ctk.CTkButton(
            self, text="ENVIAR POR CORREO", font=("Roboto", 14, "bold"),
            fg_color=COLOR_ACENTO_BLUE, hover_color="#002244", text_color=COLOR_FONDO,
            height=40, width=250, command=self.solicitar_correo_destinatario, state="disabled"
        )
        self.boton_enviar_correo.pack(pady=(5, 10))

        # Cargar logo si existe
        if os.path.exists(self.ruta_logo):
            self.load_and_place_logo()

    def load_and_place_logo(self):
        imagen_pil = Image.open(self.ruta_logo)
        ancho_original, alto_original = imagen_pil.size
        NUEVO_ANCHO = 200
        NUEVO_ALTO = int(alto_original * (NUEVO_ANCHO / ancho_original))

        self.logo_image = ctk.CTkImage(light_image=imagen_pil, size=(NUEVO_ANCHO, NUEVO_ALTO))
        self.label_logo = ctk.CTkLabel(self, text="", image=self.logo_image, fg_color="transparent")
        self.label_logo.place(relx=0.5, rely=0.96, anchor="s")

    def escribir_en_consola(self, texto):
        self.consola_texto.configure(state="normal")
        self.consola_texto.insert("end", texto)
        self.consola_texto.see("end")
        self.consola_texto.configure(state="disabled")

    def despachar_a_consola_seguro(self, texto):
        self.after(0, self.escribir_en_consola, texto)

    def iniciar_hilo_automatizacion(self):
        self.boton_iniciar.configure(state="disabled", text="EJECUTANDO...")
        self.boton_abrir_excel.configure(state="disabled")
        self.boton_enviar_correo.configure(state="disabled")
        self.progressbar.start()

        self.consola_texto.configure(state="normal")
        self.consola_texto.delete("1.0", "end")
        self.consola_texto.configure(state="disabled")

        self.escribir_en_consola(">>> Iniciando comunicación con Fluke ESA615...\n")

        hilo = threading.Thread(target=self.ejecutar_script_real)
        hilo.start()

    def ejecutar_script_real(self):
        consola_original = sys.stdout
        sys.stdout = RedirigirConsola(self.despachar_a_consola_seguro)

        try:
            exito = Main.iniciar_orquestador()
        except Exception as e:
            self.despachar_a_consola_seguro(f"\n>>> ERROR FATAL: {e}\n")
            exito = False
        finally:
            sys.stdout = consola_original
            self.after(0, self.finalizar_interfaz, exito)

    def finalizar_interfaz(self, exito):
        self.progressbar.stop()
        self.progressbar.set(0)
        self.boton_iniciar.configure(state="normal", text="INICIAR REVISIÓN")

        if exito:
            self.escribir_en_consola("\n>>> REVISIÓN FINALIZADA CORRECTAMENTE.\n")
            self.boton_abrir_excel.configure(state="normal")
            self.boton_enviar_correo.configure(state="normal")
            messagebox.showinfo("Éxito", "Proceso finalizado con éxito.")
        else:
            self.escribir_en_consola("\n>>> ERROR O CANCELACIÓN DURANTE EL PROCESO.\n")

    def abrir_archivo_excel(self):
        if os.path.exists(self.ruta_excel):
            try:
                os.startfile(self.ruta_excel)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showerror("Error", f"No se encontró el Excel en:\n{self.ruta_excel}")

    def solicitar_correo_destinatario(self):
        dialog = ctk.CTkInputDialog(text="Ingrese el correo del destinatario:", title="Enviar Reporte")
        destinatario = dialog.get_input()

        if destinatario and "@" in destinatario and "." in destinatario:
            self.escribir_en_consola(f"\n>>> Preparando correo para {destinatario}...\n")
            self.boton_enviar_correo.configure(state="disabled", text="ENVIANDO...")
            threading.Thread(target=self.enviar_correo_hilo, args=(destinatario,)).start()
        elif destinatario:
            messagebox.showwarning("Advertencia", "Ingrese un correo válido.")

    def enviar_correo_hilo(self, destinatario):
        try:
            if not os.path.exists(self.ruta_excel):
                raise FileNotFoundError("No se encontró el archivo Excel.")

            msg = EmailMessage()
            msg['Subject'] = "Reporte de Seguridad Eléctrica - Fluke ESA615"
            msg['From'] = self.correo_remitente
            msg['To'] = destinatario
            msg.set_content("Adjunto el reporte generado automáticamente.\n\nSaludos.")

            with open(self.ruta_excel, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application',
                                   subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                   filename=os.path.basename(self.ruta_excel))

            self.despachar_a_consola_seguro(">>> Conectando con servidor de Gmail...\n")

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(self.correo_remitente, self.password_app)
                smtp.send_message(msg)

            self.after(0, self.correo_exito, destinatario)
        except Exception as e:
            self.after(0, self.correo_error, str(e))

    def correo_exito(self, destinatario):
        self.escribir_en_consola(f">>> ¡Correo enviado exitosamente a {destinatario}!\n")
        self.boton_enviar_correo.configure(state="normal", text="ENVIAR POR CORREO")
        messagebox.showinfo("Correo Enviado", f"Reporte enviado a:\n{destinatario}")

    def correo_error(self, error):
        self.escribir_en_consola(f">>> ERROR al enviar correo: {error}\n")
        self.boton_enviar_correo.configure(state="normal", text="ENVIAR POR CORREO")
        messagebox.showerror("Error", f"No se pudo enviar:\n{error}")

# =======================================================
# PANTALLA DE LOGIN (evitar manipulación)
# =======================================================
class AppLogin(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=COLOR_FONDO)
        self.title("Sistema de seguridad eléctrica automatizado")
        self.geometry("400x600")
        self.resizable(False, False)

        if getattr(sys, 'frozen', False):
            directorio_recursos_internos = sys._MEIPASS
        else:
            directorio_recursos_internos = os.path.dirname(os.path.abspath(__file__))

        self.ruta_logo = os.path.join(directorio_recursos_internos, "Captura de pantalla 2026-03-07 203532.png")

        self.label_titulo = ctk.CTkLabel(self, text="Bienvenido", font=("Roboto", 24, "bold"), text_color=COLOR_TEXTO)
        self.label_titulo.pack(pady=(20, 10))

        self.label_subtitulo = ctk.CTkLabel(self, text="Sistema automatizado", font=("Roboto", 14),
                                            text_color=COLOR_TEXTO)
        self.label_subtitulo.pack(pady=(0, 20))

        self.entry_usuario = ctk.CTkEntry(self, placeholder_text="Usuario", width=250, height=40,
                                          border_color=COLOR_ACENTO_BLUE, text_color=COLOR_INPUT_TEXTO)
        self.entry_usuario.pack(pady=10)

        self.entry_password = ctk.CTkEntry(self, placeholder_text="Contraseña", width=250, height=40, show="*",
                                           border_color=COLOR_ACENTO_BLUE, text_color=COLOR_INPUT_TEXTO)
        self.entry_password.pack(pady=10)

        self.boton_login = ctk.CTkButton(self, text="INGRESAR", font=("Roboto", 14, "bold"), fg_color=COLOR_ACENTO_BLUE,
                                         hover_color="#002244", text_color=COLOR_FONDO, width=250, height=40,
                                         command=self.verificar_login)
        self.boton_login.pack(pady=(30, 10))

        if os.path.exists(self.ruta_logo):
            self.load_and_place_logo()

    def load_and_place_logo(self):
        imagen_pil = Image.open(self.ruta_logo)
        ancho_original, alto_original = imagen_pil.size
        NUEVO_ANCHO = 150
        NUEVO_ALTO = int(alto_original * (NUEVO_ANCHO / ancho_original))

        self.logo_image = ctk.CTkImage(light_image=imagen_pil, size=(NUEVO_ANCHO, NUEVO_ALTO))
        self.label_logo = ctk.CTkLabel(self, text="", image=self.logo_image, fg_color="transparent")
        self.label_logo.place(relx=0.5, rely=0.96, anchor="s")

    def verificar_login(self):
        usuario = self.entry_usuario.get()
        password = self.entry_password.get()

        # Validación directa(se penso suar SQlite pero complicaba el .exe)
        if usuario == "admin" and password == "admin123":
            self.destroy()
            app_principal = AppPrincipal("Administrador")
            app_principal.mainloop()
        else:
            messagebox.showerror("Error", "Usuario o contraseña incorrectos.")

if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = AppLogin()
    app.mainloop()