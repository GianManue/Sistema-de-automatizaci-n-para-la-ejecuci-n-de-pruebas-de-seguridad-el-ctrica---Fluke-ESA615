import serial
import time
import csv
import re
import tkinter as tk
from tkinter import messagebox
import os

#Este código sirve para recolectar y mandar información por conexión serial - fluke ESA615 a un csv
#Para que después data_to_excel pueda mandar dicha información al excel donde se encuentra el informe


#la lógica consiste en que cada celda donde debe ir los datos recopilados tiene un código,
#se le asigna el mismo código a cierto paquete de información, entonces después
# data_to_excel busca parejas de códigos para poner la información donde corresponda.

#los códigos por celda fueron asignados por mí y sirven para simplificar mucho la lógica
# del código, es un código para cada tipo de prueba.

#Segunda hoja
def alerta_usuario(mensaje):
    # Se crea la ventana raíz oculta dentro de la función para no bloquear el import
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    messagebox.showinfo("ATENCIÓN - Acción Requerida", mensaje)
    root.destroy()


def ejecutar():
    PUERTO_COM = os.environ.get('PUERTO_FLUKE', 'COM3')

    ser = serial.Serial()
    ser.port = PUERTO_COM
    ser.baudrate = 115200
    ser.bytesize = serial.EIGHTBITS
    ser.parity = serial.PARITY_NONE
    ser.stopbits = serial.STOPBITS_ONE
    ser.rtscts = True
    ser.timeout = 1

    resultados_finales = []

    def enviar_instruccion(comando, pausa=0.1):
        codificado = (comando + '\r').encode('ascii')
        ser.write(codificado)
        time.sleep(pausa)

    def ejecutar_prueba(categoria, nombre_prueba, comandos_configuracion, codigo_excel, guardar=True):
        print(f"\n[{categoria}] {nombre_prueba} - Configurando...")
        for cmd in comandos_configuracion:
            enviar_instruccion(cmd, pausa=0.1)

        time.sleep(0.4)
        ser.reset_input_buffer()

        print("Capturando 3 muestras...")
        enviar_instruccion("MREAD", pausa=0.1)

        muestras_capturadas = 0
        valores = []
        intentos_vacios = 0
        #3 muestras capturadas para realizar un promedio, el promedio es el que será enviado al excel
        while muestras_capturadas < 3:
            linea_raw = ser.readline().decode('ascii', errors='ignore').strip()

            if not linea_raw:
                intentos_vacios += 1
                if intentos_vacios >= 2:
                    enviar_instruccion("READ", pausa=0.1)
                continue

            if "*" in linea_raw and len(linea_raw) < 3:
                continue
            #Sugerencia de gemini
            numeros_encontrados = re.findall(r"[-+]?\d*\.\d+|\d+", linea_raw)

            if numeros_encontrados:
                try:
                    valor = float(numeros_encontrados[0])
                    valores.append(valor)
                    muestras_capturadas += 1
                    intentos_vacios = 0
                    print(f"  Muestra {muestras_capturadas}/3: {valor}")
                except ValueError:
                    pass

        ser.write(b'\x1B')
        time.sleep(0.2)

        promedio = sum(valores) / len(valores) if valores else 0.0
        print(f"--> PROMEDIO OBTENIDO ({nombre_prueba}): {promedio:.2f}")

        if guardar:
            resultados_finales.append({
                "Categoria": categoria,
                "Prueba": nombre_prueba,
                "Codigo": codigo_excel,
                "Promedio": round(promedio, 3)
            })

        return promedio

    try:
        if not ser.is_open:
            ser.open()

        time.sleep(1.5)
        ser.reset_input_buffer()
        ser.reset_output_buffer()

        print("--- INICIALIZACIÓN ---")
        ser.write(b'\r')
        time.sleep(0.5)

        enviar_instruccion("REMOTE", pausa=0.5)
        enviar_instruccion("STD=601", pausa=0.5)

        alerta_usuario(
            "Asegúrese de que el Monitor de Signos Vitales esté ENCENDIDO.\n\nTodas las pruebas se realizarán en este estado.\n\nPresione ACEPTAR para comenzar.")

        p_fn = ejecutar_prueba("1. Linea", "Voltaje Fase-Neutro", ["MAINS=L1-L2"], "PL_FN", guardar=False)
        p_nt = ejecutar_prueba("1. Linea", "Voltaje Neutro-Tierra", ["MAINS=L2-GND"], "PL_NT", guardar=False)
        p_ft = ejecutar_prueba("1. Linea", "Voltaje Fase-Tierra", ["MAINS=L1-GND"], "PL_FT", guardar=False)

        promedio_pl_total = (p_fn + p_nt + p_ft) / 3
        print(f"\n==> GRAN PROMEDIO DE LÍNEA CALCULADO: {promedio_pl_total:.2f}")

        resultados_finales.append({
            "Categoria": "1. Linea",
            "Prueba": "Promedio Total de Línea",
            "Codigo": "PL",
            "Promedio": round(promedio_pl_total, 3)
        })

        ejecutar_prueba("2. Resistencia Tierra", "200mA", ["ERES"], "PR")

        pruebas_tierra = [
            {"desc": "Fase Normal, N Cerrado (NC)", "cmds": ["EARTHL", "POL=N", "NEUT=C"], "cod": "TNC"},
            {"desc": "Fase Normal, N Abierto (SFC)", "cmds": ["EARTHL", "POL=N", "NEUT=O"], "cod": "TNA"},
            {"desc": "Fase OFF", "cmds": ["EARTHL", "POL=OFF"], "cod": "TOFF"},
            {"desc": "Fase Invertida, N Cerrado (NC)", "cmds": ["EARTHL", "POL=R", "NEUT=C"], "cod": "TIC"},
            {"desc": "Fase Invertida, N Abierto (SFC)", "cmds": ["EARTHL", "POL=R", "NEUT=O"], "cod": "TIA"}
        ]

        for pt in pruebas_tierra:
            ejecutar_prueba("3. Fuga a Tierra", pt['desc'], pt["cmds"], pt["cod"])

        pruebas_chasis = [
            {"desc": "F. Normal, N Cerrado, T Cerrado (NC)", "cmds": ["ENCL", "POL=N", "NEUT=C", "EARTH=C"],
             "cod": "CNCC"},
            {"desc": "F. Normal, N Cerrado, T Abierto (SFC)", "cmds": ["ENCL", "POL=N", "NEUT=C", "EARTH=O"],
             "cod": "CNCA"},
            {"desc": "F. Normal, N Abierto, T Cerrado (SFC)", "cmds": ["ENCL", "POL=N", "NEUT=O", "EARTH=C"],
             "cod": "CNAC"},
            {"desc": "F. Invertida, N Cerrado, T Cerrado (NC)", "cmds": ["ENCL", "POL=R", "NEUT=C", "EARTH=C"],
             "cod": "CICC"},
            {"desc": "F. Invertida, N Abierto, T Cerrado (SFC)", "cmds": ["ENCL", "POL=R", "NEUT=O", "EARTH=C"],
             "cod": "CIAC"},
            {"desc": "F. Invertida, N Cerrado, T Abierto (SFC)", "cmds": ["ENCL", "POL=R", "NEUT=C", "EARTH=O"],
             "cod": "CICA"}
        ]

        for pc in pruebas_chasis:
            ejecutar_prueba("4. Fuga Chasis", pc['desc'], pc["cmds"], pc["cod"])

        alerta_usuario("Conecte Latiguillos.\n\nPresione ACEPTAR para continuar.")

        partes_latiguillos = ["RA", "LL", "LA", "RL", "V1"]

        condiciones_lat = [
            ("Normal, NC, TC (NC)", "NCC", ["POL=N", "NEUT=C", "EARTH=C"]),
            ("Normal, NC, TA (SFC)", "NCA", ["POL=N", "NEUT=C", "EARTH=O"]),
            ("Normal, NA, TC (SFC)", "NAC", ["POL=N", "NEUT=O", "EARTH=C"]),
            ("Invertida, NC, TC (NC)", "ICC", ["POL=R", "NEUT=C", "EARTH=C"]),
            ("Invertida, NA, TC (SFC)", "IAC", ["POL=R", "NEUT=O", "EARTH=C"]),
            ("Invertida, NC, TA (SFC)", "ICA", ["POL=R", "NEUT=C", "EARTH=O"])
        ]

        for parte in partes_latiguillos:
            cmd_aislar = f"AP={parte}//GND"

            for desc_cond, sufijo_cod, cmds_cond in condiciones_lat:
                nombre_prueba = f"Latiguillo {parte} - {desc_cond}"
                codigo_etiqueta = f"L{parte}{sufijo_cod}"
                comandos_completos = ["PAT", "MODE=ACDC", cmd_aislar] + cmds_cond
                ejecutar_prueba("5. Fuga Latiguillos", nombre_prueba, comandos_completos, codigo_etiqueta)

        enviar_instruccion("LOCAL", pausa=1.0)
        print("\n--- EQUIPO EN MODO LOCAL ---")

        #Guardamos en la carpeta ArchivosGenerados relativa al ejecutable
        import sys
        if getattr(sys, 'frozen', False):
            directorio_base = os.path.dirname(sys.executable)
        else:
            directorio_base = os.path.dirname(os.path.abspath(__file__))

        carpeta_generados = os.path.join(directorio_base, "ArchivosGenerados")
        os.makedirs(carpeta_generados, exist_ok=True)
        nombre_archivo = os.path.join(carpeta_generados, "resultados_esa615.csv")

        with open(nombre_archivo, mode='w', newline='', encoding='utf-8') as archivo_csv:
            writer = csv.DictWriter(archivo_csv, fieldnames=["Categoria", "Prueba", "Codigo", "Promedio"])
            writer.writeheader()
            for fila in resultados_finales:
                writer.writerow(fila)

        print(f"\n¡ÉXITO! Recopilación terminada. Los datos se guardaron en:\n{nombre_archivo}")

    except serial.SerialException as e:
        print(f"\n[ERROR SERIAL] {e}")
        raise e  #Es importante relanzar el error para que el Main.py lo detecte
    except Exception as e:
        print(f"\n[ERROR INESPERADO] {e}")
        raise e
    finally:
        if ser.is_open:
            try:
                ser.write(b'\x1B')
                time.sleep(0.2)
                ser.write(b'LOCAL\r')
                time.sleep(0.5)
            except:
                pass
            ser.close()
            print("Puerto cerrado.")


if __name__ == "__main__":
    ejecutar()