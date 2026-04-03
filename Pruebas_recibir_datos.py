import serial
import time
import csv
import re
import os
import tkinter as tk
from tkinter import messagebox

#Este código sirve para recolectar y mandar información por conexión serial - fluke ESA615 a un csv
#Para que después data_to_excel pueda mandar dicha información al excel donde se encuentra el informe


#la lógica consiste en que cada celda donde debe ir los datos recopilados tiene un código,
#se le asigna el mismo código a cierto paquete de información, entonces después
# data_to_excel busca parejas de códigos para poner la información donde corresponda.

#los códigos por celda fueron asignados por mí y sirven para simplificar mucho la lógica
# del código, es un código para cada tipo de prueba.

#Primera hoja
def alerta_usuario(mensaje):
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

    def ejecutar_prueba_exacta(nombre_prueba, comandos_configuracion, codigo_excel):
        print(f"Configurando: {nombre_prueba} [{codigo_excel}]...")
        for cmd in comandos_configuracion:
            enviar_instruccion(cmd, pausa=0.1)

        time.sleep(0.4)
        ser.reset_input_buffer()

        enviar_instruccion("MREAD", pausa=0.1)
        muestras_capturadas = 0
        valores = []
        intentos_vacios = 0

        while muestras_capturadas < 3:
            linea_raw = ser.readline().decode('ascii', errors='ignore').strip()

            if not linea_raw:
                intentos_vacios += 1
                if intentos_vacios >= 2:
                    enviar_instruccion("READ", pausa=0.1)
                continue

            if "*" in linea_raw and len(linea_raw) < 3:
                continue

            numeros_encontrados = re.findall(r"[-+]?\d*\.\d+|\d+", linea_raw)

            if numeros_encontrados:
                try:
                    valor = float(numeros_encontrados[0])
                    valores.append(valor)
                    muestras_capturadas += 1
                    intentos_vacios = 0
                except ValueError:
                    pass

        ser.write(b'\x1B')
        time.sleep(0.2)

        promedio = sum(valores) / len(valores) if valores else 0.0
        print(f"  -> Valor Registrado: {promedio:.3f}")

        resultados_finales.append({
            "Codigo": codigo_excel,
            "Valor": round(promedio, 3)
        })

    def prueba_linea():
        print("\n--- INICIANDO PRUEBA DE LÍNEA ---")
        ejecutar_prueba_exacta("Fase-Neutro", ["MAINS=L1-L2"], "LFN")
        ejecutar_prueba_exacta("Neutro-Tierra", ["MAINS=L2-GND"], "LNT")
        ejecutar_prueba_exacta("Fase-Tierra", ["MAINS=L1-GND"], "LFT")

    def bloque_pruebas(n):
        print(f"\n=============================================")
        print(f"       INICIANDO BLOQUE DE PRUEBAS #{n}")
        print(f"=============================================")

        print("\n[Prueba de resistencia]")
        ejecutar_prueba_exacta("Resistencia Tierra", ["ERES"], f"PR{n}")

        print("\n[Pruebas Tierra]")
        ejecutar_prueba_exacta("ON, F-Norm, N-C", ["EARTHL", "POL=N", "NEUT=C"], f"TONC{n}")
        ejecutar_prueba_exacta("OFF, F-Norm, N-A", ["EARTHL", "POL=N", "NEUT=O"], f"TFNA{n}")
        ejecutar_prueba_exacta("OFF, F-OFF", ["EARTHL", "POL=OFF"], f"TFF{n}")
        ejecutar_prueba_exacta("OFF, F-Inv, N-A", ["EARTHL", "POL=R", "NEUT=O"], f"TFIA{n}")
        ejecutar_prueba_exacta("ON, F-Inv, N-C", ["EARTHL", "POL=R", "NEUT=C"], f"TOIC{n}")

        print("\n[Pruebas Chasis]")
        ejecutar_prueba_exacta("ON, F-Inv, N-C, T-C", ["ENCL", "POL=R", "NEUT=C", "EARTH=C"], f"COICC{n}")
        ejecutar_prueba_exacta("ON, F-Inv, N-C, T-A", ["ENCL", "POL=R", "NEUT=C", "EARTH=O"], f"COICA{n}")
        ejecutar_prueba_exacta("OFF, F-Inv, N-A, T-A", ["ENCL", "POL=R", "NEUT=O", "EARTH=O"], f"COIAA{n}")
        ejecutar_prueba_exacta("OFF, F-Inv, N-A, T-C", ["ENCL", "POL=R", "NEUT=O", "EARTH=C"], f"COIAC{n}")
        ejecutar_prueba_exacta("OFF, F-OFF, T-A", ["ENCL", "POL=OFF", "EARTH=O"], f"COOA{n}")
        ejecutar_prueba_exacta("OFF, F-OFF, T-C", ["ENCL", "POL=OFF", "EARTH=C"], f"COOC{n}")
        ejecutar_prueba_exacta("OFF, F-Norm, N-A, T-C", ["ENCL", "POL=N", "NEUT=O", "EARTH=C"], f"CONAC{n}")
        ejecutar_prueba_exacta("OFF, F-Norm, N-C, T-C", ["ENCL", "POL=N", "NEUT=C", "EARTH=C"], f"CONCC{n}")
        ejecutar_prueba_exacta("OFF, F-Norm, N-C, T-A", ["ENCL", "POL=N", "NEUT=C", "EARTH=O"], f"CONCA{n}")
        ejecutar_prueba_exacta("OFF, F-Norm, N-A, T-A", ["ENCL", "POL=N", "NEUT=O", "EARTH=O"], f"CONAA{n}")

        print("\n[Pruebas Latiguillos]")
        ejecutar_prueba_exacta("ON, F-Norm, N-C", ["PAT", "MODE=ACDC", "AP=ALL//GND", "POL=N", "NEUT=C"], f"LONC{n}")
        ejecutar_prueba_exacta("OFF, F-Norm, N-A", ["PAT", "MODE=ACDC", "AP=ALL//GND", "POL=N", "NEUT=O"], f"LONA{n}")
        ejecutar_prueba_exacta("OFF, F-OFF", ["PAT", "MODE=ACDC", "AP=ALL//GND", "POL=OFF"], f"LOO{n}")
        ejecutar_prueba_exacta("OFF, F-Inv, N-A", ["PAT", "MODE=ACDC", "AP=ALL//GND", "POL=R", "NEUT=O"], f"LOIA{n}")
        ejecutar_prueba_exacta("ON, F-Inv, N-C", ["PAT", "MODE=ACDC", "AP=ALL//GND", "POL=R", "NEUT=C"], f"LOIC{n}")

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

        alerta_usuario("Equipo encendido y listo. Iniciando ejecución de bloques continuos.\n\nPresione ACEPTAR.")

        prueba_linea()

        for i in range(1, 4):
            bloque_pruebas(i)

        import sys
        if getattr(sys, 'frozen', False):
            directorio_base = os.path.dirname(sys.executable)
        else:
            directorio_base = os.path.dirname(os.path.abspath(__file__))

        carpeta_generados = os.path.join(directorio_base, "ArchivosGenerados")
        os.makedirs(carpeta_generados, exist_ok=True)
        ruta_csv = os.path.join(carpeta_generados, "resultados_bloques_esa615.csv")

        with open(ruta_csv, mode='w', newline='', encoding='utf-8') as archivo_csv:
            writer = csv.DictWriter(archivo_csv, fieldnames=["Codigo", "Valor"])
            writer.writeheader()
            for fila in resultados_finales:
                writer.writerow(fila)

        print(f"\n¡ÉXITO! Recopilación terminada. Se generaron {len(resultados_finales)} mediciones.")
        print(f"Los datos se guardaron en: {ruta_csv}")

    except serial.SerialException as e:
        print(f"\n[ERROR SERIAL] {e}")
        raise e
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