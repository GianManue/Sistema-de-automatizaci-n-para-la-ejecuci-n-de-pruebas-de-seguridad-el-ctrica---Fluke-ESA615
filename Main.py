import os
import sys
import time
import serial
import serial.tools.list_ports

#Este código recopila el funcionamiento EN ORDEN de los otros 3 archivos, es como un orquestador
#sirve para que Appfluke o la interfaz accedan a los 3 archivos de forma directa


# Importamos los scripts anteriores
import Pruebas_recibir_info
import Pruebas_recibir_datos
import data_to_excel


def auto_detectar_fluke():
    print("\n>>> Buscando analizador Fluke ESA615 en puertos USB...")
    puertos_disponibles = serial.tools.list_ports.comports()
    #lógica buscar puertos
    for puerto in puertos_disponibles:
        try:
            ser = serial.Serial(
                port=puerto.device,
                baudrate=115200,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                rtscts=True,
                timeout=2
            )
            time.sleep(1)
            ser.write(b'IDENT\r')
            #esto fue corrección de IA no se bien que es pero fue efectivo
            respuesta = ser.read_until().decode('utf-8', errors='ignore').strip()
            ser.close()

            #Verificamos si la respuesta tiene sentido
            if respuesta:
                print(f">>> [ÉXITO] Fluke ESA615 detectado en: {puerto.device}")
                return puerto.device
        except Exception:
            pass

    print(">>> [ERROR] No se detectó ningún equipo conectado.")
    return None


def ejecutar_script_modulo(nombre_script, funcion_ejecutar):
    print(f"\n{'=' * 60}")
    print(f" INICIANDO MÓDULO: {nombre_script}")
    print(f"{'=' * 60}\n")
    try:
        funcion_ejecutar()
        print(f"\n [OK] {nombre_script} se ejecutó y finalizó correctamente.")
        return True
    except Exception as e:
        print(f"\n[ERROR EN MÓDULO {nombre_script}] {e}")
        return False


def iniciar_orquestador():
    print("\n" + "*" * 60)
    print("   🤖 INICIANDO AUTOMATIZACIÓN COMPLETA - FLUKE ESA615 🤖")
    print("*" * 60)

    tiempo_inicio = time.time()

    # fase de autodetección del puerto (también está arriba)
    puerto_detectado = auto_detectar_fluke()
    if not puerto_detectado:
        print("\n[PROCESO CANCELADO] Por favor, conecte el analizador e intente de nuevo.")
        return False

    #Guardamos el puerto en una variable de memoria para que los otros scripts la lean
    os.environ['PUERTO_FLUKE'] = puerto_detectado


    modulos_a_ejecutar = [
        ("Pruebas_recibir_info", Pruebas_recibir_info.ejecutar),
        ("Pruebas_recibir_datos", Pruebas_recibir_datos.ejecutar),
        ("data_to_excel", data_to_excel.ejecutar)
    ]

    for nombre, funcion in modulos_a_ejecutar:
        exito = ejecutar_script_modulo(nombre, funcion)

        if not exito:
            print("\n[PROCESO DETENIDO] La automatización se detuvo por seguridad.")
            return False

        if nombre != modulos_a_ejecutar[-1][0]:
            print(f" Esperando liberación del puerto serial...")
            time.sleep(2.5)

    tiempo_total = round(time.time() - tiempo_inicio, 2)
    print("\n" + "*" * 60)
    print("¡ANALISÍS DE SEGURIDAD FINALIZADO CON ÉXITO! ")
    print(f"Tiempo total de ejecución: {tiempo_total} segundos.")
    print("*" * 60 + "\n")
    return True


if __name__ == "__main__":
    iniciar_orquestador()