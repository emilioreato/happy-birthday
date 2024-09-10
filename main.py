"""
Proyecto de calendario para cumpleaños hecho por Emilio Reato el 10 de septiembre de 2024.
Despliega un menu en la consola con el cual se puede operar y modificar los recordatorios.
Cuando llegue el día de uno de los cumpleaños registrados, sonará una notificacion de windows para que no se te olvide!.

He recibido la ayuda de chatgpt para investigar cuales librerias podrian funcionarme mejor y para crear ciertas funciones simples como check_registry_key_exists o ocultar_ventana.
El resultado final óptimo del proyecto es un .exe empaquetado gracias a auto_py_to_exe o pyinstaller --onedir que cuenta con una carpeta auxiliar(_internal) donde se guardan todos los archivos necesarios, aunque tambien funciona si se ejecuta el archivo .py (no se si en el startup)
"""

import time
import datetime
import csv
import sys
import os
import re
import ctypes
import psutil
import winreg as reg
from winotify import Notification, audio


os.chdir(os.path.dirname(os.path.abspath(__file__)))  # establece el directorio actual al del archivo
ctypes.windll.kernel32.SetConsoleTitleW("Happy Birthday by Emilio Reato")  # establecemos el nombre de la ventana
data = 'data.csv'  # nombre de los archivos necesarios
notification_register = 'already_notified.txt'
notification_icon = 'notification_icon.ico'


def kill_process(proc):  # FUNCION PARA MATAR PROCESOS
    proc_name = proc.name()
    try:
        # kill_list = get_executable_name()
        if get_executable_name().lower() in proc_name.lower() or "happy birthday" in proc_name.lower():

            if (os.getpid() != proc.pid):  # NO NOS ELIMINEMOS A NOSOTROS MISMOS, ELIMINEMOS PRIMERO A NUESTRA COPIA EN SEGUNDO PLANO
                print("\nTerminando el programa en segundo plano")  # f"\nTerminando proceso {proc_name} (PID: {proc.pid})"
                os.kill(proc.pid, 9)  # Enviar la señal kill para forzar la terminación
                time.sleep(1)
                print("\nTerminando este mismo proceso")
                time.sleep(1)
                exit()  # luego de eliminarlo, nos cerramos a nosotros mismos

    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass  # Si no podemos acceder al proceso o ya no existe, lo ignoramos
        print("unable")


def get_executable_name():  # FUNCION Q OBTIENE EL NOMBRE DEL PROPIO ARCHIVO, EL EJECTABLE
    if getattr(sys, 'frozen', False):  # El script está en un ejecutable
        return os.path.basename(sys.argv[0])  # Devuelve el nombre del archivo .exe
    else:
        return os.path.basename(__file__)  # Devuelve el nombre del script .py


def check_registry_key_exists(exe_path, app_name):  # SE FIJA SI CIERTO REGISTRO YA EXISTE EN EL REGEDIT
    try:
        # Intentar abrir la clave del registro \Software\Microsoft\Windows\CurrentVersion\Run

        with reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_READ) as key:
            if (str(exe_path)+" /startup" != reg.QueryValueEx(key, app_name)[0]):  # si no existe el valor, o no coincide sus datos,fallara en esta linea.
                return False
            return True  # La clave existe
    except FileNotFoundError:
        return False  # La clave no existe


def regedit_conf(exe_path, app_name):  # IF THERE ISNT A REGEDIT START UP CONFIG FOR THIS PROGRAM, CREATE ONE
    if not check_registry_key_exists(exe_path, app_name):
        try:
            key = reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_WRITE)  # Open the registry key with write access
            reg.SetValueEx(key, app_name, 0, reg.REG_SZ, str(exe_path)+" /startup")  # Add the entry for the executable
            reg.CloseKey(key)  # Close the registry key
            print(f"Instalación finalizada: {app_name} ha sido añadido al startup.")

        except Exception as e:
            print(f"No se ha podido añadir al startup. Instalación fallida: {e}")
    else:
        # print("already in start up")
        pass


def ocultar_ventana():  # OCULTA LA CONSOLA DEL CMD Y PASA EL PROCESO A "SEGUNDO PLANO"
    ventana = ctypes.windll.kernel32.GetConsoleWindow()  # Obtener el identificador de la ventana de la consola
    if ventana != 0:
        ctypes.windll.user32.ShowWindow(ventana, 0)  # Ocultar la ventana


def leer_linea_csv(nombre_archivo, numero_linea):  # LEE LINEA ESPECIFICA EN UN CSV
    with open(nombre_archivo, mode='r', newline='') as archivo_csv:
        filas = csv.reader(archivo_csv)
        for i, fila in enumerate(filas, start=1):
            if i == numero_linea:
                return fila
    return None  # Si la línea no existe


def escribir_linea_csv(nombre_archivo, numero_linea, nueva_fila):  # REESCRIBE UNA LINEA ESPECIFICA EN UN CSV
    # Leer todo el contenido del archivo CSV
    with open(nombre_archivo, mode='r', newline='') as archivo_csv:
        lineas = list(csv.reader(archivo_csv))

    # Modificar la línea deseada
    if 0 < numero_linea <= len(lineas):
        lineas[numero_linea - 1] = nueva_fila
    else:
        # print(f"La línea {numero_linea} no existe.")
        return

    # Volver a escribir todo el contenido en el archivo
    with open(nombre_archivo, mode='w', newline='') as archivo_csv:
        escritor = csv.writer(archivo_csv)
        escritor.writerows(lineas)
        # print(f"Línea {numero_linea} actualizada correctamente.")


def segundos_para_fin_dia():  # DEVUELVE EL TOTAL DE LOS SEGUNDOS QUE QUEDAN PARA Q TERMINE EL DIA
    # Obtener la hora actual
    ahora = datetime.datetime.now()

    # Calcular el fin del día (medianoche del siguiente día)
    fin_dia = (ahora + datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)

    # Calcular la diferencia en segundos
    segundos_restantes = (fin_dia - ahora).total_seconds()

    return int(segundos_restantes)


def get_index(date_dd_mm):  # whe get how many days has passed since the beginning of the year, aka, the row in the file(called index).
    date = datetime.datetime.strptime("2024-"+date_dd_mm, "%Y-%d-%m")
    index = date.timetuple().tm_yday  # get the day since the year started/index
    return index+1  # +1 because there is a title row in the file


def check_n_format_date(ugly_date):  # SE FIJA SI LA FECHA Q LE PASAMOS ESTA BIEN Y LA REFORMATE A LA DESEADA PARA ESTE CODIGO

    if re.search(r"^[0123]?\d{1}(\W+)?[01]?\d{1}$", ugly_date):  # check the date introduced is valid
        try:
            # REFORMATEAMOS EL WHEN PARA Q SE PUEDA INGRESAR EN CUALQUIER FORMATO
            ugly_date = re.sub(r'[^\d]', '-', ugly_date)  # quitamos espacios y ponemos guion
            ugly_date = re.sub(r'-+', '-', ugly_date)  # si hay varios guines(anteriormente espacios u otros caracteres) ponemos uno solo

            if not "-" in ugly_date:  # creamos la fecha con el formato correspondiente
                ugly_date = datetime.datetime.strptime(ugly_date, "%d%m")
            else:
                ugly_date = datetime.datetime.strptime(ugly_date, "%d-%m")
            nice_date = ugly_date.strftime("%d-%m")  # cambiamos el formato al duplicado, dd-mm y no solo d-m
            return True, nice_date
        except:
            False, None  # if it wasnt possible to format, return false
    return False, None  # return false if it is not even valid


def display_menu():  # MENU PRINCIPAL

    print("\"add\"  --> para añadir un nuevo cumpleañero.")
    print("\"del\"  --> para borrar un cumpleañero.")
    print("\"ok\"   --> para entrar en standbyte.")
    print("\n\"show\" --> para mostrar todo el registro")
    print("\"exit\" --> para retroceder/salir.")
    print("\"kill\" --> para matar el programa por completo")

    valido = ""

    while True:

        choice = input("\nIngrese un comando" + valido + ": ").strip().lower()

        if (choice == "add"):

            while True:

                name = input("\nPerfecto, ingrese el nombre: ").strip()

                if (name.lower() == "exit"):
                    print("\n<--", end="")
                    break  # si se usa el comando exit salimos del menu proximo

                while True:
                    when = input("\nIngrese el día [dd-mm]: ").strip()

                    if (when == "exit"):
                        print("\n<--", end="")
                        break  # si se usa el comando exit salimos del menu proximo

                    is_valid, when_formatted = check_n_format_date(when)

                    if (is_valid):

                        index = get_index(when_formatted)

                        updated_content = leer_linea_csv(data, index)  # read what is already written in the register
                        updated_content.append(name)  # add the new name
                        escribir_linea_csv(data, index, updated_content)  # write the updated register

                        os.system('cls')
                        print(f"\n>>>({when_formatted}) {name.title()} agregado.\n")
                        display_menu()

        elif (choice == "del"):

            while True:

                entry = input("\nIngrese un nombre o una fecha[dd-mm] para buscar: ").strip().lower()

                if (entry == "exit"):
                    print("\n<--", end="")
                    break  # si se usa el comando exit salimos del menu proximo

                is_valid, when_formatted = check_n_format_date(entry)  # obtenemos la validez de la fecha y su version formateada

                if (is_valid):
                    # when = input("\nIngrese el día del que desea eliminar [dd-mm]: ").strip()
                    index = get_index(when_formatted)  # obtenemos el index/dia desde inicio del año

                    registered_names_n_date = leer_linea_csv(data, index)  # leemos la linea numero (index) y guardamos la lista de los valores separados por comas

                    if (len(registered_names_n_date) > 1):  # si hay mas de un elemento en la linea leida, es decir, un nombre ademas de la fecha:
                        print()
                        print(f"Coincidencias encontradas para {when_formatted}:")
                        for i, name in enumerate(registered_names_n_date, start=-1):  # imprimimos todos los valores/nombres dentro de la lista
                            if (name != when_formatted):  # todos excepto la fecha xd
                                print(f"{i+1} - {name}")  # le agregamos un valor numero i al lado para q el usuario pueda elegir

                        while True:

                            choice = input("\nIngrese el número de quien desea eliminar: ").strip()

                            if (choice.lower() == "exit"):  # si lo ingresado fue exit, salimos del bucle
                                print("\n<--", end="")
                                break

                            while True:  # bucle para asegurarnos que sea correcto lo que ingresa
                                while (not re.search(r"^\d{1,2}$", choice)):
                                    choice = input("\nIngrese el número de quien desea eliminar: ").strip()
                                if (int(choice) <= len(registered_names_n_date)-1):
                                    break
                                choice = ""

                            deleted = registered_names_n_date.pop(int(choice))  # eliminamos lo el valor q el usuario elegió de la copia q teniamos de la linea leida

                            escribir_linea_csv(data, index, registered_names_n_date)  # acutalizamos la linea dentro del fichero data
                            os.system('cls')
                            print(f"\n({when_formatted}) {deleted.title()} borrado.\n")
                            display_menu()  # volvermos al menu princical
                    else:
                        print(f">No hay cumpleaños registrados en la fecha: {when_formatted}.")
                else:

                    coincidences = list()

                    while True:

                        # leemos el archivo y vamos fila por filas, valor por valor separado en coma y dentro de el buscamos coincidencias con lo ingresado en input.
                        with open(data, mode='r', newline='') as archivo_csv:

                            filas = csv.reader(archivo_csv)

                            for i, fila in enumerate(filas, start=0):
                                for x, value in enumerate(fila, start=0):
                                    if entry in value.lower():  # si hay coincidencias guardamos un par de datos necesarios para luego identificar el elemento al q borrar y tambien para poder mostrar la info en pantalla
                                        dic = {"value": value,  # el valor/nombre de la coincidencia
                                               "fecha": fila[0],  # la fecha donde se ubica de la coincidencia
                                               # "i": i,  # el index/dias desde el incio de año (no necesario porque dps uso get_index con ["fecha"])
                                               "x": x,
                                               "fila": fila}
                                        coincidences.append(dic)  # añadimos el diccionario a la lista

                        cantidad_coincidencias = len(coincidences)  # averiguamos la cantidad de coincidencias
                        if (cantidad_coincidencias > 0):  # si se han encontrado coincidencias

                            if (cantidad_coincidencias > 1):  # si son plurales las coincidencias

                                print("\nCoincidencias encontradas:")
                                for val, _ in enumerate(coincidences):  # imprimimos los datos de cada coincidencia en la pantalla para q el usuario pueda eligir cual borrar
                                    print(f"{val+1} - ", coincidences[val]['value'].title(), f"({coincidences[val]['fecha']})")
                                out = False
                                while not out:  # bucle para asegurar q lo ingresado no rompa el codigo/ no sea valido
                                    which = ""
                                    while (not re.search(r"\d+", which) and not out):
                                        which = input("\nIngrese el número de quien desea borrar: ").strip().lower()  # pedimos el numero de cual desea borrar
                                        if (which.strip().lower() == "exit"):  # si ha ingresado exit lo retornamos al menu anterior
                                            out = True
                                            break
                                    try:
                                        which = int(which)
                                        if (which <= cantidad_coincidencias):  # si inserto un numero mas alto al de las coincidencias encontradas, le pedimos q lo haga denuevo
                                            break
                                    except:
                                        pass

                                if (out):
                                    print("\n<--", end="")
                                    break

                            else:  # en caso de que solo haya 1 coincidencia
                                print("\n>Se ha encontrado a:", coincidences[0]["value"].title(), f"({coincidences[0]['fecha']})")  # imprimimos los datos de la coincidencia
                                if (input("\n¿Desea borrarlo del registro? [si/no]: ").strip().lower() == "si"):  # si
                                    which = 1  # esto para q dps cuando se le reste 1 quede con 0 y se pueda referir al unico elemento de la lista de coincidencias
                                    pass
                                else:  # si ha optado por no borrarlo lo devolvemos al menu anterior
                                    print("Operacion cancelada.")
                                    print("\n<--", end="")
                                    break

                            # borramos el valor q el usuario eligió de la linea que lo contenia, asi luego la reescribimos sin él
                            deleted = coincidences[which-1]['fila'].pop(coincidences[which-1]['x'])
                            # reescribimos la linea ahora sin el valor que ha deseado eliminar el usuario. se usan los datos previamente almacenados de la coincidencia elegida. (coincidences[which-1]["i"]+1 --> not useful anymore)
                            escribir_linea_csv(data, get_index(coincidences[which-1]["fecha"]), coincidences[which-1]['fila'])

                            os.system('cls')
                            print(f"\n({coincidences[which-1]['fecha']}) {deleted.title()} borrado.\n")
                            display_menu()  # salimos al menu principal
                        else:
                            print(">No hay nadie registrado con ese nombre.")
                            break

        elif (choice == "ok"):

            occurance = 0

            for process in psutil.process_iter(['pid', 'name']):  # iteramos entre todos los procesos q se estan ejecutando ahora mismo
                # si alguno de ellos se llama como este propio programa, es decir, este mismo programa se esta ejecutando, subimos un contador de ocurrencia
                if "happy birthday" in process.name().lower() or re.sub(r'\\_internal', '', os.path.abspath(get_executable_name())) in process.name().lower():
                    occurance += 1

            if (occurance > 1):  # si hay mas de 1 proceso con el nombre de este programa/archivo es porque ya se estaba ejecutando en segundo plano porque ya ha sido puesto en standbyte (manualmente o por el propio startup)
                print(f"\n(ya ejecutandose en segundo plano({occurance})) saliendo", end="")  # avisamos
                for _ in range(3):  # for para crear una linea de puntos
                    time.sleep(0.42)
                    print(".", end="")
                    sys.stdout.flush()
                time.sleep(1.4)

                exit()  # terminamos la ejecucion de este mismo programa dejando al otro tranquilo ejecutandose en segundo plano

            else:  # si solo hay 1 programa ejecutandose,osea este, entonces lo hacemos entrar en standbyte
                print("\nponiendo en modo standbyte", end="")  # avisamos
                for _ in range(3):  # for para crear una linea de puntos
                    time.sleep(0.42)
                    print(".", end="")
                    sys.stdout.flush()
                time.sleep(1.4)

                ocultar_ventana()
                standbyte()  # entramos en modo espera

        elif (choice.lower() == "exit"):
            exit()  # si se usa el comando exit salimos del programa
        elif (choice.lower() == "kill"):  # si el usuario quiere matar al programa definitivamente y q no se ejecute en segundo plano:
            for proc in psutil.process_iter(['pid', 'name']):
                kill_process(proc)  # llamamos a una funcion q elimina primero el otro proceso ya ejecutandose en segundo plano, y luego se termina a si mismo. si no lo hay, no hace nada

            # si si se hubiera encontrado un proceso en segundo plano esta linea nunca se ejecutaria
            time.sleep(0.8)
            print("\n>No se ha encontrado otro proceso en segundo plano")
            time.sleep(1.2)
            print("\n>Se procederá a cerrar únicamente este")
            time.sleep(2)
            exit()

        else:
            valido = " válido"
            continue


def check_coincidences():  # SE FIJA SI HAY ALGUN NOMBRE EN EL EL ARCHIVO DATA EN LA FILA DEL DIA DE HOY, MANDA NOTIFICACION SI ES EL CASO Y GUARDA EL ENVIO DE LA NOTIFICACION PARA NO VOLVER A HACERLO EN EL DIA

    with open('already_notified.txt', mode='r', newline='') as file:  # leemos el registro de notificaciones diarias
        register = file.read()

    today = datetime.datetime.today().date()
    today = today.strftime("%d-%m")  # obtenemos la fecha de hoy

    if not today in register:  # si la fecha de hoy no se ecuentra en el registro es porque todavia no hemos notificado

        registered_names_today = leer_linea_csv("data.csv", get_index(today))  # leemos los cumpleaños de hoy
        registered_names_today.pop(0)

        cantidad = len(registered_names_today)

        if (cantidad > 0):  # si alguien cumple años

            if (cantidad > 1):
                plural = "s"
                nacimiento = "nacieron"
            else:
                plural = ""
                nacimiento = "nació"
            if (cantidad > 3):
                names_string = " "+", ".join(registered_names_today).title()  # si son mas de 3 no van a entrar en fromato lista en la notificacion entonces los concatenamos
            else:
                names_string = "\n> "+"\n> ".join(registered_names_today).title()  # si son 3 o menos los cumpleañeros los ponemos en formato lista asi queda mejor visualmente

            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "noti_icon.png")  # cargamos el icono de la notificacion

            # estas 3 lineas para darle el formato de texto, es decir, 9 de septiembre.
            today2 = datetime.datetime.now()
            meses = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"}
            fecha_formateada_a_texto = f"{today2.day} de {meses[today2.month]}"  # Formatear la fecha en formato "día de mes"

            toast = Notification(title=f'Hoy es el cumpleaños de {cantidad} persona'+plural+".",
                                 msg=f'El {fecha_formateada_a_texto} {nacimiento}:' + names_string+".",
                                 duration="long",
                                 app_id=" "*30+"Happy Birthday",
                                 icon=icon_path)
            toast.set_audio(audio.Reminder, loop=False)
            toast.show()

            """notification.notify(  # notificamos con todos los nombres y la cantidad de cumpleañeros
                title=f'Hoy es el cumpleaños de {cantidad} persona'+plural+".",
                message=f'El {today} nacieron:' + names_string+".",
                app_name='Happy Birthday',
                timeout=20,
                app_icon=notification_icon)"""  # old notificacion method

            with open('already_notified.txt', mode='w', newline='') as file:  # escribimos la fecha de hoy asi no volveremos a notificar
                file.write(today)
        else:
            # print("no coincidences/ no birthdays")
            pass
    else:
        # print("already notified ")
        pass


def standbyte():  # estado de standbyte donde borramos todo lo innecesario, ocultamos la ventana, notificamos cumpleaños en caso de que haya y dormimos hasta el proximo dia

    ocultar_ventana()

    time.sleep(120)

    while True:

        check_coincidences()  # funcion q busca si hay alguien registrado en el dia de hoy, y si es asi, notificamos.

        if (segundos_para_fin_dia() < 21600):
            time.sleep(segundos_para_fin_dia()+180)  # +120 para asegurarnos de que haga tiempo de checkear con el dia actualizado y no q se vuelva a dormir
        else:
            time.sleep(21600)  # hay un check_coincidences cada 6 horas, 4 por dia

        with open("intervales_register.txt", mode='a', newline='') as file:
            file.write(str(datetime.datetime.today())+"\n")  # escribimos en el archivo el horario y fecha del check_coincidences()


def main():

    # añadimos nustro programa al start up del regedit si es que ya no lo esta. le pasamos la ruta donde esta nuestro .exe y como se llama nuestro .exe asi le da ese nombre y no hay confusiones
    regedit_conf(re.sub(r'\\_internal', '', os.path.abspath(get_executable_name())), get_executable_name())  # __file__

    if not os.path.exists(data) or os.path.getsize(data) == 0:  # en caso de que los archivos auxiliares no existan los creamos y llenamos

        all_dates = ((datetime.datetime(2024, 1, 1) + datetime.timedelta(days=i)).strftime('%d-%m') for i in range(366))  # creamos una lista con todos los dias del año. 2024 because it has 2/29

        with open(data, mode='w', newline='') as file:  # abrimos (o creamos) el archivo.
            writer = csv.writer(file)
            writer.writerow(["dd-mm,nombres*"])  # escribimos una primera fila como titulo de cada columna

            for date in all_dates:
                writer.writerow([date])  # escribimos cada una de las fechas en cada linea del nuevo archivo

    if not os.path.exists(notification_register):
        with open(notification_register, mode='w', newline='') as file:  # tambien creamos el archivo q guarda la info de si ya hemos notificado hoy
            file.write("")

    # podria tambien crear el archivo q guarda cada intervalo de checks pero ni ganas

    print("\n>>>Bienvenido a Happy Birthday!!\n")
    display_menu()


if (__name__ == "__main__"):

    # si el script está siendo ejecutado por el sistema: (porque cuando se llama a la ejecucion del archivo desde el stratup va con el argumento "/startup" porque previamente se le agrego para poder asi identificarlo)
    if "/startup" in sys.argv:
        # print("standbyte")
        time.sleep(0.2)
        standbyte()
    else:  # si fue ejecutado manualmente por el usuario
        # print("main executed")
        time.sleep(0.2)
        main()
