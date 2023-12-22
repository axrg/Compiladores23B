import json
import webbrowser
import speech_recognition as sr
import subprocess
import tkinter as tk
from tkinter import simpledialog, filedialog
from tkinter import ttk
import pyttsx3
import os
import pythoncom
import win32com.client
import threading
import time
import psutil  # Asegúrate de instalar psutil con pip install psutil
import pygetwindow as gw
from PIL import Image
from PIL import ImageTk
import win32gui  # Asegúrate de importar win32gui



class ComandosApp:
    def __init__(self):
        self.ventana = tk.Tk()
        self.ventana.title("VoiCo")
        self.ventana.attributes('-topmost', True)  # Mantener la ventana en la parte superior
        self.ventana.resizable(False, False)  # Deshabilitar la capacidad de cambiar el tamaño
        #self.ventana.overrideredirect(True)  # Eliminar el marco de la ventana
        #self.ventana.attributes('-alpha', 0.8)  # Establecer la transparencia (0.0 a 1.0)



        self.comandos = self.cargar_comandos()
        print("Comandos cargados al inicio:", self.comandos)

        self.etiqueta = tk.Label(self.ventana, text="Listo.")
        self.etiqueta.pack(pady=5)

        self.imagen_btn3 = tk.PhotoImage(file="agregar.png").subsample(20)  # Ajusta el factor según sea necesario
        self.boton_agregar = ttk.Button(self.ventana, text="Agregar Nuevo Comando", image=self.imagen_btn3, compound=tk.LEFT, command=self.agregar_comando)
        self.boton_agregar.pack(pady=1)

        self.imagen_btn4 = tk.PhotoImage(file="rutina.png").subsample(1)  # Ajusta el factor según sea necesario
        self.boton_agregar_rutina = ttk.Button(self.ventana, text="     Agregar Rutina     ", image=self.imagen_btn4, compound=tk.LEFT, command=self.agregar_rutina)
        self.boton_agregar_rutina.pack(pady=1)
   
        self.imagen_btn = tk.PhotoImage(file="escuchar.png").subsample(20)  # Ajusta el factor según sea necesario
        self.boton_mostrar = ttk.Button(self.ventana, text="Mostrar Comandos", image=self.imagen_btn, compound=tk.LEFT, command=self.mostrar_comandos)
        self.boton_mostrar.pack(pady=1)

        self.imagen_btn1 = tk.PhotoImage(file="dictar.png").subsample(20)  # Ajusta el factor según sea necesario
        self.boton_escuchar = ttk.Button(self.ventana, text="  Dictar Comando  ",  image=self.imagen_btn1, compound=tk.LEFT, command=self.escuchar_y_ejecutar)
        self.boton_escuchar.pack(pady=1)

        self.imagen_btn2 = tk.PhotoImage(file="escuchaactiva.png").subsample(20)  # Ajusta el factor según sea necesario
        self.boton_escucha_activa = ttk.Button(self.ventana, text="   Escucha Activa   ", image=self.imagen_btn2, compound=tk.LEFT, command=self.iniciar_escucha_activa)
        self.boton_escucha_activa.pack(pady=1)

        self.engine = pyttsx3.init()
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[0].id)
        self.engine.setProperty('rate', 150)  # Adjust the rate as needed

        self.hilo_escucha = None
        self.detener_escucha = threading.Event()

    def on_closing(self):
        self.guardar_comandos()
        self.detener_escucha.set()
        self.ventana.destroy()

    def cargar_comandos(self):
        try:
            with open('comandos.json', 'r') as file:
                return json.load(file)
        except FileNotFoundError:
            return {
                "abrir bloc de notas": "notepad.exe",
                "abrir facebook": "https://www.facebook.com/",
            }

    def guardar_comandos(self):
        with open('comandos.json', 'w') as file:
            json.dump(self.comandos, file, indent=2)

    def agregar_comando(self):
        nuevo_comando = simpledialog.askstring("Agregar Comando", "Ingrese el nuevo comando:")
        if nuevo_comando:
            tipo_comando = simpledialog.askstring("Agregar Comando", "Ingrese el tipo de comando (link/programa/carpeta):").lower()
            if tipo_comando == "link":
                nueva_aplicacion = simpledialog.askstring("Agregar Comando", "Ingrese la nueva URL:")
                if nueva_aplicacion:
                    self.comandos[nuevo_comando] = nueva_aplicacion
                    self.actualizar_etiqueta(f"Comando '{nuevo_comando}' agregado.")
                    self.guardar_comandos()
            elif tipo_comando == "programa":
                aplicacion = self.buscar_aplicacion()
                if aplicacion:
                    self.comandos[nuevo_comando] = aplicacion
                    self.actualizar_etiqueta(f"Comando '{nuevo_comando}' agregado.")
                    self.guardar_comandos()
            elif tipo_comando == "carpeta":
                carpeta = self.buscar_carpeta()
                if carpeta:
                    self.comandos[nuevo_comando] = carpeta
                    self.actualizar_etiqueta(f"Comando '{nuevo_comando}' agregado.")
                    self.guardar_comandos()
            else:
                self.actualizar_etiqueta("Tipo de comando no válido. Ingrese 'link', 'programa' o 'carpeta'.")

    def buscar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta")
        return carpeta    

    def agregar_rutina(self):
        nombre_rutina = simpledialog.askstring("Agregar Rutina", "Ingrese el nombre de la rutina:")
        if nombre_rutina:
            comandos_rutina = simpledialog.askstring("Agregar Rutina", "Ingrese los comandos de la rutina (separados por coma):")
            if comandos_rutina:
                comandos_lista = [cmd.strip() for cmd in comandos_rutina.split(',')]
                self.comandos[nombre_rutina] = comandos_lista
                self.actualizar_etiqueta(f"Rutina '{nombre_rutina}' agregada.")
                self.guardar_comandos()

    def mostrar_comandos(self):
        comandos_nombres = "\n".join(self.comandos.keys())
        mensaje = "Los comandos almacenados son:\n" + comandos_nombres
        self.actualizar_etiqueta(mensaje)

    def buscar_aplicacion(self):
        file_path = filedialog.askopenfilename(title="Seleccionar aplicación .exe", filetypes=[("Archivos ejecutables", "*.exe")])
        return file_path

    def actualizar_etiqueta(self, mensaje):
        self.etiqueta.config(text=mensaje)
        self.hablar(mensaje)

    def hablar(self, mensaje):
        self.engine.say(mensaje)
        self.engine.runAndWait()

    def escuchar_y_ejecutar(self):
        comando = self.escuchar_comando()
        if comando:
            self.ejecutar_comando(comando)

    def escucha_activa(self):
        recognizer = sr.Recognizer()

        while not self.detener_escucha.is_set():
            with sr.Microphone() as source:
                print("Escuchando...")
                recognizer.adjust_for_ambient_noise(source)

                audio = recognizer.listen(source)

            try:
                comando = recognizer.recognize_google(audio, language="es-ES").lower()
                print("Comando detectado:", comando)
                self.ejecutar_comando(comando)
            except sr.UnknownValueError:
                print("No se pudo entender el comando")
            except sr.RequestError as e:
                print(f"Error en la solicitud a Google API: {e}")

        print("Escucha detenida.")

    def iniciar_escucha_activa(self):
        if self.hilo_escucha and self.hilo_escucha.is_alive():
            self.detener_escucha.set()
            self.hilo_escucha.join()
            self.boton_escucha_activa.config(text="Escucha Activa")
        else:
            self.detener_escucha.clear()
            self.hilo_escucha = threading.Thread(target=self.escucha_activa)
            self.hilo_escucha.start()
            self.boton_escucha_activa.config(text="Detener Escucha")

    def escuchar_comando(self):
        recognizer = sr.Recognizer()
          # Delay for 2 seconds


        with sr.Microphone() as source:
            print("Di el comando:")
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)

        try:
            comando = recognizer.recognize_google(audio, language="es-ES").lower()
            print("Comando detectado:", comando)
            return comando
        except sr.UnknownValueError:
            print("No se pudo entender el comando")
            self.hablar("No se pudo entender el comando")
            return None
        except sr.RequestError as e:
            print(f"Error en la solicitud a Google API: {e}")
            return None

    def ejecutar_comando(self, comando):
        if comando.startswith("abrir ") and comando.endswith(" carpeta"):
            carpeta_a_abrir = comando.split("abrir ")[1].replace(" carpeta", "")
            self.abrir_carpeta(carpeta_a_abrir)        
        if comando.startswith("cerrar "):
            app_a_cerrar = comando.split("cerrar ")[1].lower()
            self.cerrar_aplicacion(app_a_cerrar)
        elif isinstance(self.comandos[comando], list):
            # Handle routine
            for subcomando in self.comandos[comando]:
                self.ejecutar_comando(subcomando)
        elif comando in self.comandos:
            aplicacion = self.comandos[comando]
            if aplicacion.endswith(".lnk"):
                ejecutable = self.obtener_ruta_desde_acceso_directo(aplicacion)
                if ejecutable:
                    os.startfile(ejecutable)
                    self.actualizar_etiqueta(f"{comando.capitalize()} abierto.")
                else:
                    self.actualizar_etiqueta(f"No se pudo abrir el acceso directo '{aplicacion}'.")
            elif aplicacion.startswith("http"):
                webbrowser.open(aplicacion)
                self.actualizar_etiqueta(f"{comando.capitalize()} abierto.")
            

            elif os.path.isdir(aplicacion):  # Nuevo bloque para abrir carpeta
                os.startfile(aplicacion)
                self.actualizar_etiqueta(f"Carpeta '{comando.capitalize()}' abierta.")


            else:
                subprocess.Popen([aplicacion], shell=True)
                self.actualizar_etiqueta(f"{comando.capitalize()} abierto.")
        else:
            self.actualizar_etiqueta("Comando no reconocido.")

    def abrir_carpeta(self, carpeta):
        if os.path.isdir(carpeta):
            os.startfile(carpeta)
            self.actualizar_etiqueta(f" '{carpeta}' abierta.")
        else:
            self.actualizar_etiqueta(f"No se encontró la carpeta '{carpeta}' para abrir.")


    def cerrar_aplicacion(self, nombre_aplicacion):
        nombre_aplicacion = nombre_aplicacion.lower()

        if nombre_aplicacion == "calculadora":
            try:
                os.system("taskkill /f /im CalculatorApp.exe")
                self.actualizar_etiqueta("Calculadora cerrada.")
            except Exception as e:
                print(f"Error al intentar cerrar Calculadora: {e}")
                self.actualizar_etiqueta("No se pudo cerrar Calculadora.")
    

        elif nombre_aplicacion == "word":
            try:
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Quit()
                self.actualizar_etiqueta(f"Microsoft Word cerrado.")
            except Exception as e:
                print(f"Error al intentar cerrar Microsoft Word: {e}")
                self.actualizar_etiqueta(f"No se pudo cerrar Microsoft Word.")


        elif nombre_aplicacion == "excel":
            try:
                word_app = win32com.client.Dispatch("Excel.Application")
                word_app.Quit()
                self.actualizar_etiqueta(f"Microsoft Excel cerrado.")
            except Exception as e:
                print(f"Error al intentar cerrar Excel Word: {e}")
                self.actualizar_etiqueta(f"No se pudo cerrar Microsoft Excel.")    

        elif nombre_aplicacion == "bloc de notas":
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'notepad.exe'], check=True)
                self.actualizar_etiqueta(f"Bloc de notas cerrado.")
            except subprocess.CalledProcessError as e:
                print(f"Error al intentar cerrar Bloc de notas: {e}")
                self.actualizar_etiqueta(f"No se pudo cerrar Bloc de notas.")

        elif nombre_aplicacion.startswith("youtube") or nombre_aplicacion.startswith("hora") or nombre_aplicacion.startswith("facebook"):
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'msedge.exe'], check=True)
                self.actualizar_etiqueta(f"Microsoft Edge cerrado.")
            except subprocess.CalledProcessError as e:
                print(f"Error al intentar cerrar Microsoft Edge: {e}")
                self.actualizar_etiqueta(f"No se pudo cerrar Microsoft Edge.")

        elif nombre_aplicacion == "lolero":
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'RiotClientServices.exe'], check=True, cwd="C:\\Riot Games\\Riot Client")
                self.actualizar_etiqueta(f"La aplicación 'lolero' ha sido cerrada.")
            except subprocess.CalledProcessError as e:
                print(f"Error al intentar cerrar 'lolero': {e}")
                self.actualizar_etiqueta(f"No se pudo cerrar 'lolero'.")

        else:
            self.actualizar_etiqueta(f"No se encontró la aplicación '{nombre_aplicacion}' para cerrar.")

 

    def obtener_ruta_desde_acceso_directo(self, acceso_directo):
        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(acceso_directo)
            return os.path.abspath(shortcut.Targetpath)
        except Exception as e:
            print(f"Error al obtener la ruta desde el acceso directo: {e}")
            return None

    def run(self):
        self.ventana.protocol("WM_DELETE_WINDOW", self.guardar_comandos)
        self.ventana.mainloop()


    def on_closing(self):
        self.guardar_comandos()
        self.detener_escucha.set()
        self.ventana.destroy()

    def run(self):
        self.ventana.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.ventana.mainloop()

if __name__ == "__main__":
    app = ComandosApp()
    app.run()
