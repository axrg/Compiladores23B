# ComandosApp

ComandosApp es una aplicación interactiva en Python que combina una interfaz gráfica de usuario (GUI) con reconocimiento de voz para permitir a los usuarios ejecutar comandos mediante comandos hablados. La aplicación ofrece funciones para agregar, gestionar y ejecutar comandos, así como otras características como el cierre de aplicaciones específicas y retroalimentación auditiva.

## Características Principales

- **Reconocimiento de Voz:** Utiliza la biblioteca `speech_recognition` para reconocer comandos hablados.

- **Interfaz Gráfica de Usuario (GUI):** Implementa una interfaz gráfica intuitiva con botones para diversas acciones.

- **Gestión de Comandos:** Permite agregar nuevos comandos y rutinas, ejecutando acciones como abrir programas, enlaces web o carpetas.

- **Síntesis de Voz:** Proporciona retroalimentación auditiva a través de la biblioteca `pyttsx3`.

- **Escucha Activa Continua:** Permite la detección continua de comandos de voz en tiempo real.

- **Persistencia de Datos:** Almacena y carga comandos desde un archivo JSON (`comandos.json`).

- **Cierre de Aplicaciones:** Puede cerrar aplicaciones específicas como parte de los comandos.

## Instalación

1. Clona o descarga el repositorio.

    ```bash
    git clone https://github.com/axrg/ComandosApp.git
    cd ComandosApp
    ```

2. Instala las dependencias utilizando `pip`:

    ```bash
    pip install -r requirements.txt
    ```

3. Ejecuta la aplicación:

    ```bash
    python ComandosApp.py
    ```

## Configuración y Uso

- **Agregar Comandos:** Utiliza la interfaz gráfica para agregar nuevos comandos, especificando el tipo (link, programa, carpeta).

- **Ejecutar Comandos:** Dicta comandos de voz para ejecutar acciones. Puedes cerrar aplicaciones específicas también.

- **Escucha Activa:** Inicia la escucha continua para detectar comandos de voz en tiempo real.

- **Guardar Comandos:** Los comandos se guardan automáticamente en `comandos.json` al cerrar la aplicación.

## Contribución

Si encuentras errores o tienes sugerencias de mejora, no dudes en crear un [issue](https://github.com/axrg/ComandosApp/issues) o enviar un [pull request](https://github.com/axrg/ComandosApp/pulls).

## Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo [LICENSE](LICENSE) para más detalles.
