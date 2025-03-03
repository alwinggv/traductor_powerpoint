# Traductor de Presentaciones PowerPoint

Este proyecto permite traducir automáticamente el contenido de presentaciones PowerPoint de un idioma a otro utilizando dos métodos de traducción:
1. Google Translate (a través de la biblioteca deep-translator)
2. OpenAI (ChatGPT) para traducciones de mayor calidad

## Requisitos

- Python 3.6 o superior
- Entorno virtual (venv)
- Clave de API de OpenAI (solo si se desea usar el método de OpenAI)

## Dependencias

El proyecto utiliza las siguientes bibliotecas:
- python-pptx: Para leer y modificar archivos PowerPoint
- deep-translator: Para realizar las traducciones con Google Translate
- openai: Para realizar traducciones con la API de OpenAI
- python-dotenv: Para gestionar variables de entorno
- tqdm: Para mostrar barras de progreso durante la traducción

## Estructura del proyecto

El proyecto está organizado de la siguiente manera:
- `originales/`: Carpeta donde se almacenan los archivos PowerPoint originales
- `traducidos/`: Carpeta donde se guardan los archivos PowerPoint traducidos
- `traductor_ppt.py`: Script principal para traducir presentaciones
- `requirements.txt`: Lista de dependencias del proyecto
- `.env`: Archivo para almacenar la clave de API de OpenAI (debes crearlo)

## Configuración del entorno

1. Clonar o descargar este repositorio
2. Activar el entorno virtual:
   ```
   .\venv\Scripts\activate
   ```
3. Las dependencias ya están instaladas en el entorno virtual
4. Si deseas usar OpenAI, crea un archivo `.env` en la raíz del proyecto con el siguiente contenido:
   ```
   OPENAI_API_KEY=tu_clave_de_api_aqui
   ```

## Uso

### Flujo de trabajo manual

1. Coloca tus archivos PowerPoint en la carpeta `originales/`
2. Ejecuta el script principal:
   ```
   python traductor_ppt.py "originales/nombre_de_tu_archivo.pptx" idioma_destino idioma_origen [metodo_traduccion]
   ```
3. Los archivos traducidos se guardarán automáticamente en la carpeta `traducidos/`

Ejemplos:
```
python traductor_ppt.py "originales/mi_presentacion.pptx" en es
```
Esto traducirá la presentación del español (es) al inglés (en) usando Google Translate (método por defecto).

```
python traductor_ppt.py "originales/mi_presentacion.pptx" fr es openai
```
Esto traducirá la presentación del español (es) al francés (fr) usando la API de OpenAI para obtener traducciones de mayor calidad.

### Métodos de traducción disponibles

- `google`: Utiliza Google Translate a través de la biblioteca deep-translator (método por defecto)
- `openai`: Utiliza la API de OpenAI (ChatGPT) para obtener traducciones de mayor calidad

### Códigos de idioma comunes

- es: Español
- en: Inglés
- fr: Francés
- de: Alemán
- it: Italiano
- pt: Portugués
- ja: Japonés
- zh-CN: Chino (Simplificado)
- ru: Ruso

### Organización de archivos

- Los archivos originales se deben colocar en la carpeta `originales/`
- Los archivos traducidos se guardarán automáticamente en la carpeta `traducidos/`

## Notas

- Este proyecto es para uso educativo y personal.
- Las traducciones con OpenAI suelen ser de mayor calidad pero requieren una clave de API válida y tienen un costo asociado.
- Las traducciones con Google Translate son gratuitas pero pueden ser menos precisas, especialmente para textos técnicos o especializados. 