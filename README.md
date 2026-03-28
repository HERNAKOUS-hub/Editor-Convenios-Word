# 📄 Editor Automático de Convenios Word

Una aplicación web desarrollada con **Spring Boot** para automatizar la edición y cumplimentación de documentos Word (`.docx`), especialmente diseñada para convenios oficiales y documentos legales con tablas anidadas.

La aplicación escanea el documento, detecta de forma inteligente los huecos a rellenar (marcadores en rojo, "NNNN", "xxxx", etc.), muestra una interfaz web limpia con el contexto de cada campo para que el usuario introduzca los datos, y genera un documento final limpio y listo para imprimir.

## ✨ Características Principales

* **Detección Inteligente de Marcadores:** Identifica campos a rellenar basándose en el color de la fuente (rojos puros o de Office) y en patrones de texto (`NNNN`, `xxxx`, espacios especiales).
* **Escáner Profundo Recursivo:** Capaz de encontrar marcadores ocultos dentro de tablas anidadas (tablas dentro de tablas), muy comunes en documentos oficiales de la Administración.
* **Filtro Legal Avanzado:** Ignora automáticamente párrafos legales largos que estén en rojo por motivos informativos, capturando únicamente las variables reales a modificar.
* **Interfaz Web con Contexto:** Genera formularios dinámicos mostrando la frase completa (contexto) donde se encuentra el marcador, facilitando enormemente la cumplimentación.
* **Limpieza de Formato Automática:** El documento final generado se descarga con los textos reemplazados en color negro estándar y sin resaltados, integrándose perfectamente en el documento original.

## 🛠️ Tecnologías Utilizadas

* **Java 21**
* **Spring Boot 3** (Web, MVC)
* **Apache POI** (Manipulación avanzada de archivos `.docx` y XML interno)
* **Thymeleaf** (Motor de plantillas HTML para vistas dinámicas)
* **HTML5 / CSS3** (Diseño frontend responsivo y sin dependencias externas)
* **Maven** (Gestión de dependencias y empaquetado)

## 🚀 Cómo ejecutar la aplicación

Tienes dos formas de utilizar esta aplicación:

### Opción 1: Usar la versión compilada (Recomendado para usuarios finales)
1. Ve a la sección **Releases** de este repositorio y descarga el archivo `cambiarword-X.X.X.jar`.
2. Asegúrate de tener Java instalado en tu equipo.
3. Haz doble clic sobre el archivo `.jar` (o ejecútalo desde la terminal con `java -jar cambiarword-X.X.X.jar`).
4. Abre tu navegador web y entra en `http://localhost:8080`.

### Opción 2: Compilar desde el código fuente (Para desarrolladores)
1. Clona este repositorio:
   ```bash
   git clone [https://github.com/TU_USUARIO/TU_REPOSITORIO.git](https://github.com/TU_USUARIO/TU_REPOSITORIO.git)
