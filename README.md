# Proyecto Ciclovia Autodocs

## Descripción
Esta software genera cronogramas de actividades mensual en formato de hojas de cálculo a partir de información recolectada a través de formularios en línea. El proyecto optimizó la recolección de datos de contratistas y simplificó la generación de documentación necesaria para la validación de cuentas de cobro.

Este repositorio contiene dos proyectos relacionados:
- **Auto-cronograma**: Una biblioteca que proporciona funcionalidades para generar cronogramas.
- **Formulario-sheets**: Un proyecto que se ejecuta desde una hoja de cálculo de Google Sheets y utiliza la biblioteca Auto-cronograma.

## Estructura del Proyecto
    Auto-cronograma/ src/ 
        .clasp.json 
        appsscript.json 
        globals.js main.js 

    Formulario-sheets/ src/ 
        .clasp.json 
        appsscript.json 
        Código.js 
        Globals.js 
    package.json


## Despliegue

### Auto-cronograma

1. Navega a la carpeta `Auto-cronograma/src`.
2. Ejecuta `clasp push` para desplegar los cambios.

### Formulario-sheets

1. Navega a la carpeta `Formulario-sheets/src`.
2. Ejecuta `clasp push` para desplegar los cambios.

## Uso

1. Abre la hoja de cálculo de Google Sheets asociada con `Formulario-sheets`.
2. Utiliza el menú personalizado "Autofill Form Ciclovida" para ejecutar las funciones disponibles.

Link al proyecto Apps Scripts, Formularios y formatos en Drive - 
https://drive.google.com/drive/folders/1DEJwIOkSNp_e4JuthQHZIwvHaMQ7R1jY?usp=sharing