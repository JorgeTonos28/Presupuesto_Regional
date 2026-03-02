# App Segmentación de Presupuesto Regional

Aplicación web desarrollada en Google Apps Script para gestionar la segmentación del presupuesto regional. Permite visualizar el presupuesto base vs. lo segmentado, aplicar nuevas segmentaciones proporcionales sobre el monto disponible, y consultar detalles por regional.

## Características Clave

*   **Autenticación Institucional:** Acceso restringido a usuarios del dominio configurado.
*   **Gestión de Segmentaciones:** Creación de segmentaciones por porcentaje, afectando a todas las regionales proporcionalmente según su disponible.
*   **Dashboard en Tiempo Real:** Visualización de totales, disponible y listas de segmentaciones.
*   **Detalle por Regional:** Vista detallada con histórico de movimientos por regional.
*   **Administración de Usuarios:** Panel para gestionar roles (ADMIN / COLAB) y estados de acceso.
*   **Indicadores de Carga:** Feedback visual (spinner) para todas las operaciones asíncronas.
*   **Diseño Responsivo:** Adaptable a móviles y escritorio.

## Estructura del Repositorio

*   **Código.js:** Lógica del servidor (backend), API, manejo de hojas de cálculo y autenticación.
*   **Index.html:** Plantilla principal que orquesta la carga de vistas y recursos.
*   **app.html:** Estructura base de la UI (topbar, footer, modales).
*   **styles.html:** Estilos CSS (variables, grid, componentes).
*   **js.html:** Lógica del cliente (frontend), manejo de estado, llamadas al servidor y renderizado.
*   **view_dashboard.html, view_regional.html, view_admin.html:** Vistas parciales de la aplicación.
*   **Denied.html:** Pantalla de acceso denegado.
*   **appsscript.json:** Manifiesto del proyecto (configuración de dependencias y servicios).
*   **AGENTS.md:** Instrucciones para desarrolladores y control de versiones.

## Requisitos Previos

1.  **Cuenta Google Workspace:** Necesaria para el despliegue y acceso.
2.  **Hoja de Cálculo:** Debe existir una Google Sheet con las siguientes pestañas (el script `setup()` las crea si no existen):
    *   `Config`: Parámetros globales.
    *   `Usuarios`: Lista de usuarios y roles.
    *   `BaseDatos`: Datos presupuestarios origen (Año, Regional, Presupuesto, etc.).
    *   `Segmentaciones`: Cabeceras de segmentaciones creadas.
    *   `SegmentacionesDetalle`: Desglose por regional de cada segmentación.
3.  **Servicios Avanzados:**
    *   **Drive API (v2):** Debe estar habilitada en el editor de Apps Script (Servicios +) y en la consola de Google Cloud del proyecto (si aplica). Esto es crucial para visualizar logos e imágenes alojadas en Drive.

## Configuración Inicial (Setup)

1.  **Crear Proyecto:** Crear un nuevo proyecto de Apps Script vinculado a una Hoja de Cálculo o independiente.
2.  **Copiar Archivos:** Subir los archivos `.html` y `.js` de este repositorio al proyecto.
3.  **Habilitar Drive API:**
    *   En el editor, clic en `+` al lado de "Servicios".
    *   Seleccionar **Drive API**, versión **v2**. Identificador: `Drive`.
4.  **Ejecutar Setup:**
    *   Abrir `Código.js`.
    *   Ejecutar la función `setup()`. Esto creará las hojas necesarias y registrará al usuario actual como ADMIN.
5.  **Configurar Logos (Opcional):**
    *   Subir el logo y firma a Google Drive.
    *   Asegurar que los archivos tengan permiso de lectura ("Cualquier usuario con el enlace").
    *   Copiar el **ID del archivo** (la parte alfanumérica de la URL).
    *   Ir a la hoja `Config` y pegar los IDs en `logo_url` y `signature_url`.
    *   Opcionalmente, ajustar `logo_width`, `logo_height`, `signature_width`, `signature_height` (ej: `100px`, `auto`).
6.  **Desplegar:**
    *   Clic en "Implementar" > "Nueva implementación".
    *   Tipo: "Aplicación web".
    *   Ejecutar como: "Usuario que implementa la aplicación".
    *   Quién tiene acceso: "Cualquier usuario de [Tu Organización]".

## Uso de la Aplicación

*   **Acceso:** Ingresar a la URL proporcionada tras el despliegue.
*   **Dashboard:** Seleccionar el año para ver el estado del presupuesto. Los Administradores verán el formulario para crear nuevas segmentaciones.
*   **Regional:** Consultar el detalle de una regional específica.
*   **Administración:** (Solo Admin) Agregar o modificar permisos de usuarios.

## Solución de Problemas (Troubleshooting)

*   **Error "Cannot read properties of null (reading 'ok')":**
    *   Generalmente ocurre si el script del servidor falla silenciosamente o retorna nulo. Revisar los logs de ejecución en el editor.
    *   Asegurar que la sesión de usuario es válida y que el navegador no bloquea cookies de terceros necesarias para la autenticación de Apps Script.
*   **Imágenes Rotas (Logo/Firma):**
    *   Verificar que el ID en la hoja `Config` sea correcto.
    *   Verificar que el archivo en Drive sea **Público** o compartido con la organización.
    *   Asegurar que el servicio **Drive API** esté habilitado en el proyecto.

## Contacto

Desarrollado por el equipo de Innovación.
