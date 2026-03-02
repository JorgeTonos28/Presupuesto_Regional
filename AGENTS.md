# Instrucciones para Agentes de Desarrollo

## Versionado y Control de Cambios

1.  **Versión de la App:** Siempre incrementar la constante `APP_VERSION` en `Código.js` cada vez que se realicen modificaciones funcionales o correcciones de bugs.
2.  **Mensajes de Commit:** Al hacer commit, mencionar la versión actualizada en el mensaje (ej: `Bump v1.0.2: ...`). Esto facilita el rastreo de cambios.
3.  **Logs de Servidor:** Utilizar `Logger.log()` para errores críticos, especialmente en `doGet`, `apiBootstrap` y `api...`. Estos logs son vitales para depurar errores "silenciosos" del lado del cliente.

## Estilo de Código y Patrones

1.  **HTML/CSS:**
    *   Mantener el CSS en `styles.html` usando variables CSS (`:root`).
    *   Utilizar `normalizeDriveUrl_` en el servidor para cualquier URL de imagen externa, asegurando compatibilidad con los permisos de Drive.
    *   Evitar inyectar estilos inline salvo excepciones muy específicas.
2.  **JavaScript (Cliente):**
    *   Usar `async/await` para llamadas al servidor (`callServer`).
    *   Manejar SIEMPRE errores con `try...catch` y mostrar feedback al usuario mediante `modal()`.
    *   Usar `Busy.show()` antes de operaciones asíncronas y `Busy.hide()` en `finally`.
3.  **JavaScript (Servidor):**
    *   Retornar objetos estandarizados `{ ok: boolean, data: any, message: string }`.
    *   Validar permisos (`requireAdmin_`) al inicio de funciones sensibles.
    *   Usar `SpreadsheetApp.flush()` después de operaciones de escritura (crear, borrar, actualizar) para asegurar consistencia inmediata en lecturas subsecuentes.

## Configuración y Dependencias

*   **appsscript.json:** Es la fuente de verdad para los servicios habilitados. Si se requiere Drive API u otros servicios avanzados, deben estar declarados en `dependencies.enabledAdvancedServices`. No eliminar ni modificar sin razón justificada.
*   **Hoja Config:** Los parámetros globales (`logo_url`, `locale`) se leen de esta hoja. Al añadir nuevas configuraciones, actualizar `ensureConfigDefaults_` en `setup()`.
