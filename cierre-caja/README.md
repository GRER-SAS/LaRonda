# Cierre de Caja — GRER S.A.S.

Aplicación web para el cierre de caja de los cinco hoteles administrados por GRER S.A.S. en Cuenca, Ecuador.

## Hoteles incluidos
- Azul de la Plaza
- Chat Noir
- La Ronda
- Magdalena
- Sanchez

## Regla principal
El cierre solo se acepta cuando la diferencia en efectivo es exactamente **$0,00**.

**Efectivo esperado = saldo anterior + ingresos en efectivo - egresos**

Transferencias y tarjetas se registran para control, pero no forman parte del efectivo físico de caja.

El sistema trabaja internamente en centavos enteros para evitar errores de redondeo.

## Instalación
1. Abrir el Google Sheet indicado por GRER S.A.S.
2. Abrir `Extensiones → Apps Script`.
3. Crear un proyecto de Apps Script.
4. Copiar el contenido de `Code.gs` y `Index.html` del repositorio.
5. Ejecutar una vez `setup()` y autorizar el acceso al Spreadsheet.
6. En Apps Script elegir `Implementar → Nueva implementación → Aplicación web`.
7. Ejecutar como el propietario del proyecto.
8. Elegir quién tiene acceso según la política interna de GRER S.A.S.
9. Compartir la URL de la aplicación con recepción.

## Hoja de cálculo
El backend usa el Spreadsheet ID proporcionado en `Code.gs`. Si la pestaña `Cierres de Caja` no existe, `setup()` la crea automáticamente.

## Seguridad recomendada
La URL de la aplicación debe compartirse solo con personal autorizado. Para una versión productiva se recomienda añadir autenticación por usuario corporativo, permisos por hotel y un panel administrativo con auditoría.

## Próximas mejoras posibles
- Usuario/login por recepcionista.
- Selección de denominaciones de billetes y monedas para validar el conteo físico.
- Número de turno automático.
- Firma digital del recepcionista y supervisor.
- Bloqueo de cierres duplicados por hotel/fecha/turno.
- Panel gerencial con filtros por hotel, fecha, recepcionista y diferencias.
- Exportación PDF del cierre.
- Registro de anulaciones o correcciones con auditoría.
