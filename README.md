# RevisaBitacoras

Aplicación Streamlit para validar y completar las hojas **Flujo Nuevo** y
**Flujo Viejo** de la bitácora diaria. Usa el `exported-logs.csv` del PMS para
estados y devoluciones, y el `EXPORT_*.xlsx` de SAP para las cantidades P/TRATAS.
El archivo final se devuelve codificado con la plantilla HSEQ/SIG **F-AGP-04**,
incluida en `assets/` para que funcione tanto localmente como en Streamlit Cloud.
La hoja `PNC` se conserva como registro operativo, se limpia y se traslada a la
plantilla codificada sin validaciones adicionales por ahora.

## Flujo de auditoría guiada

La aplicación conserva el expediente en la sesión y obliga a resolver cada etapa
antes de avanzar:

1. **Estados:** presenta únicamente las diferencias, una por una. El auditor puede
   conservar la bitácora, corregir el estado y sus datos, o excluir la fila. Las
   devoluciones (`verify_formats`, `2.6.1-Dibujo en revision` y `2-Valorizado`) no solicitan AGP ni
   protocolo y recuperan automáticamente el comentario del PMS. AGP y protocolo se
   solicitan al corregir a un estado de dibujo finalizado. Para toda salida de dibujo
   que pase por auditoría se muestran y confirman también la fecha/hora de llegada y
   de salida.
2. **Casos no reportados:** permite agregar cada hallazgo a `Flujo Nuevo` o
   `Flujo Viejo`, o justificar por qué no debe agregarse. Pide AGP y protocolo,
   pero deja el conteo P/TRATA para la etapa siguiente. Si el hallazgo es una
   devolución (`verify_formats` o `2.6.1-Dibujo en revision`), no solicita AGP ni
   protocolo y escribe directamente la devolución recuperada del PMS. El origen de
   actualización se elige de una lista cerrada con los valores operativos autorizados.
3. **Protocolos y cantidades:** completa primero protocolos vacíos, mal separados o
   que no comienzan por `1000`,
   genera un Excel para solicitar el EXPORT de SAP y solo entonces habilita su
   carga. La hoja `Protocolos` contiene una tabla de una sola columna con protocolos
   únicos; la hoja `Detalle` conserva su relación con hoja, QR y estado. Los
   protocolos encontrados y faltantes se presentan por separado. Cada faltante se
   revisa en una tarjeta copiable y permite: corregir el código en todas sus
   referencias, registrar cantidades PEL/TRATA manuales con justificación e imagen
   obligatoria, o marcarlo como `No aplica` con justificación.
4. **Excel final:** entrega la bitácora codificada lista para cargar y agrega la hoja
   `Auditoria guiada`, donde queda registrada cada decisión, estado, AGP, protocolo
   y justificación. Cuando hubo soportes manuales agrega además `Soportes protocolos`,
   con la imagen incrustada dentro del mismo libro.

## Regla aplicada

1. La fecha de producción se lee del nombre de la bitácora (`DD-MM-AAAA`).
2. Por cada QR se reconstruyen ciclos de dibujo. En flujo nuevo la llegada es el
   primer `drawing_request` o `assigned_draftsman` posterior a la salida previa; en
   flujo viejo es el primer `2.5-Enviado a dibujo` o `2.6-Dibujante asignado`.
   Cualquier salida real anterior cierra el ciclo, aunque su actor no pertenezca a
   la lista autorizada que se usa exclusivamente para `Casos no reportados`.
3. La salida debe ser un movimiento real hecho por el mismo profesional que figura
   en `USUARIO`: `verify_formats`/`drawing_review` para flujo nuevo y
   `2-Valorizado`/`2.6.1-Dibujo en revision`/`2.7-Dibujo finalizado` para flujo viejo.
   Las filas técnicas duplicadas y las actualizaciones sin cambio de estado se
   ignoran. `1.3-RG Gestionado EEC` nunca se considera salida de dibujo.
4. Antes de validar, se comprueba la cobertura real de `date_created` dentro del
   CSV. Si el log termina antes de la fecha de producción, el proceso se detiene y
   solicita un `exported-logs` más reciente; el prefijo de fecha del nombre no se usa.
5. Se compara contra `ESTADO DE LA OT`. Si el estado existe pero lo realizó un
   profesional diferente al reportado, el caso pasa obligatoriamente por auditoría.
   La comparación tolera apellidos omitidos por PMS cuando conserva al menos tres
   componentes inequívocos del nombre, además de orden, puntos, tildes y mayúsculas.
6. Por defecto, el Excel final conserva solo las filas que coinciden.
7. Las filas válidas se copian a la plantilla fija codificada `F-AGP-04`.
8. El mismo Excel agrega al final la hoja `Resultado Validacion` con el detalle de
   coincidencias, diferencias, cantidades, devoluciones y trazabilidad temporal.
   Las hojas `Flujo Nuevo` y `Flujo Viejo` reciben las columnas
   `FECHA/HORA DE LLEGADA` y `FECHA/HORA DE SALIDA` como fechas reales de Excel.
9. También agrega la hoja `Casos no reportados`: incluye los QR ausentes de ambas
   hojas de la bitácora que tuvieron ese día un evento `2.7-Dibujo finalizado`,
   `2.6.1-Dibujo en revision`, `drawing_review` o `verify_formats`. Si un QR tuvo más
   de uno, se conserva únicamente el evento objetivo más reciente del día. Para los
   estados cerrados `2.7-Dibujo finalizado` y `drawing_review`, solo se incluyen los
   movimientos hechos por la lista autorizada de profesionales. El usuario se
   identifica desde `username` o `user_created_last_name` por cédula, cédula con
   prefijo `CO` o nombre sin depender del orden, puntos, tildes o mayúsculas. Los
   estados de devolución no llevan este filtro de profesional.
10. El logo OCA Global se re-inserta desde la plantilla en los encabezados codificados.
11. Para `2.7-Dibujo finalizado` y `drawing_review`, se buscan los protocolos en
   `Sheet1` del EXPORT y se cuentan los `GIS_GUID` únicos por tipo `P` y `T`.
12. Para `2-Valorizado`, `2.6.1-Dibujo en revision` y `verify_formats`, se copia a `OBSERVACIONES`
    el último `comments` no vacío del mismo estado en el PMS. La búsqueda se hace
    por QR, estado y profesional hasta la salida seleccionada: aunque después aparezca
    otro movimiento del caso, se conserva la devolución más reciente que coincida con
    el estado aprobado en la bitácora.
13. Los marcadores `N/A`, `No aplica`, `S/M` y `SM` se eliminan de
    `OBSERVACIONES`; los demás textos se conservan.
14. En las cinco hojas operativas (`Flujo Nuevo`, `Flujo Viejo`, `Odm`,
    `Correo-Rotulacion` y `PNC`) se eliminan de cualquier columna los valores
    aislados `N/A`, `S/N` y `0`. En `PNC` también se limpian los marcadores aislados
    `No aplica`, `S/M` y `SM`. Los textos que solamente contienen esas secuencias
    como parte de información válida y las fórmulas no se eliminan. La fórmula de
    `TOTAL ELEMENTOS` devuelve vacío cuando ambas cantidades están vacías.
15. En `Correo-Rotulacion`, únicamente las filas cuyo `ORIGEN DE ACTUALIZACIÓN` sea
    `Solicitud por correo` y cuyo `ESTADO DE LA OT` sea `Finalizado` se cruzan contra
    el EXPORT y reciben las cantidades PEL/TRATAS de todos sus protocolos. Las solicitudes
    en otros estados y las filas con origen `Rotulación` quedan excluidas. Los
    resultados y errores se detallan en `Validacion Correos`.
16. `PNC` se copia al formato F-AGP-04 con sus diez columnas, desde `FECHA DE
    LLEGADA` hasta `FECHA DE REVISIÓN`. Sus registros de origen `Correo` o `PNC`
    no participan todavía en la validación de estados, casos no reportados ni
    cantidades P/TRATA.

Si una celda contiene varios protocolos, deben estar separados obligatoriamente por
punto y coma (`1000162837;1000162838`). Las cantidades se escriben en el mismo orden:
por ejemplo, PEL `1;0` y TRATAS `1;1`. `TOTAL ELEMENTOS` sigue siendo un número y
corresponde a la suma de todos los valores (`3` en el ejemplo). Los protocolos
faltantes quedan señalados como `Protocolos incompletos en EXPORT`; el reporte
muestra el subtotal de los encontrados hasta que el auditor corrija el código,
registre cantidades manuales o justifique que no aplica. Si se detectan varios
protocolos sin `;`, la fila queda marcada con `Separador inválido entre protocolos: usa ;`.
Antes de generar el Excel para SAP, la etapa 3 bloquea el avance y solicita corregir
protocolos separados con comas, barras verticales, saltos de línea, espacios o
concatenados sin `;`. Cada protocolo debe comenzar por `1000`; cualquier otro
prefijo también bloquea el avance y solicita revisión. Después de cargar el EXPORT,
un protocolo ausente puede sustituirse por su valor correcto —o por varios valores
separados por `;`— y la aplicación actualiza todas las referencias antes de repetir
el cruce P/TRATA. Las cantidades justificadas manualmente aceptan cero y requieren
una imagen de respaldo; la justificación y el soporte se conservan en el Excel final.

Cuando un caso tiene varios códigos AGP, también deben registrarse con punto y coma
(`2000133299;2000134328`). La interfaz rechaza comas, saltos de línea y listas de AGP
detectadas sin `;`.

Los estados posteriores a la fecha de producción no alteran una validación histórica.
Cuando un QR no tuvo movimiento en ese día, la aplicación muestra también el último
estado anterior disponible como contexto, pero no lo considera una coincidencia.

## Ejecución

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

Para iniciar, la interfaz solo solicita la bitácora y `exported-logs.csv`. El
`EXPORT_*.xlsx` se carga después, en la etapa de protocolos, cuando ya se descargó
el listado depurado. Su cargador permanece bloqueado durante el inicio y las etapas
1 y 2, se habilita exclusivamente en la etapa 3 y vuelve a bloquearse al cerrar las
cantidades. Acepta CSV grandes (hasta 500 MB), procesa el log por bloques y
toma automáticamente la plantilla F-AGP-04 incluida en `assets/`. En ejecución local
se conserva la antigua ruta de Drive únicamente como respaldo.

## Despliegue en Streamlit Community Cloud

1. Publica este repositorio en GitHub. Es recomendable mantenerlo **privado**, pues
   incluye la plantilla operativa F-AGP-04.
2. En Streamlit Community Cloud selecciona el repositorio, la rama `main` y
   `app.py` como archivo principal.
3. No se requieren secretos. Las dependencias se instalan desde `requirements.txt`
   y `runtime.txt` fija Python 3.12.
4. El límite de carga está configurado en 500 MB mediante `.streamlit/config.toml`.

Los archivos operativos, exportaciones, resultados, entornos virtuales y secretos
locales están excluidos mediante `.gitignore` y no deben subirse al repositorio.
