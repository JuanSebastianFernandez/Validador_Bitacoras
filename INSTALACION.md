# Instalación para el usuario

El instalador está en `release/RevisorBitacoras-Setup-1.0.0.exe`. La persona que
va a usar el sistema no necesita instalar Python, Streamlit ni dependencias, y la
instalación se realiza únicamente para su usuario de Windows.

## Paso a paso

1. Copia `RevisorBitacoras-Setup-1.0.0.exe` al computador donde se usará el
   sistema. No lo ejecutes directamente desde un correo o una carpeta comprimida.
2. Haz doble clic en el instalador.
3. Si Windows SmartScreen muestra **Windows protegió su PC**, confirma que el
   archivo proviene de la persona responsable del sistema, pulsa **Más
   información** y luego **Ejecutar de todas formas**. El aviso aparece porque
   esta versión interna no tiene un certificado comercial de firma digital.
4. Conserva marcada la opción de crear un acceso directo y pulsa **Instalar**.
5. Al finalizar, pulsa **Terminar**. La aplicación abrirá automáticamente el
   navegador predeterminado.

## Uso diario

- Abre **Revisor de Bitácoras** desde el escritorio o el menú Inicio.
- Mientras esté activo aparecerá un icono verde en el área de notificaciones,
  junto al reloj. Si cierras la pestaña del navegador, haz clic derecho en ese
  icono y selecciona **Abrir en el navegador**.
- Para cerrar completamente el sistema, haz clic derecho en el icono verde y
  selecciona **Cerrar aplicación**.
- El programa se ejecuta solo en el computador local y no publica el servidor en
  la red. Los archivos se cargan desde el navegador y el resultado se descarga en
  la carpeta elegida por el usuario.

## Desinstalación

En Windows abre **Configuración > Aplicaciones > Aplicaciones instaladas**, busca
**Revisor de Bitácoras** y selecciona **Desinstalar**.

## Si no abre

Revisa el archivo `%LOCALAPPDATA%\RevisaBitacoras\app.log`. Si la empresa bloquea
programas internos sin firma digital, el área de TI deberá autorizar el instalador
o firmarlo con el certificado corporativo.

