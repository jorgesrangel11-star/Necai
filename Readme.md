# 🌿 Necai - Sistema de Control Operativo y Digitalización de Procesos

Sistema de gestión y optimización de procesos administrativos desarrollado para la empresa **Necai Sistemas de Riego y Jardinería**, enfocado en la digitalización de reportes de campo y el control operativo de mantenimiento de áreas verdes.

## 🚀 Características Principales

* **Digitalización de Reportes:** Captura y registro de datos en campo por parte de supervisores desde dispositivos móviles.
* **Gestión de Evidencias Fotográficas:** Organización y almacenamiento automático de imágenes en Google Drive con compresión optimizada.
* **Control de Inventario en Tiempo Real:** Tablero de control dinámico con inventario activo e historial cronológico de mantenimiento.
* **Centralización de Datos:** Integración eficiente para el procesamiento y consulta centralizada de reportes[cite: 1].

## 📂 Estructura y Automatización en Google Drive

El backend en Google Apps Script (`codigo.gs`) automatiza la creación jerárquica de carpetas y almacenamiento de evidencias fotográficas directamente en Google Drive[cite: 1]:

```text
📁 Necai - Control Operativo (Google Drive)
 ├── 📁 Reportes_[Año]
 │    ├── 📁 [Mes]
 │    │    └── 📁 Reporte_[Fecha]_[Supervisión]
 │    │         ├── 🖼️ Evidencia_Antes.jpg
 │    │         ├── 🖼️ Evidencia_Durante.jpg
 │    │         └── 🖼️ Evidencia_Despues.jpg