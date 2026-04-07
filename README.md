# 📊 Email Campaign Analytics Dashboard

Sistema unificado de dashboards para monitoreo de campañas de email marketing de múltiples compañías.

## 🌐 [Ver Dashboard en Vivo](https://tu-usuario.github.io/email-dashboard/)

*(Reemplaza esta URL después de publicar en GitHub Pages)*

---

## 🏢 Compañías Incluidas

- **TeamFicient** - Azul
- **AccuSights** - Morado/Violeta
- **HireALatino** - Azul Cielo
- **MyDrive Academy** - Verde
- **iDrivio** - Naranja
- **ArchFicient** - Azul Oscuro/Dorado
- **DocFicient** - Violeta

---

## ✨ Características

- 📈 **KPIs en Tiempo Real**: Total leads, open rate, clicks, y replies
- 📊 **Visualizaciones Interactivas**: Gráficos de línea, barras y engagement
- 🎨 **Diseño Personalizado**: Cada compañía con su esquema de colores único
- 📱 **Responsive**: Funciona perfectamente en desktop, tablet y móvil
- 🔍 **Filtros Avanzados**: Por campaña, industria y rango de fechas
- 💬 **Sistema de Replies**: Visualización detallada de respuestas categorizadas

---

## 🚀 Cómo Actualizar los Dashboards

### Requisitos
- Python 3.x
- pandas
- openpyxl

### Pasos

1. **Reemplaza el archivo Excel** con los datos actualizados:
   ```
   Email & Leads Campaign Summary & Plan for All Companies.xlsx
   ```

2. **Ejecuta el script de actualización**:
   ```bash
   python update_all_dashboards.py
   ```

3. **¡Listo!** Todos los 7 dashboards se regeneran automáticamente

---

## 📁 Estructura del Proyecto

```
unidashboard/
├── index.html                          # Página principal (selector de compañías)
├── dashboard_teamficient.html          # Dashboard de TeamFicient
├── dashboard_accusights.html           # Dashboard de AccuSights
├── dashboard_hirealatino.html          # Dashboard de HireALatino
├── dashboard_mydriveacademy.html       # Dashboard de MyDrive Academy
├── dashboard_idrivio.html              # Dashboard de iDrivio
├── dashboard_archficient.html          # Dashboard de ArchFicient
├── dashboard_docficient.html           # Dashboard de DocFicient
├── replies_teamficient.html            # Página de replies de TeamFicient
├── replies_accusights.html             # Página de replies de AccuSights
├── update_all_dashboards.py            # Script de actualización automática
└── README.txt                          # Instrucciones en español
```

---

## 🎨 Colores por Compañía

| Compañía | Colores |
|----------|---------|
| TeamFicient | `#3B82F6, #2563EB` |
| AccuSights | `#667eea, #764ba2` |
| HireALatino | `#0ea5e9, #0284c7` |
| MyDrive Academy | `#10b981, #059669` |
| iDrivio | `#f59e0b, #d97706` |
| ArchFicient | `#003B5C, #FDB913` |
| DocFicient | `#8B5CF6, #7C3AED` |

---

## 📊 Métricas Rastreadas

- **Leads Generated**: Total de contactos por campaña
- **Open Rate**: Porcentaje de emails abiertos
- **Click Rate**: Porcentaje de clicks en emails
- **Replies**: Respuestas recibidas categorizadas por tipo
- **Delivered**: Emails entregados exitosamente

---

## 🔒 Seguridad

- Los datos del archivo Excel **NO** se suben al repositorio (protegido por `.gitignore`)
- Los dashboards HTML contienen datos embebidos para funcionar de forma standalone
- No se requiere servidor backend

---

## 💻 Tecnologías

- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)
- **Charts**: Chart.js
- **Backend**: Python 3 (para generación de dashboards)
- **Data Processing**: pandas, openpyxl
- **Hosting**: GitHub Pages

---

## 📝 Notas

- Solo se muestran campañas con `Status = "Sent"`
- Los nombres de campañas provienen de la columna `Industry` (excepto ArchFicient que usa `Source`)
- Los dashboards funcionan sin servidor (archivos estáticos)
- Actualización manual mediante script Python

---

## 🤝 Contribuciones

Este es un proyecto privado para uso interno de las compañías listadas.

---

## 📧 Contacto

Para preguntas o soporte, contacta al administrador del sistema.

---

**Última actualización**: Abril 2026
