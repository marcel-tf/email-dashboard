===================================================================
   SISTEMA UNIFICADO DE DASHBOARDS - INSTRUCCIONES
===================================================================

CÓMO ACTUALIZAR TODOS LOS DASHBOARDS:

1. Reemplaza el archivo Excel:
   "Email & Leads Campaign Summary & Plan for All Companies.xlsx"
   con la nueva versión (mantén el mismo nombre)

2. Ejecuta el script de actualización:
   python update_all_dashboards.py

3. ¡LISTO! Todos los 7 dashboards se actualizan automáticamente

===================================================================

ARCHIVOS EN ESTE FOLDER:

📄 index.html - Página principal con selector de compañías

📊 DASHBOARDS (7 archivos):
   - dashboard_teamficient.html
   - dashboard_accusights.html
   - dashboard_hirealatino.html
   - dashboard_mydriveacademy.html
   - dashboard_idrivio.html
   - dashboard_archficient.html
   - dashboard_docficient.html

📊 EXCEL:
   - Email & Leads Campaign Summary & Plan for All Companies.xlsx

🔧 SCRIPT DE ACTUALIZACIÓN:
   - update_all_dashboards.py

===================================================================

COLORES DE CADA COMPAÑÍA:

TeamFicient    - Azul (#3B82F6)
AccuSights     - Morado/Violeta (#667eea, #764ba2)
HireALatino    - Azul Cielo (#0ea5e9, #0284c7)
MyDrive Academy- Verde (#10b981, #059669)
iDrivio        - Naranja (#f59e0b, #d97706)
ArchFicient    - Azul Oscuro/Dorado (#003B5C, #FDB913)
DocFicient     - Violeta (#8B5CF6, #7C3AED)

===================================================================

PARA USAR:

1. Abre index.html en tu navegador
2. Haz clic en cualquier compañía
3. Verás su dashboard con sus colores y datos

===================================================================

NOTAS:

- Los datos están embebidos en cada HTML (no necesita servidor)
- Cada dashboard tiene los colores originales de las versiones individuales
- Un solo comando actualiza todos los dashboards
- Los dashboards antiguos fuera de esta carpeta ya no son necesarios

===================================================================
