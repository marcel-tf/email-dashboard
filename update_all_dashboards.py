#!/usr/bin/env python3
"""
Script para actualizar TODOS los dashboards y replies desde el Excel unificado
Solo reemplaza el Excel y ejecuta este script - todo se actualiza!
"""

import pandas as pd
import json
from datetime import datetime
import numpy as np
import datetime as dt

# Configuración de compañías con sus colores originales
COMPANIES = {
    'TeamFicient': {
        'sheet': 'TeamFicient',
        'replies_sheet': 'TeamFicient Replies',
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #3B82F6 0%, #2563EB 100%)',
            'header_gradient': 'linear-gradient(135deg, #3B82F6 0%, #2563EB 100%)',
            'primary': '#3B82F6',
            'secondary': '#2563EB',
            'accent': '#1D4ED8'
        }
    },
    'AccuSights': {
        'sheet': 'Accusights',
        'replies_sheet': 'AccuSights Replies',
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            'header_gradient': 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            'primary': '#667eea',
            'secondary': '#764ba2',
            'accent': '#5a67d8'
        }
    },
    'HireALatino': {
        'sheet': 'HireALatino',
        'replies_sheet': 'HAL Replies',
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%)',
            'header_gradient': 'linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%)',
            'primary': '#0ea5e9',
            'secondary': '#0284c7',
            'accent': '#0369a1'
        }
    },
    'MyDriveAcademy': {
        'sheet': 'MyDriveAcademy',
        'replies_sheet': None,
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
            'header_gradient': 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
            'primary': '#10b981',
            'secondary': '#059669',
            'accent': '#047857'
        }
    },
    'iDrivio': {
        'sheet': 'Medicar SafetyiDrivio',
        'replies_sheet': 'Medicar Safety Replies',
        'sales_sheet': 'Medicar Safety SALES',
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
            'header_gradient': 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
            'primary': '#f59e0b',
            'secondary': '#d97706',
            'accent': '#b45309'
        }
        # Sin filter_company para incluir TODOS los datos de la hoja (iDrivio + Medicar Safety)
    },
    'ArchFicient': {
        'sheet': 'ArchFicient',
        'replies_sheet': 'ArchFicient Replies',
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #003B5C 0%, #001F3F 100%)',
            'header_gradient': 'linear-gradient(135deg, #2D2926 0%, #1a1918 100%)',
            'primary': '#003B5C',
            'secondary': '#FDB913',
            'accent': '#002844'
        }
    },
    'DocFicient': {
        'sheet': 'Docficient',
        'replies_sheet': None,
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #8B5CF6 0%, #7C3AED 100%)',
            'header_gradient': 'linear-gradient(135deg, #8B5CF6 0%, #7C3AED 100%)',
            'primary': '#8B5CF6',
            'secondary': '#7C3AED',
            'accent': '#6D28D9'
        }
    }
}

EXCEL_FILE = 'Email & Leads Campaign Summary & Plan for All Companies.xlsx'

def format_date(date_val):
    """Formatea una fecha de manera segura"""
    if pd.isna(date_val):
        return ''
    try:
        # Manejar datetime.datetime, pd.Timestamp, y otros objetos con strftime
        if isinstance(date_val, (pd.Timestamp, dt.datetime)):
            return date_val.strftime('%Y-%m-%d')
        elif hasattr(date_val, 'strftime') and not isinstance(date_val, str):
            return date_val.strftime('%Y-%m-%d')
        elif isinstance(date_val, str):
            # Try to parse string date
            dt_obj = pd.to_datetime(date_val, errors='coerce')
            if pd.notna(dt_obj):
                return dt_obj.strftime('%Y-%m-%d')
        return str(date_val)
    except:
        return ''

def read_company_data(sheet_name, filter_company=None):
    """Lee los datos de una compañía desde el Excel - SOLO campañas con Status = 'Sent'"""
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, skiprows=[0])
        df = df.dropna(how='all')

        # FILTRAR por compañía si se especifica (para sheets compartidas como Medicar SafetyiDrivio)
        if filter_company and 'Company' in df.columns:
            df = df[df['Company'].str.contains(filter_company, case=False, na=False)]
            print(f"  Filtradas {len(df)} filas para compañía '{filter_company}'")

        # FILTRAR SOLO Status/Lead Status = 'Sent'
        if 'Status' in df.columns:
            df = df[df['Status'].str.strip().str.lower() == 'sent']
            print(f"  Filtradas {len(df)} campañas con Status='Sent'")
        elif 'Lead Status' in df.columns:
            df = df[df['Lead Status'].str.strip().str.lower() == 'sent']
            print(f"  Filtradas {len(df)} campañas con Lead Status='Sent'")

        records = []
        for _, row in df.iterrows():
            record = {}
            for col in df.columns:
                val = row[col]
                if pd.notna(val):
                    # Manejar fechas (datetime, Timestamp, etc)
                    if isinstance(val, (pd.Timestamp, dt.datetime)) or (hasattr(val, 'strftime') and not isinstance(val, str)):
                        record[col] = format_date(val)
                    elif isinstance(val, (int, float, np.integer, np.floating)):
                        record[col] = float(val) if '.' in str(val) or 'Rate' in col else int(val)
                    else:
                        record[col] = str(val).strip()
            if record:
                records.append(record)

        return records
    except Exception as e:
        print(f"  ERROR leyendo {sheet_name}: {e}")
        return []

def read_replies_data(sheet_name):
    """Lee los datos de replies desde el Excel"""
    if not sheet_name:
        return []

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=0)
        df = df.dropna(how='all')

        replies = []
        for _, row in df.iterrows():
            reply = {}
            for col in df.columns:
                val = row[col]
                if pd.notna(val):
                    reply[col] = str(val).strip()
            if reply:
                replies.append(reply)

        return replies
    except Exception as e:
        print(f"  ERROR leyendo replies {sheet_name}: {e}")
        return []

def read_sales_data(sheet_name):
    """Lee los datos de SALES desde el Excel"""
    if not sheet_name:
        return []

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=0)
        df = df.dropna(how='all')

        sales = []
        for _, row in df.iterrows():
            sale = {}
            for col in df.columns:
                val = row[col]
                if pd.notna(val):
                    sale[col] = str(val).strip()
            if sale:
                sales.append(sale)

        return sales
    except Exception as e:
        print(f"  ERROR leyendo sales {sheet_name}: {e}")
        return []

def get_category_class(category):
    """Mapea categoría a clase CSS"""
    category_lower = category.lower().strip()

    # SOLD es especial - ventas generadas!
    if 'sold' in category_lower:
        return 'sold'
    elif 'not interested' in category_lower:
        return 'not-interested'
    elif 'invalid' in category_lower or 'wrong contact' in category_lower:
        return 'invalid'
    elif 'hot' in category_lower or 'meeting' in category_lower:
        return 'hot-lead'
    elif 'engaged' in category_lower and 'interested' in category_lower:
        return 'engaged-interested'
    elif 'follow' in category_lower:
        return 'follow-up'
    elif 'pipeline' in category_lower or 'circle back' in category_lower:
        return 'pipelined'
    elif 'no action' in category_lower:
        return 'no-action'
    elif 'existing client' in category_lower:
        return 'existing-client'
    else:
        return 'neutral'

def get_category_color(cat_class):
    """Retorna el color para una categoría - ACTUALIZADO con nuevos colores"""
    colors = {
        'sold': '#fbbf24',              # Gold - Maximum achievement
        'not-interested': '#9ca3af',     # Gray - Not interested
        'invalid': '#4b5563',            # Dark Gray - Invalid contact
        'hot-lead': '#3b82f6',           # Blue - Hot lead/Meeting scheduled
        'engaged-interested': '#10b981', # Green - Engaged/Interested
        'follow-up': '#d4a574',          # Beige - Follow-up needed
        'pipelined': '#06b6d4',          # Cyan - Pipelined/Circle back
        'no-action': '#8b5cf6',          # Purple - No action needed
        'existing-client': '#ec4899',    # Pink - Existing client
        'neutral': '#94a3b8'             # Slate - Default
    }
    return colors.get(cat_class, '#94a3b8')

def generate_replies_html(company_name, company_data, replies_data):
    """Genera el HTML de replies"""

    colors = company_data['colors']

    # Calcular estadísticas - USAR CATEGORÍAS REALES
    total_replies = len(replies_data)

    # Contar por categoría REAL
    real_category_counts = {}
    for reply in replies_data:
        category = reply.get('Category', 'Unknown')
        if category not in real_category_counts:
            real_category_counts[category] = 0
        real_category_counts[category] += 1

    # Ordenar categorías: Sold primero, luego el resto alfabéticamente
    sorted_categories = sorted(real_category_counts.keys(), key=lambda x: (x.lower() != 'sold', x.lower()))

    # Función para parsear fecha (DD-MM-YYYY o YYYY-MM-DD HH:MM:SS)
    def parse_date(date_str):
        """Convierte fecha string a formato comparable YYYY-MM-DD"""
        if not date_str:
            return '1900-01-01'
        # Limpiar espacios y manejar timestamps
        date_str = str(date_str).strip().split(' ')[0]
        parts = date_str.split('-')
        if len(parts) == 3:
            # Si el primer elemento tiene 2 dígitos, es DD-MM-YYYY
            if len(parts[0]) == 2:
                return f"{parts[2]}-{parts[1]}-{parts[0]}"  # Convertir a YYYY-MM-DD
            # Si tiene 4 dígitos, ya es YYYY-MM-DD
            else:
                return date_str
        return '1900-01-01'

    # Función para formatear fecha a DD/MM/YYYY para mostrar
    def format_date_display(date_str):
        """Convierte cualquier formato de fecha a DD/MM/YYYY para mostrar"""
        if not date_str:
            return 'Unknown'
        # Limpiar espacios y manejar timestamps
        date_str = str(date_str).strip().split(' ')[0]
        parts = date_str.split('-')
        if len(parts) == 3:
            # Si el primer elemento tiene 2 dígitos, es DD-MM-YYYY
            if len(parts[0]) == 2:
                return f"{parts[0]}/{parts[1]}/{parts[2]}"  # DD/MM/YYYY
            # Si tiene 4 dígitos, es YYYY-MM-DD, convertir a DD/MM/YYYY
            else:
                return f"{parts[2]}/{parts[1]}/{parts[0]}"  # DD/MM/YYYY
        return 'Unknown'

    # Ordenar replies por fecha (más reciente primero)
    sorted_replies = sorted(replies_data, key=lambda x: parse_date(x.get('Date', '')), reverse=True)

    # Generar HTML de tarjetas de respuestas
    replies_html = ''
    for reply in sorted_replies:
        who = reply.get('Who', 'Unknown')
        email = reply.get('Email', '')
        message = reply.get('Reply/Calls', '')
        action = reply.get('Action Taken', '')
        subject = reply.get('Subject if email ', reply.get('Subject', ''))
        campaign = reply.get('Email Campaign', '')
        category = reply.get('Category', 'Neutral')
        date_raw = reply.get('Date', '')
        date_display = format_date_display(date_raw)  # Formato consistente DD/MM/YYYY
        date_sortable = parse_date(date_raw)  # Formato para ordenar YYYY-MM-DD
        cat_class = get_category_class(category)

        replies_html += f'''
        <div class="response-card {cat_class}" data-category="{cat_class}" data-date="{date_sortable}">
            <div class="response-header">
                <div>
                    <div class="response-who">{who}</div>
                    <div class="response-email">{email}</div>
                </div>
                <div class="response-category {cat_class}">{category}</div>
            </div>
            <div class="response-date">
                <span class="date-icon">📅</span> {date_display}
            </div>
            <div class="response-details">
                <div class="response-detail-item">
                    <span class="response-detail-label">Subject:</span>
                    <span class="response-detail-value">{subject}</span>
                </div>
                <div class="response-detail-item">
                    <span class="response-detail-label">Campaign:</span>
                    <span class="response-detail-value">{campaign}</span>
                </div>
                <div class="response-detail-item">
                    <span class="response-detail-label">Action Taken:</span>
                    <span class="response-detail-value">{action}</span>
                </div>
            </div>
            <div class="response-message">"{message}"</div>
        </div>
        '''

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Response Analysis - {company_name}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {colors['body_gradient']};
            padding: 20px;
            min-height: 100vh;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }}

        header {{
            background: {colors['header_gradient']};
            color: white;
            padding: 30px 40px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        header h1 {{
            font-size: 32px;
            font-weight: 600;
        }}

        .back-btn {{
            padding: 12px 25px;
            background: rgba(255, 255, 255, 0.2);
            color: white;
            border: 2px solid white;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-block;
        }}

        .back-btn:hover {{
            background: white;
            color: {colors['primary']};
        }}

        .content {{
            padding: 40px;
        }}

        .summary-section {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
            margin-bottom: 40px;
        }}

        .chart-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        }}

        .chart-card h2 {{
            color: #2c3e50;
            font-size: 22px;
            margin-bottom: 20px;
            font-weight: 600;
        }}

        .chart-container {{
            position: relative;
            height: 350px;
        }}

        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 15px;
        }}

        .stat-item {{
            padding: 20px;
            background: #f8f9fa;
            border-radius: 8px;
            border-left: 5px solid;
        }}

        .stat-item.not-interested {{
            border-left-color: #9ca3af;
        }}

        .stat-item.engaged-interested {{
            border-left-color: #10b981;
        }}

        .stat-item.hot-lead {{
            border-left-color: #3b82f6;
        }}

        .stat-item.no-action {{
            border-left-color: #8b5cf6;
        }}

        .stat-item.follow-up {{
            border-left-color: #d4a574;
        }}

        .stat-item.pipelined {{
            border-left-color: #06b6d4;
        }}

        .stat-item.invalid {{
            border-left-color: #4b5563;
        }}

        .stat-item.sold {{
            border-left-color: #fbbf24;
            background: linear-gradient(to right, rgba(251, 191, 36, 0.05), white);
        }}

        .stat-item.existing-client {{
            border-left-color: #ec4899;
        }}

        .stat-item.neutral {{
            border-left-color: #94a3b8;
        }}

        .stat-label {{
            font-size: 13px;
            color: #7f8c8d;
            font-weight: 500;
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .stat-value {{
            font-size: 32px;
            font-weight: 700;
            color: #2c3e50;
        }}

        .stat-percentage {{
            font-size: 16px;
            color: #7f8c8d;
            margin-left: 8px;
        }}

        .responses-section {{
            margin-top: 40px;
        }}

        .responses-section h2 {{
            color: #2c3e50;
            font-size: 24px;
            margin-bottom: 20px;
            font-weight: 600;
        }}

        .response-card {{
            background: white;
            border-radius: 10px;
            padding: 25px;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            border-left: 5px solid;
            transition: transform 0.2s ease;
        }}

        .response-card:hover {{
            transform: translateX(5px);
        }}

        .response-card.negative {{
            border-left-color: #EF4444;
        }}

        .response-card.interested {{
            border-left-color: #3B82F6;
        }}

        .response-card.hot-lead {{
            border-left-color: #10B981;
        }}

        .response-card.neutral {{
            border-left-color: #8B5CF6;
        }}

        .response-header {{
            display: flex;
            justify-content: space-between;
            align-items: start;
            margin-bottom: 15px;
        }}

        .response-who {{
            font-size: 20px;
            font-weight: 700;
            color: #2c3e50;
        }}

        .response-email {{
            font-size: 14px;
            color: #7f8c8d;
            margin-top: 5px;
        }}

        .response-category {{
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        /* Updated category colors */
        .response-category.not-interested {{
            background: #9ca3af;
            color: white;
        }}

        .response-category.engaged-interested {{
            background: #10b981;
            color: white;
        }}

        .response-category.hot-lead {{
            background: #3b82f6;
            color: white;
        }}

        .response-category.no-action {{
            background: #8b5cf6;
            color: white;
        }}

        .response-category.follow-up {{
            background: #d4a574;
            color: white;
        }}

        .response-category.pipelined {{
            background: #06b6d4;
            color: white;
        }}

        .response-category.invalid {{
            background: #4b5563;
            color: white;
        }}

        .response-category.sold {{
            background: linear-gradient(135deg, #fbbf24, #f59e0b);
            color: #000;
            font-weight: 700;
            font-size: 13px;
            box-shadow: 0 4px 15px rgba(251, 191, 36, 0.5);
            animation: pulse-sold 2s ease-in-out infinite;
        }}

        @keyframes pulse-sold {{
            0%, 100% {{ box-shadow: 0 4px 15px rgba(251, 191, 36, 0.5); }}
            50% {{ box-shadow: 0 4px 25px rgba(251, 191, 36, 0.8); }}
        }}

        .response-category.existing-client {{
            background: #ec4899;
            color: white;
        }}

        .response-category.neutral {{
            background: #94a3b8;
            color: white;
        }}

        .response-date {{
            background: #f0f9ff;
            border: 1px solid #bfdbfe;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            color: #1e40af;
            margin-bottom: 15px;
            display: inline-block;
        }}

        .date-icon {{
            margin-right: 5px;
        }}

        .response-details {{
            margin-bottom: 15px;
        }}

        .response-detail-item {{
            display: flex;
            margin-bottom: 8px;
            font-size: 14px;
        }}

        .response-detail-label {{
            font-weight: 600;
            color: #2c3e50;
            min-width: 120px;
        }}

        .response-detail-value {{
            color: #5a6c7d;
        }}

        .response-message {{
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            font-size: 14px;
            line-height: 1.6;
            color: #2c3e50;
            font-style: italic;
        }}

        .filters-container {{
            margin-bottom: 25px;
        }}

        .date-filters {{
            display: flex;
            gap: 15px;
            margin-bottom: 15px;
            align-items: center;
            flex-wrap: wrap;
        }}

        .date-filter-group {{
            display: flex;
            flex-direction: column;
            gap: 5px;
        }}

        .date-filter-group label {{
            font-size: 12px;
            font-weight: 600;
            color: #2c3e50;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .date-input {{
            padding: 8px 12px;
            border: 2px solid #e9ecef;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s ease;
        }}

        .date-input:focus {{
            outline: none;
            border-color: #3b82f6;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
        }}

        .clear-dates-btn {{
            padding: 8px 16px;
            background: #ef4444;
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            align-self: flex-end;
        }}

        .clear-dates-btn:hover {{
            background: #dc2626;
            transform: translateY(-1px);
        }}

        .filter-buttons {{
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}

        .filter-btn {{
            padding: 10px 20px;
            border: 2px solid #e9ecef;
            background: white;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
        }}

        .filter-btn:hover {{
            transform: translateY(-2px);
        }}

        .filter-btn.active {{
            color: white;
        }}

        .filter-btn.all.active {{
            background: #2c3e50;
            border-color: #2c3e50;
        }}

        .filter-btn.not-interested.active {{
            background: #9ca3af;
            border-color: #9ca3af;
        }}

        .filter-btn.engaged-interested.active {{
            background: #10b981;
            border-color: #10b981;
        }}

        .filter-btn.hot-lead.active {{
            background: #3b82f6;
            border-color: #3b82f6;
        }}

        .filter-btn.no-action.active {{
            background: #8b5cf6;
            border-color: #8b5cf6;
        }}

        .filter-btn.follow-up.active {{
            background: #d4a574;
            border-color: #d4a574;
        }}

        .filter-btn.pipelined.active {{
            background: #06b6d4;
            border-color: #06b6d4;
        }}

        .filter-btn.invalid.active {{
            background: #4b5563;
            border-color: #4b5563;
        }}

        .filter-btn.sold {{
            border: 3px solid #fbbf24;
            font-weight: 700;
        }}

        .filter-btn.sold.active {{
            background: linear-gradient(135deg, #fbbf24, #f59e0b);
            border-color: #fbbf24;
            color: #000;
            box-shadow: 0 4px 15px rgba(251, 191, 36, 0.5);
        }}

        .filter-btn.existing-client.active {{
            background: #ec4899;
            border-color: #ec4899;
        }}

        .filter-btn.neutral.active {{
            background: #94a3b8;
            border-color: #94a3b8;
        }}

        @media (max-width: 768px) {{
            .summary-section {{
                grid-template-columns: 1fr;
            }}

            .stats-grid {{
                grid-template-columns: 1fr;
            }}

            header {{
                flex-direction: column;
                gap: 15px;
                text-align: center;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>Response Analysis - {company_name}</h1>
            <a href="dashboard_{company_name.lower().replace(' ', '')}.html" class="back-btn">← Back to Dashboard</a>
        </header>

        <div class="content">
            <div class="summary-section">
                <div class="chart-card">
                    <h2>Response Distribution</h2>
                    <div class="chart-container">
                        <canvas id="responseChart"></canvas>
                    </div>
                </div>

                <div class="chart-card">
                    <h2>Response Statistics by Category</h2>
                    <div class="stats-grid">
                        {''.join([f'''<div class="stat-item {get_category_class(cat)}">
                            <div class="stat-label">{cat}</div>
                            <div class="stat-value">{real_category_counts[cat]} <span class="stat-percentage">({real_category_counts[cat]/total_replies*100 if total_replies > 0 else 0:.0f}%)</span></div>
                        </div>''' for cat in sorted_categories])}
                    </div>
                </div>
            </div>

            <div class="responses-section">
                <h2>All Responses ({total_replies})</h2>

                <div class="filters-container">
                    <div class="date-filters">
                        <div class="date-filter-group">
                            <label for="startDate">From Date</label>
                            <input type="date" id="startDate" class="date-input" onchange="applyFilters()">
                        </div>
                        <div class="date-filter-group">
                            <label for="endDate">To Date</label>
                            <input type="date" id="endDate" class="date-input" onchange="applyFilters()">
                        </div>
                        <button class="clear-dates-btn" onclick="clearDateFilters()">Clear Dates</button>
                    </div>

                    <div class="filter-buttons">
                        <button class="filter-btn all active" onclick="filterResponses('all')">All ({total_replies})</button>
                        {''.join([f'''<button class="filter-btn {get_category_class(cat)}" onclick="filterResponses('{get_category_class(cat)}')">{cat} ({real_category_counts[cat]})</button>''' for cat in sorted_categories])}
                    </div>
                </div>

                <div id="responsesContainer">
                    {replies_html}
                </div>
            </div>
        </div>
    </div>

    <script>
        // Create response distribution chart with REAL categories (same style as dashboard)
        const ctx = document.getElementById('responseChart');
        new Chart(ctx, {{
            type: 'doughnut',
            data: {{
                labels: [{','.join([f"'{cat}'" for cat in sorted_categories])}],
                datasets: [{{
                    data: [{','.join([str(real_category_counts[cat]) for cat in sorted_categories])}],
                    backgroundColor: [
                        {','.join([f"'{get_category_color(get_category_class(cat))}'" for cat in sorted_categories])}
                    ],
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        position: 'right',
                        labels: {{
                            padding: 12,
                            font: {{
                                size: 11
                            }},
                            usePointStyle: true,
                            pointStyle: 'circle'
                        }}
                    }},
                    tooltip: {{
                        callbacks: {{
                            label: function(context) {{
                                const label = context.label || '';
                                const value = context.parsed || 0;
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                                return '  ' + label + ': ' + value + ' (' + percentage + '%)';
                            }}
                        }}
                    }}
                }}
            }}
        }});

        // State for filters
        let currentCategory = 'all';

        // Filter responses by category
        function filterResponses(category) {{
            currentCategory = category;
            const buttons = document.querySelectorAll('.filter-btn');

            // Update button states
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');

            applyFilters();
        }}

        // Apply all filters (category + date)
        function applyFilters() {{
            const cards = document.querySelectorAll('.response-card');
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;

            cards.forEach(card => {{
                let show = true;

                // Category filter
                if (currentCategory !== 'all') {{
                    if (card.dataset.category !== currentCategory) {{
                        show = false;
                    }}
                }}

                // Date filter
                if (show && (startDate || endDate)) {{
                    const cardDate = card.dataset.date;
                    if (startDate && cardDate < startDate) {{
                        show = false;
                    }}
                    if (endDate && cardDate > endDate) {{
                        show = false;
                    }}
                }}

                card.style.display = show ? 'block' : 'none';
            }});

            // Update count in "All" button
            const visibleCards = document.querySelectorAll('.response-card[style="display: block;"]').length;
            const allButton = document.querySelector('.filter-btn.all');
            if (allButton) {{
                const totalMatch = allButton.textContent.match(/\d+/);
                if (totalMatch) {{
                    allButton.textContent = `All (${{visibleCards}})`;
                }}
            }}
        }}

        // Clear date filters
        function clearDateFilters() {{
            document.getElementById('startDate').value = '';
            document.getElementById('endDate').value = '';
            applyFilters();
        }}
    </script>
</body>
</html>
'''
    return html

def generate_dashboard_html(company_name, company_data, data_records, replies_count):
    """Genera el HTML del dashboard con datos embebidos"""

    colors = company_data['colors']

    # Convertir datos a formato JavaScript
    js_data = json.dumps(data_records, indent=12)

    # Botón de replies AL FINAL (solo si hay replies)
    replies_button = ''
    if replies_count > 0:
        replies_button = f'''
                <div class="kpi-card info clickable" onclick="window.location.href='replies_{company_name.lower().replace(' ', '')}.html'">
                    <div class="kpi-label">Total Replied</div>
                    <div class="kpi-value">{replies_count}</div>
                    <div class="kpi-subtitle">Customer responses (Click for details)</div>
                </div>
        '''
    else:
        replies_button = '''
                <div class="kpi-card info">
                    <div class="kpi-label">Total Replied</div>
                    <div class="kpi-value">0</div>
                    <div class="kpi-subtitle">No responses yet</div>
                </div>
        '''

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{company_name} - Email Campaign Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {colors['body_gradient']};
            padding: 20px;
            min-height: 100vh;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }}

        header {{
            background: {colors['header_gradient']};
            color: white;
            padding: 30px 40px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        header h1 {{
            font-size: 32px;
            font-weight: 600;
        }}

        .back-btn {{
            padding: 12px 25px;
            background: rgba(255, 255, 255, 0.2);
            color: white;
            border: 2px solid white;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-block;
        }}

        .back-btn:hover {{
            background: white;
            color: {colors['primary']};
        }}

        .filters-section {{
            background: #f8f9fa;
            padding: 25px 40px;
            border-bottom: 2px solid #e9ecef;
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
            align-items: center;
        }}

        .filter-group {{
            display: flex;
            flex-direction: column;
            gap: 8px;
            flex: 1;
            min-width: 200px;
        }}

        .filter-group label {{
            font-size: 13px;
            font-weight: 600;
            color: #2c3e50;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .filter-group select {{
            padding: 10px 15px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 14px;
            background: white;
            cursor: pointer;
            transition: border-color 0.3s ease;
        }}

        .filter-group select:focus {{
            outline: none;
            border-color: {colors['primary']};
        }}

        .filter-group .date-input {{
            padding: 10px 15px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 14px;
            background: white;
            cursor: pointer;
            transition: border-color 0.3s ease;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}

        .filter-group .date-input:focus {{
            outline: none;
            border-color: {colors['primary']};
        }}

        .clear-filters-btn {{
            padding: 10px 25px;
            background: {colors['body_gradient']};
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s ease;
            margin-top: 20px;
        }}

        .clear-filters-btn:hover {{
            transform: translateY(-2px);
        }}

        .dashboard-content {{
            padding: 40px;
        }}

        .kpi-section {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}

        .kpi-card {{
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
            transition: transform 0.3s ease;
        }}

        .kpi-card:hover {{
            transform: translateY(-5px);
        }}

        .kpi-card.clickable {{
            cursor: pointer;
        }}

        .kpi-card.clickable:hover {{
            transform: translateY(-8px);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.15);
        }}

        .kpi-card.primary {{
            background: {colors['body_gradient']};
            color: white;
        }}

        .kpi-card.success {{
            background: linear-gradient(135deg, #10B981 0%, #059669 100%);
            color: white;
        }}

        .kpi-card.warning {{
            background: linear-gradient(135deg, #F59E0B 0%, #D97706 100%);
            color: white;
        }}

        .kpi-card.info {{
            background: linear-gradient(135deg, #3B82F6 0%, #2563EB 100%);
            color: white;
        }}

        .kpi-label {{
            font-size: 13px;
            font-weight: 500;
            opacity: 0.9;
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .kpi-value {{
            font-size: 36px;
            font-weight: 700;
            margin-bottom: 4px;
        }}

        .kpi-subtitle {{
            font-size: 12px;
            opacity: 0.8;
        }}

        .charts-section {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(450px, 1fr));
            gap: 30px;
            margin-bottom: 30px;
        }}

        .chart-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        }}

        .chart-card h3 {{
            color: #2c3e50;
            font-size: 18px;
            margin-bottom: 20px;
            font-weight: 600;
        }}

        .chart-container {{
            position: relative;
            height: 300px;
        }}

        .full-width {{
            grid-column: 1 / -1;
        }}

        @media (max-width: 768px) {{
            .charts-section {{
                grid-template-columns: 1fr;
            }}

            .dashboard-content {{
                padding: 20px;
            }}

            header {{
                flex-direction: column;
                gap: 15px;
                text-align: center;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>{company_name} - Campaign Dashboard</h1>
            <a href="index.html" class="back-btn">← Back to Companies</a>
        </header>

        <div class="filters-section">
            <div class="filter-group">
                <label for="campaignFilter">Campaign</label>
                <select id="campaignFilter">
                    <option value="">All Campaigns</option>
                </select>
            </div>
            <div class="filter-group">
                <label for="startDate">Start Date</label>
                <input type="date" id="startDate" class="date-input">
            </div>
            <div class="filter-group">
                <label for="endDate">End Date</label>
                <input type="date" id="endDate" class="date-input">
            </div>
            <button class="clear-filters-btn" onclick="clearFilters()">Clear Filters</button>
        </div>

        <div class="dashboard-content">
            <div class="kpi-section">
                <div class="kpi-card primary">
                    <div class="kpi-label">Total Leads</div>
                    <div class="kpi-value" id="totalLeads">0</div>
                    <div class="kpi-subtitle">Generated across campaigns</div>
                </div>

                <div class="kpi-card success">
                    <div class="kpi-label">Avg Open Rate</div>
                    <div class="kpi-value" id="avgOpenRate">0%</div>
                    <div class="kpi-subtitle">Campaign engagement</div>
                </div>

                <div class="kpi-card warning">
                    <div class="kpi-label">Total Clicked</div>
                    <div class="kpi-value" id="totalClicked">0</div>
                    <div class="kpi-subtitle">Link interactions</div>
                </div>

                {replies_button}
            </div>

            <div class="charts-section">
                <div class="chart-card full-width">
                    <h3>Performance Over Time</h3>
                    <div class="chart-container">
                        <canvas id="performanceChart"></canvas>
                    </div>
                </div>

                <div class="chart-card">
                    <h3>Leads by Campaign</h3>
                    <div class="chart-container">
                        <canvas id="campaignChart"></canvas>
                    </div>
                </div>

                <div class="chart-card">
                    <h3>Engagement Metrics</h3>
                    <div class="chart-container">
                        <canvas id="engagementChart"></canvas>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Data embedded from Excel - Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
        // IMPORTANTE: Solo campañas con Status='Sent'
        const allData = {js_data};

        let charts = {{
            performance: null,
            campaign: null,
            engagement: null
        }};

        // Process data into standard format - USAR INDUSTRY COMO NOMBRE DE CAMPAÑA
        function processData(data) {{
            return data.map(item => {{
                // Validar y formatear fecha
                let dateStr = item['Date Sent'] || item['date'] || '';
                if (dateStr && dateStr !== 'NaT') {{
                    try {{
                        // Asegurarse de que la fecha esté en formato correcto
                        const dateParts = dateStr.split('-');
                        if (dateParts.length === 3) {{
                            dateStr = `${{dateParts[0]}}-${{dateParts[1].padStart(2, '0')}}-${{dateParts[2].padStart(2, '0')}}`;
                        }}
                    }} catch (e) {{
                        dateStr = '';
                    }}
                }}

                return {{
                    date: dateStr,
                    campaign: item['Source'] || item['source'] || item['Industry'] || item['industry'] || 'Unknown',  // USAR SOURCE (ArchFicient) o INDUSTRY
                    industry: item['Industry'] || item['industry'] || '',
                    leads: parseInt(item['Leads Generated'] || item['leads'] || 0),
                    opened: parseInt(item['Opened'] || item['opened'] || 0),
                    clicked: parseInt(item['Clicked'] || item['clicked'] || 0),
                    delivered: parseInt(item['Delivered'] || item['delivered'] || item['Leads Generated'] || 0),
                    openRate: parseFloat(item['Open Rate %'] || item['openRate'] || 0)
                }};
            }}).filter(item => item.date && item.date !== '' && item.date !== 'NaT');  // Filtrar fechas inválidas
        }}

        function formatDisplayDate(dateStr) {{
            if (!dateStr || dateStr === 'NaT') return 'Unknown';
            try {{
                const date = new Date(dateStr + 'T00:00:00');
                if (isNaN(date.getTime())) return 'Unknown';
                return date.toLocaleDateString('en-US', {{ month: 'short', day: 'numeric', year: 'numeric' }});
            }} catch (e) {{
                return 'Unknown';
            }}
        }}

        function formatChartDate(dateStr) {{
            if (!dateStr || dateStr === 'NaT') return 'Unknown';
            try {{
                const date = new Date(dateStr + 'T00:00:00');
                if (isNaN(date.getTime())) return 'Unknown';
                return date.toLocaleDateString('en-US', {{ month: 'short', day: 'numeric' }});
            }} catch (e) {{
                return 'Unknown';
            }}
        }}

        function initializeFilters() {{
            const processedData = processData(allData);
            const campaigns = [...new Set(processedData.map(d => d.campaign).filter(c => c))];

            const campaignFilter = document.getElementById('campaignFilter');
            const startDateInput = document.getElementById('startDate');
            const endDateInput = document.getElementById('endDate');

            campaigns.forEach(campaign => {{
                const option = document.createElement('option');
                option.value = campaign;
                option.textContent = campaign;
                campaignFilter.appendChild(option);
            }});

            // Event listeners for filters
            campaignFilter.addEventListener('change', updateDashboard);
            startDateInput.addEventListener('change', updateDashboard);
            endDateInput.addEventListener('change', updateDashboard);
        }}

        function clearFilters() {{
            document.getElementById('campaignFilter').value = '';
            document.getElementById('startDate').value = '';
            document.getElementById('endDate').value = '';
            updateDashboard();
        }}

        function getFilteredData() {{
            const campaignFilter = document.getElementById('campaignFilter').value;
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const processedData = processData(allData);

            return processedData.filter(item => {{
                const matchesCampaign = !campaignFilter || item.campaign === campaignFilter;

                // Date range filtering
                let matchesDate = true;
                if (startDate || endDate) {{
                    const itemDate = item.date;
                    if (startDate && itemDate < startDate) matchesDate = false;
                    if (endDate && itemDate > endDate) matchesDate = false;
                }}

                return matchesCampaign && matchesDate;
            }});
        }}

        function updateDashboard() {{
            const data = getFilteredData();

            // Calculate KPIs
            const totalLeads = data.reduce((sum, row) => sum + row.leads, 0);
            const totalOpened = data.reduce((sum, row) => sum + row.opened, 0);
            const totalClicked = data.reduce((sum, row) => sum + row.clicked, 0);
            const avgOpenRate = data.length > 0 ? (data.reduce((sum, row) => sum + row.openRate, 0) / data.length * 100).toFixed(1) : 0;

            document.getElementById('totalLeads').textContent = totalLeads.toLocaleString();
            document.getElementById('avgOpenRate').textContent = avgOpenRate + '%';
            document.getElementById('totalClicked').textContent = totalClicked.toLocaleString();

            // Update charts
            updatePerformanceChart(data);
            updateCampaignChart(data);
            updateEngagementChart(data);
        }}

        function updatePerformanceChart(data) {{
            const dates = [...new Set(data.map(d => d.date))].sort();
            const dateData = dates.map(date => {{
                const dayData = data.filter(d => d.date === date);
                return {{
                    date,
                    leads: dayData.reduce((sum, d) => sum + d.leads, 0),
                    opened: dayData.reduce((sum, d) => sum + d.opened, 0),
                    clicked: dayData.reduce((sum, d) => sum + d.clicked, 0)
                }};
            }});

            const ctx = document.getElementById('performanceChart');

            if (charts.performance) {{
                charts.performance.destroy();
            }}

            charts.performance = new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: dateData.map(d => formatChartDate(d.date)),
                    datasets: [
                        {{
                            label: 'Leads Generated',
                            data: dateData.map(d => d.leads),
                            borderColor: '#3B82F6',
                            backgroundColor: 'rgba(59, 130, 246, 0.1)',
                            tension: 0.4,
                            fill: true
                        }},
                        {{
                            label: 'Emails Opened',
                            data: dateData.map(d => d.opened),
                            borderColor: '#10B981',
                            backgroundColor: 'rgba(16, 185, 129, 0.1)',
                            tension: 0.4,
                            fill: true
                        }},
                        {{
                            label: 'Links Clicked',
                            data: dateData.map(d => d.clicked),
                            borderColor: '#F43F5E',
                            backgroundColor: 'rgba(244, 63, 94, 0.1)',
                            tension: 0.4,
                            fill: true
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'top',
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true
                        }}
                    }}
                }}
            }});
        }}

        function updateCampaignChart(data) {{
            const campaigns = {{}};
            data.forEach(row => {{
                campaigns[row.campaign] = (campaigns[row.campaign] || 0) + row.leads;
            }});

            const ctx = document.getElementById('campaignChart');

            if (charts.campaign) {{
                charts.campaign.destroy();
            }}

            charts.campaign = new Chart(ctx, {{
                type: 'doughnut',
                data: {{
                    labels: Object.keys(campaigns),
                    datasets: [{{
                        data: Object.values(campaigns),
                        backgroundColor: [
                            '#3B82F6', '#10B981', '#F43F5E', '#F59E0B', '#8B5CF6',
                            '#06b6d4', '#ec4899', '#fbbf24', '#a855f7', '#14b8a6',
                            '#f97316', '#84cc16', '#06b6d4', '#6366f1', '#d946ef'
                        ]
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'bottom',
                        }}
                    }}
                }}
            }});
        }}

        function updateEngagementChart(data) {{
            const totalOpened = data.reduce((sum, row) => sum + row.opened, 0);
            const totalClicked = data.reduce((sum, row) => sum + row.clicked, 0);
            const totalDelivered = data.reduce((sum, row) => sum + row.delivered, 0);

            const ctx = document.getElementById('engagementChart');

            if (charts.engagement) {{
                charts.engagement.destroy();
            }}

            charts.engagement = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: ['Opened', 'Clicked'],
                    datasets: [{{
                        label: 'Count',
                        data: [totalOpened, totalClicked],
                        backgroundColor: ['#10B981', '#F43F5E']
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            display: false
                        }},
                        tooltip: {{
                            callbacks: {{
                                afterLabel: function(context) {{
                                    if (totalDelivered > 0) {{
                                        const percentage = ((context.parsed.y / totalDelivered) * 100).toFixed(2);
                                        return percentage + '% of delivered';
                                    }}
                                    return '';
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            ticks: {{
                                precision: 0
                            }}
                        }}
                    }}
                }}
            }});
        }}

        // Initialize on page load
        initializeFilters();
        updateDashboard();
    </script>
</body>
</html>
'''
    return html

def generate_improved_dashboard_teamficient(company_name, company_data, data_records, replies_data):
    """Genera el dashboard MEJORADO para TeamFicient con todas las funcionalidades avanzadas"""

    colors = company_data['colors']

    # Convertir datos a formato JavaScript
    js_campaign_data = json.dumps(data_records, indent=2, ensure_ascii=False)
    js_replies_data = json.dumps(replies_data, indent=2, ensure_ascii=False)

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{company_name} - Email Campaign Dashboard (Enhanced)</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {colors['body_gradient']};
            padding: 20px;
            min-height: 100vh;
        }}

        .container {{
            max-width: 1600px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }}

        header {{
            background: {colors['header_gradient']};
            color: white;
            padding: 30px 40px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        header h1 {{
            font-size: 32px;
            font-weight: 600;
        }}

        .back-btn {{
            padding: 12px 25px;
            background: rgba(255, 255, 255, 0.2);
            color: white;
            border: 2px solid white;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-block;
        }}

        .back-btn:hover {{
            background: white;
            color: {colors['primary']};
        }}

        .filters-section {{
            background: #f8f9fa;
            padding: 25px 40px;
            border-bottom: 2px solid #e9ecef;
        }}

        .filter-row {{
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
            align-items: flex-end;
            margin-bottom: 15px;
        }}

        .filter-group {{
            display: flex;
            flex-direction: column;
            gap: 8px;
            flex: 1;
            min-width: 200px;
        }}

        .filter-group label {{
            font-size: 13px;
            font-weight: 600;
            color: #2c3e50;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .filter-group .date-input {{
            padding: 10px 15px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 14px;
            background: white;
            cursor: pointer;
            transition: border-color 0.3s ease;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}

        .filter-group .date-input:focus {{
            outline: none;
            border-color: {colors['primary']};
        }}

        .multi-select-container {{
            position: relative;
            flex: 2;
            min-width: 300px;
        }}

        .multi-select-button {{
            padding: 10px 15px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 14px;
            background: white;
            cursor: pointer;
            transition: border-color 0.3s ease;
            display: flex;
            justify-content: space-between;
            align-items: center;
            width: 100%;
        }}

        .multi-select-button:hover {{
            border-color: {colors['primary']};
        }}

        .multi-select-dropdown {{
            position: absolute;
            top: 100%;
            left: 0;
            right: 0;
            margin-top: 5px;
            background: white;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            max-height: 300px;
            overflow-y: auto;
            display: none;
            z-index: 1000;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        }}

        .multi-select-dropdown.active {{
            display: block;
        }}

        .multi-select-option {{
            padding: 10px 15px;
            cursor: pointer;
            transition: background 0.2s;
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        .multi-select-option:hover {{
            background: #f8f9fa;
        }}

        .multi-select-option input[type="checkbox"] {{
            width: 18px;
            height: 18px;
            cursor: pointer;
        }}

        .clear-filters-btn {{
            padding: 10px 25px;
            background: {colors['body_gradient']};
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s ease;
        }}

        .clear-filters-btn:hover {{
            transform: translateY(-2px);
        }}

        .dashboard-content {{
            padding: 40px;
        }}

        .kpi-section {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}

        .kpi-card {{
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
            transition: transform 0.3s ease;
        }}

        .kpi-card:hover {{
            transform: translateY(-5px);
        }}

        .kpi-card.clickable {{
            cursor: pointer;
        }}

        .kpi-card.clickable:hover {{
            transform: translateY(-8px);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.15);
        }}

        .kpi-card.primary {{
            background: {colors['body_gradient']};
            color: white;
        }}

        .kpi-card.success {{
            background: linear-gradient(135deg, #10B981 0%, #059669 100%);
            color: white;
        }}

        .kpi-card.warning {{
            background: linear-gradient(135deg, #F59E0B 0%, #D97706 100%);
            color: white;
        }}

        .kpi-card.info {{
            background: linear-gradient(135deg, #8B5CF6 0%, #7C3AED 100%);
            color: white;
        }}

        .kpi-card.purple {{
            background: linear-gradient(135deg, #A855F7 0%, #9333EA 100%);
            color: white;
        }}

        .kpi-card.cyan {{
            background: linear-gradient(135deg, #06B6D4 0%, #0891B2 100%);
            color: white;
        }}

        .kpi-label {{
            font-size: 11px;
            font-weight: 500;
            opacity: 0.9;
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .kpi-value {{
            font-size: 28px;
            font-weight: 700;
            margin-bottom: 4px;
        }}

        .kpi-subtitle {{
            font-size: 11px;
            opacity: 0.8;
        }}

        .charts-section {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 30px;
            margin-bottom: 30px;
        }}

        .chart-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        }}

        .chart-card h3 {{
            color: #2c3e50;
            font-size: 18px;
            margin-bottom: 20px;
            font-weight: 600;
        }}

        .chart-container {{
            position: relative;
            height: 350px;
        }}

        .chart-container.small {{
            height: 300px;
        }}

        .chart-container.large {{
            height: 500px;
        }}

        .full-width {{
            grid-column: 1 / -1;
        }}

        @media (max-width: 1024px) {{
            .charts-section {{
                grid-template-columns: 1fr;
            }}

            .dashboard-content {{
                padding: 20px;
            }}

            header {{
                flex-direction: column;
                gap: 15px;
                text-align: center;
            }}

            .kpi-section {{
                grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>{company_name} - Campaign Dashboard</h1>
            <a href="index.html" class="back-btn">← Back to Companies</a>
        </header>

        <div class="filters-section">
            <div class="filter-row">
                <div class="multi-select-container">
                    <label>Campaigns (Multi-Select)</label>
                    <div class="multi-select-button" onclick="toggleMultiSelect()">
                        <span id="selectedCampaignsText">All Campaigns Selected</span>
                        <span>▼</span>
                    </div>
                    <div class="multi-select-dropdown" id="campaignDropdown">
                        <div class="multi-select-option">
                            <input type="checkbox" id="selectAll" checked onchange="toggleSelectAll()">
                            <label for="selectAll" style="cursor: pointer; font-weight: 600;">Select All</label>
                        </div>
                        <div id="campaignCheckboxes"></div>
                    </div>
                </div>
                <div class="filter-group">
                    <label for="startDate">Start Date</label>
                    <input type="date" id="startDate" class="date-input" onchange="updateDashboard()">
                </div>
                <div class="filter-group">
                    <label for="endDate">End Date</label>
                    <input type="date" id="endDate" class="date-input" onchange="updateDashboard()">
                </div>
                <button class="clear-filters-btn" onclick="clearFilters()">Clear Filters</button>
            </div>
        </div>

        <div class="dashboard-content">
            <div class="kpi-section">
                <div class="kpi-card primary">
                    <div class="kpi-label">Total Leads</div>
                    <div class="kpi-value" id="totalLeads">0</div>
                    <div class="kpi-subtitle">Generated across campaigns</div>
                </div>

                <div class="kpi-card success">
                    <div class="kpi-label">Avg Open Rate</div>
                    <div class="kpi-value" id="avgOpenRate">0%</div>
                    <div class="kpi-subtitle">Email engagement</div>
                </div>

                <div class="kpi-card warning">
                    <div class="kpi-label">Total Clicked</div>
                    <div class="kpi-value" id="totalClicked">0</div>
                    <div class="kpi-subtitle">Link interactions</div>
                </div>

                <div class="kpi-card info clickable" onclick="window.location.href='replies_teamficient.html'">
                    <div class="kpi-label">Total Replied</div>
                    <div class="kpi-value" id="totalReplied">0</div>
                    <div class="kpi-subtitle">Customer responses</div>
                </div>

                <div class="kpi-card purple">
                    <div class="kpi-label">Reply Rate</div>
                    <div class="kpi-value" id="replyRate">0%</div>
                    <div class="kpi-subtitle">Of delivered emails</div>
                </div>

                <div class="kpi-card cyan">
                    <div class="kpi-label">Click Rate</div>
                    <div class="kpi-value" id="clickRate">0%</div>
                    <div class="kpi-subtitle">CTR (click-through)</div>
                </div>
            </div>

            <div class="charts-section">
                <div class="chart-card full-width">
                    <h3>📊 Performance Over Time - Complete Campaign Analysis</h3>
                    <p style="margin: -10px 0 15px 0; color: #64748b; font-size: 13px;">
                        Shows Leads, Delivered, Opened, Clicked, Replied count + Open Rate % for each selected campaign
                    </p>
                    <div class="chart-container large">
                        <canvas id="performanceChart"></canvas>
                    </div>
                </div>

                <div class="chart-card">
                    <h3>Conversion Funnel</h3>
                    <div class="chart-container small">
                        <canvas id="funnelChart"></canvas>
                    </div>
                </div>

                <div class="chart-card">
                    <h3>Reply Categories Distribution</h3>
                    <div class="chart-container small">
                        <canvas id="replyCategoriesChart"></canvas>
                    </div>
                </div>

                <div class="chart-card">
                    <h3>Top Performing Campaigns</h3>
                    <div class="chart-container small">
                        <canvas id="topCampaignsChart"></canvas>
                    </div>
                </div>

                <div class="chart-card">
                    <h3>Replies Timeline</h3>
                    <div class="chart-container small">
                        <canvas id="repliesTimelineChart"></canvas>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        // ENHANCED DASHBOARD - Data embedded from Excel
        window.campaignData = {js_campaign_data};
        window.repliesData = {js_replies_data};

        // Color scheme for charts
        const chartColors = {{
            primary: '{colors['primary']}',
            success: '#10b981',
            warning: '#f59e0b',
            info: '#3b82f6',
            purple: '#9333ea',
            cyan: '#06b6d4'
        }};

        // Initialize dashboard on load
        window.addEventListener('DOMContentLoaded', function() {{
            initDashboard();
        }});

        // State management
        let charts = {{}};
        let selectedCampaigns = new Set();

        // Helper function to normalize campaign names for flexible matching
        function normalizeCampaignName(name) {{
            return name
                .toLowerCase()
                .replace(/\s+/g, ' ')  // Normalize multiple spaces to single space
                .replace(/\s*-\s*/g, ' ')  // Remove hyphens with surrounding spaces
                .replace('email drip sequence', '')
                .replace('batch 1', '')
                .replace('batch', '')
                .replace(/\s+/g, ' ')  // Normalize spaces again
                .trim();
        }}

        // Helper function to check if a reply matches a campaign
        function replyMatchesCampaign(replyCampaignName, campaignName) {{
            const replyNormalized = normalizeCampaignName(replyCampaignName);
            const campaignNormalized = normalizeCampaignName(campaignName);

            // Try exact match first
            if (replyNormalized === campaignNormalized) return true;

            // Try partial match (reply campaign name contains the campaign name)
            if (replyNormalized.includes(campaignNormalized)) return true;

            // Try reverse partial match (campaign name contains reply campaign name)
            if (campaignNormalized.includes(replyNormalized)) return true;

            return false;
        }}

        // Initialize dashboard
        function initDashboard() {{
            populateCampaignCheckboxes();
            updateDashboard();
        }}

        // Populate campaign checkboxes
        function populateCampaignCheckboxes() {{
            const container = document.getElementById('campaignCheckboxes');
            const campaigns = [...new Set(window.campaignData.map(c => c.Industry))].sort();

            campaigns.forEach(campaign => {{
                selectedCampaigns.add(campaign);
                const div = document.createElement('div');
                div.className = 'multi-select-option';
                div.innerHTML = `
                    <input type="checkbox" id="campaign_${{campaign.replace(/[^a-zA-Z0-9]/g, '_')}}"
                           value="${{campaign}}" checked onchange="updateCampaignSelection()">
                    <label for="campaign_${{campaign.replace(/[^a-zA-Z0-9]/g, '_')}}" style="cursor: pointer;">${{campaign}}</label>
                `;
                container.appendChild(div);
            }});
        }}

        // Toggle multi-select dropdown
        function toggleMultiSelect() {{
            const dropdown = document.getElementById('campaignDropdown');
            dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
        }}

        // Toggle select all checkbox
        function toggleSelectAll() {{
            const selectAll = document.getElementById('selectAll');
            const checkboxes = document.querySelectorAll('#campaignCheckboxes input[type="checkbox"]');
            checkboxes.forEach(cb => {{
                cb.checked = selectAll.checked;
            }});
            updateCampaignSelection();
        }}

        // Update campaign selection
        function updateCampaignSelection() {{
            selectedCampaigns.clear();
            const checkboxes = document.querySelectorAll('#campaignCheckboxes input[type="checkbox"]:checked');
            checkboxes.forEach(cb => selectedCampaigns.add(cb.value));

            const selectAll = document.getElementById('selectAll');
            const allCheckboxes = document.querySelectorAll('#campaignCheckboxes input[type="checkbox"]');
            selectAll.checked = checkboxes.length === allCheckboxes.length;

            const text = selectedCampaigns.size === 0 ? 'No campaigns selected' :
                        selectedCampaigns.size === allCheckboxes.length ? 'All Campaigns Selected' :
                        `${{selectedCampaigns.size}} Campaign(s) Selected`;
            document.getElementById('selectedCampaignsText').textContent = text;

            updateDashboard();
        }}

        // Clear all filters
        function clearFilters() {{
            document.getElementById('startDate').value = '';
            document.getElementById('endDate').value = '';
            document.querySelectorAll('#campaignCheckboxes input[type="checkbox"]').forEach(cb => cb.checked = true);
            document.getElementById('selectAll').checked = true;
            selectedCampaigns.clear();
            window.campaignData.forEach(c => selectedCampaigns.add(c.Industry));
            document.getElementById('selectedCampaignsText').textContent = 'All Campaigns Selected';
            updateDashboard();
        }}

        // Close dropdown when clicking outside
        window.addEventListener('click', function(e) {{
            if (!e.target.matches('.multi-select-button') && !e.target.matches('.multi-select-button *')) {{
                document.getElementById('campaignDropdown').style.display = 'none';
            }}
        }});

        // Filter data based on selections
        function getFilteredData() {{
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;

            return window.campaignData.filter(campaign => {{
                if (selectedCampaigns.size > 0 && !selectedCampaigns.has(campaign.Industry)) {{
                    return false;
                }}

                if (startDate && campaign['Date Sent'] < startDate) return false;
                if (endDate && campaign['Date Sent'] > endDate) return false;

                return true;
            }});
        }}

        // Update dashboard with filtered data
        function updateDashboard() {{
            const filteredData = getFilteredData();
            updateKPIs(filteredData);
            renderAllCharts(filteredData);
        }}

        // Update KPIs
        function updateKPIs(data) {{
            const totalLeads = data.reduce((sum, c) => sum + (parseFloat(c['Leads Generated']) || 0), 0);
            const totalDelivered = data.reduce((sum, c) => sum + (parseFloat(c.Delivered) || 0), 0);
            const totalOpened = data.reduce((sum, c) => sum + (parseFloat(c.Opened) || 0), 0);
            const totalClicked = data.reduce((sum, c) => sum + (parseFloat(c.Clicked) || 0), 0);

            // Filter replies based on selected campaigns and dates (with flexible matching)
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;

            const filteredReplies = window.repliesData.filter(reply => {{
                // Check if reply matches any of the selected campaigns
                const matchesCampaign = data.some(campaign =>
                    replyMatchesCampaign(reply['Email Campaign'], campaign.Industry)
                );

                if (!matchesCampaign) return false;
                if (startDate && reply.Date < startDate) return false;
                if (endDate && reply.Date > endDate) return false;
                return true;
            }});

            const totalReplied = filteredReplies.length;
            const avgOpenRate = totalDelivered > 0 ? (totalOpened / totalDelivered * 100) : 0;
            const replyRate = totalDelivered > 0 ? (totalReplied / totalDelivered * 100) : 0;
            const clickRate = totalDelivered > 0 ? (totalClicked / totalDelivered * 100) : 0;

            document.getElementById('totalLeads').textContent = totalLeads.toLocaleString();
            document.getElementById('avgOpenRate').textContent = avgOpenRate.toFixed(1) + '%';
            document.getElementById('totalClicked').textContent = totalClicked.toLocaleString();
            document.getElementById('totalReplied').textContent = totalReplied.toLocaleString();
            document.getElementById('replyRate').textContent = replyRate.toFixed(2) + '%';
            document.getElementById('clickRate').textContent = clickRate.toFixed(2) + '%';
        }}

        // Render all charts
        function renderAllCharts(data) {{
            renderPerformanceChart(data);
            renderFunnelChart(data);
            renderReplyCategoriesChart();
            renderTopCampaignsChart(data);
            renderRepliesTimelineChart();
        }}

        // Performance Chart (Enhanced Multi-Metric with Campaign Details)
        function renderPerformanceChart(data) {{
            const ctx = document.getElementById('performanceChart');
            if (charts.performance) charts.performance.destroy();

            // Sort and prepare data with campaign information
            const sortedData = [...data].sort((a, b) => a['Date Sent'].localeCompare(b['Date Sent']));

            // Get replies for filtered campaigns
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const selectedCampaignNames = new Set(data.map(c => c.Industry));

            // Create labels with campaign names
            const labels = sortedData.map(c => {{
                const date = new Date(c['Date Sent']).toLocaleDateString();
                const campaignShort = c.Industry.length > 30 ? c.Industry.substring(0, 30) + '...' : c.Industry;
                return `${{date}}\\n${{campaignShort}}`;
            }});

            // Prepare datasets with all metrics
            const leadsData = sortedData.map(c => c['Leads Generated'] || 0);
            const deliveredData = sortedData.map(c => c.Delivered || 0);
            const openedData = sortedData.map(c => c.Opened || 0);
            const clickedData = sortedData.map(c => c.Clicked || 0);

            // Calculate replies per campaign (with flexible matching using helper function)
            const repliesData = sortedData.map(c => {{
                const campaignReplies = window.repliesData.filter(r => {{
                    // Use helper function for flexible campaign matching
                    if (!replyMatchesCampaign(r['Email Campaign'], c.Industry)) return false;

                    // Apply date filters
                    if (startDate && r.Date < startDate) return false;
                    if (endDate && r.Date > endDate) return false;

                    return true;
                }});

                return campaignReplies.length;
            }});

            // Calculate open rate as percentage
            const openRateData = sortedData.map(c => ((c['Open Rate %'] || 0) * 100));

            charts.performance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [
                        {{
                            label: 'Leads Generated',
                            data: leadsData,
                            backgroundColor: 'rgba(99, 102, 241, 0.8)',
                            borderColor: chartColors.primary,
                            borderWidth: 2,
                            yAxisID: 'y',
                            type: 'bar',
                            order: 2
                        }},
                        {{
                            label: 'Delivered',
                            data: deliveredData,
                            backgroundColor: 'rgba(59, 130, 246, 0.6)',
                            borderColor: chartColors.info,
                            borderWidth: 2,
                            yAxisID: 'y',
                            type: 'bar',
                            order: 2
                        }},
                        {{
                            label: 'Opened',
                            data: openedData,
                            backgroundColor: 'rgba(16, 185, 129, 0.7)',
                            borderColor: chartColors.success,
                            borderWidth: 2,
                            yAxisID: 'y',
                            type: 'bar',
                            order: 2
                        }},
                        {{
                            label: 'Clicked',
                            data: clickedData,
                            backgroundColor: 'rgba(245, 158, 11, 0.7)',
                            borderColor: chartColors.warning,
                            borderWidth: 2,
                            yAxisID: 'y',
                            type: 'bar',
                            order: 2
                        }},
                        {{
                            label: 'Replied',
                            data: repliesData,
                            backgroundColor: 'rgba(147, 51, 234, 0.8)',
                            borderColor: chartColors.purple,
                            borderWidth: 2,
                            yAxisID: 'y',
                            type: 'bar',
                            order: 2
                        }},
                        {{
                            label: 'Open Rate %',
                            data: openRateData,
                            borderColor: '#ef4444',
                            backgroundColor: 'rgba(239, 68, 68, 0.1)',
                            borderWidth: 3,
                            yAxisID: 'y1',
                            type: 'line',
                            tension: 0.4,
                            fill: false,
                            pointRadius: 5,
                            pointHoverRadius: 7,
                            order: 1
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {{
                        mode: 'index',
                        intersect: false
                    }},
                    plugins: {{
                        legend: {{
                            display: true,
                            position: 'top',
                            labels: {{
                                usePointStyle: true,
                                padding: 15,
                                font: {{
                                    size: 12,
                                    weight: 'bold'
                                }}
                            }}
                        }},
                        tooltip: {{
                            enabled: true,
                            mode: 'index',
                            intersect: false,
                            backgroundColor: 'rgba(0, 0, 0, 0.8)',
                            titleFont: {{
                                size: 14,
                                weight: 'bold'
                            }},
                            bodyFont: {{
                                size: 13
                            }},
                            padding: 12,
                            callbacks: {{
                                title: function(context) {{
                                    const index = context[0].dataIndex;
                                    const campaign = sortedData[index];
                                    const date = new Date(campaign['Date Sent']).toLocaleDateString();
                                    return `${{date}} - ${{campaign.Industry}}`;
                                }},
                                afterTitle: function(context) {{
                                    const index = context[0].dataIndex;
                                    const campaign = sortedData[index];
                                    return `Status: ${{campaign.Status}} | CRM: ${{campaign['CRM Status']}}`;
                                }},
                                label: function(context) {{
                                    const label = context.dataset.label || '';
                                    const value = context.parsed.y;
                                    if (label === 'Open Rate %') {{
                                        return `  ${{label}}: ${{value.toFixed(1)}}%`;
                                    }}
                                    return `  ${{label}}: ${{value.toLocaleString()}}`;
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        x: {{
                            stacked: false,
                            ticks: {{
                                maxRotation: 45,
                                minRotation: 45,
                                font: {{
                                    size: 10
                                }},
                                callback: function(value, index) {{
                                    const campaign = sortedData[index];
                                    const date = new Date(campaign['Date Sent']).toLocaleDateString();
                                    const shortName = campaign.Industry.length > 20 ?
                                        campaign.Industry.substring(0, 20) + '...' :
                                        campaign.Industry;
                                    return [date, shortName];
                                }}
                            }},
                            grid: {{
                                display: false
                            }}
                        }},
                        y: {{
                            type: 'linear',
                            position: 'left',
                            stacked: false,
                            title: {{
                                display: true,
                                text: 'Count (Leads, Delivered, Opened, Clicked, Replied)',
                                font: {{
                                    size: 12,
                                    weight: 'bold'
                                }}
                            }},
                            beginAtZero: true,
                            grid: {{
                                color: 'rgba(0, 0, 0, 0.05)'
                            }}
                        }},
                        y1: {{
                            type: 'linear',
                            position: 'right',
                            title: {{
                                display: true,
                                text: 'Open Rate (%)',
                                font: {{
                                    size: 12,
                                    weight: 'bold'
                                }},
                                color: '#ef4444'
                            }},
                            beginAtZero: true,
                            max: 100,
                            grid: {{
                                drawOnChartArea: false
                            }},
                            ticks: {{
                                callback: function(value) {{
                                    return value + '%';
                                }},
                                color: '#ef4444'
                            }}
                        }}
                    }}
                }}
            }});
        }}

        // Funnel Chart
        function renderFunnelChart(data) {{
            const ctx = document.getElementById('funnelChart');
            if (charts.funnel) charts.funnel.destroy();

            const delivered = data.reduce((sum, c) => sum + (parseFloat(c.Delivered) || 0), 0);
            const opened = data.reduce((sum, c) => sum + (parseFloat(c.Opened) || 0), 0);
            const clicked = data.reduce((sum, c) => sum + (parseFloat(c.Clicked) || 0), 0);

            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;

            const filteredReplies = window.repliesData.filter(r => {{
                // Use flexible campaign matching
                const matchesCampaign = data.some(campaign =>
                    replyMatchesCampaign(r['Email Campaign'], campaign.Industry)
                );

                if (!matchesCampaign) return false;
                if (startDate && r.Date < startDate) return false;
                if (endDate && r.Date > endDate) return false;
                return true;
            }});

            charts.funnel = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: ['Delivered', 'Opened', 'Clicked', 'Replied'],
                    datasets: [{{
                        label: 'Conversion Funnel',
                        data: [delivered, opened, clicked, filteredReplies.length],
                        backgroundColor: [chartColors.primary, chartColors.success, chartColors.warning, chartColors.info]
                    }}]
                }},
                options: {{
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false
                }}
            }});
        }}

        // Reply Categories Chart
        function renderReplyCategoriesChart() {{
            const ctx = document.getElementById('replyCategoriesChart');
            if (charts.replyCategories) charts.replyCategories.destroy();

            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const filteredData = getFilteredData();

            const filteredReplies = window.repliesData.filter(r => {{
                // Use flexible campaign matching
                const matchesCampaign = filteredData.some(campaign =>
                    replyMatchesCampaign(r['Email Campaign'], campaign.Industry)
                );

                if (!matchesCampaign) return false;
                if (startDate && r.Date < startDate) return false;
                if (endDate && r.Date > endDate) return false;
                return true;
            }});

            const categories = {{}};
            filteredReplies.forEach(r => {{
                const cat = r.Category || 'Unknown';
                categories[cat] = (categories[cat] || 0) + 1;
            }});

            // Function to get color based on category name
            const getCategoryColor = (category) => {{
                const cat = category.toLowerCase();

                // Not Interested - Gray
                if (cat.includes('not interested')) return '#9ca3af';

                // Engaged - Interested - Green
                if (cat.includes('engaged') && cat.includes('interested')) return '#10b981';

                // Hot Lead - Meeting Scheduled - Blue
                if (cat.includes('hot lead') || cat.includes('meeting scheduled')) return '#3b82f6';

                // No Action Needed - Purple
                if (cat.includes('no action')) return '#8b5cf6';

                // Follow-Up Needed - Beige
                if (cat.includes('follow') || cat.includes('follow-up')) return '#d4a574';

                // Pipelined / Circle Back - Cyan (light blue)
                if (cat.includes('pipeline') || cat.includes('circle back')) return '#06b6d4';

                // Invalid / Wrong Contact - Dark Gray
                if (cat.includes('invalid') || cat.includes('wrong contact')) return '#4b5563';

                // Sold - Gold (maximum achievement)
                if (cat.includes('sold')) return '#fbbf24';

                // Existing Client - Pink
                if (cat.includes('existing client')) return '#ec4899';

                // Default color for unknown categories
                return '#94a3b8';
            }};

            // Generate colors array based on actual categories
            const categoryLabels = Object.keys(categories);
            const categoryColors = categoryLabels.map(cat => getCategoryColor(cat));

            charts.replyCategories = new Chart(ctx, {{
                type: 'doughnut',
                data: {{
                    labels: categoryLabels,
                    datasets: [{{
                        data: Object.values(categories),
                        backgroundColor: categoryColors,
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'right',
                            labels: {{
                                padding: 12,
                                font: {{
                                    size: 11
                                }},
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }}
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    const label = context.label || '';
                                    const value = context.parsed;
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = ((value / total) * 100).toFixed(1);
                                    return `  ${{label}}: ${{value}} (${{percentage}}%)`;
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        }}

        // Top Campaigns Chart
        function renderTopCampaignsChart(data) {{
            const ctx = document.getElementById('topCampaignsChart');
            if (charts.topCampaigns) charts.topCampaigns.destroy();

            const sorted = [...data].sort((a, b) => (b['Open Rate %'] || 0) - (a['Open Rate %'] || 0)).slice(0, 5);
            const labels = sorted.map(c => c.Industry.substring(0, 20) + (c.Industry.length > 20 ? '...' : ''));
            const rates = sorted.map(c => ((c['Open Rate %'] || 0) * 100).toFixed(1));

            charts.topCampaigns = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: 'Open Rate %',
                        data: rates,
                        backgroundColor: chartColors.success
                    }}]
                }},
                options: {{
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false
                }}
            }});
        }}

        // Replies Timeline Chart
        function renderRepliesTimelineChart() {{
            const ctx = document.getElementById('repliesTimelineChart');
            if (charts.repliesTimeline) charts.repliesTimeline.destroy();

            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const filteredData = getFilteredData();

            const filteredReplies = window.repliesData.filter(r => {{
                // Use flexible campaign matching
                const matchesCampaign = filteredData.some(campaign =>
                    replyMatchesCampaign(r['Email Campaign'], campaign.Industry)
                );

                if (!matchesCampaign) return false;
                if (startDate && r.Date < startDate) return false;
                if (endDate && r.Date > endDate) return false;
                return true;
            }});

            const dateGroups = {{}};
            filteredReplies.forEach(r => {{
                let dateStr = r.Date;
                if (dateStr && dateStr.includes(' ')) {{
                    dateStr = dateStr.split(' ')[0];
                }}
                if (dateStr) {{
                    const parts = dateStr.split('-');
                    if (parts.length === 3 && parts[0].length === 2) {{
                        dateStr = `2026-${{parts[1]}}-${{parts[0]}}`;
                    }}
                    dateGroups[dateStr] = (dateGroups[dateStr] || 0) + 1;
                }}
            }});

            const sortedDates = Object.keys(dateGroups).sort();
            const counts = sortedDates.map(d => dateGroups[d]);
            const labels = sortedDates.map(d => new Date(d).toLocaleDateString());

            charts.repliesTimeline = new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: 'Replies per Day',
                        data: counts,
                        borderColor: chartColors.info,
                        backgroundColor: 'rgba(59, 130, 246, 0.1)',
                        tension: 0.3,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        x: {{ ticks: {{ maxRotation: 45, minRotation: 45 }} }}
                    }}
                }}
            }});
        }}
    </script>
</body>
</html>'''

    return html

def generate_medicar_replies_html(company_name, company_data, replies_data):
    """Genera el HTML de replies para Medicar con distinción entre Calls y Replies"""

    colors = company_data['colors']

    # Separar entre Calls y Replies
    calls = [r for r in replies_data if r.get('Reply/Calls', '').lower().strip() == 'call']
    emails = [r for r in replies_data if r.get('Reply/Calls', '').lower().strip() == 'reply']
    total_count = len(replies_data)

    # Contar por categoría para todos
    real_category_counts = {}
    for reply in replies_data:
        category = reply.get('Category', 'Unknown')
        if category not in real_category_counts:
            real_category_counts[category] = 0
        real_category_counts[category] += 1

    sorted_categories = sorted(real_category_counts.keys(), key=lambda x: (x.lower() != 'sold', x.lower()))

    # Función auxiliar para formatear fecha
    def parse_date(date_str):
        if not date_str:
            return '1900-01-01'
        date_str = str(date_str).strip().split(' ')[0]
        parts = date_str.split('-')
        if len(parts) == 3:
            if len(parts[0]) == 2:
                return f"{parts[2]}-{parts[1]}-{parts[0]}"
            else:
                return date_str
        return '1900-01-01'

    def format_date_display(date_str):
        if not date_str:
            return 'Unknown'
        date_str = str(date_str).strip().split(' ')[0]
        parts = date_str.split('-')
        if len(parts) == 3:
            if len(parts[0]) == 2:
                return f"{parts[0]}/{parts[1]}/{parts[2]}"
            else:
                return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return 'Unknown'

    # Generar tarjetas para ambos tipos
    def generate_cards(items, item_type):
        cards_html = ''
        sorted_items = sorted(items, key=lambda x: parse_date(x.get('Date', '')), reverse=True)

        for item in sorted_items:
            who = item.get('Who', 'Unknown')
            email = item.get('Email', '')
            message = item.get('Reply/Calls', '')
            action = item.get('Action Taken', '')
            subject = item.get('Subject if email ', '')
            campaign = item.get('Email Campaign', '')
            category = item.get('Category', 'Neutral')
            date_raw = item.get('Date', '')
            date_display = format_date_display(date_raw)
            date_sortable = parse_date(date_raw)
            cat_class = get_category_class(category)

            type_badge = '📞 Call' if item_type == 'call' else '📧 Email Reply'

            cards_html += f'''
            <div class="response-card {cat_class}" data-category="{cat_class}" data-date="{date_sortable}" data-type="{item_type}">
                <div class="response-header">
                    <div>
                        <div class="response-who">{who}</div>
                        <div class="response-email">{email}</div>
                    </div>
                    <div style="display: flex; gap: 10px; align-items: center;">
                        <span class="type-badge {item_type}">{type_badge}</span>
                        <div class="response-category {cat_class}">{category}</div>
                    </div>
                </div>
                <div class="response-date">
                    <span class="date-icon">📅</span> {date_display}
                </div>
                <div class="response-details">
                    {f'<div class="response-detail-item"><span class="response-detail-label">Subject:</span><span class="response-detail-value">{subject}</span></div>' if subject else ''}
                    <div class="response-detail-item">
                        <span class="response-detail-label">Campaign:</span>
                        <span class="response-detail-value">{campaign}</span>
                    </div>
                    <div class="response-detail-item">
                        <span class="response-detail-label">Action Taken:</span>
                        <span class="response-detail-value">{action}</span>
                    </div>
                </div>
            </div>
            '''

        return cards_html

    all_cards = generate_cards(calls, 'call') + generate_cards(emails, 'reply')

    # Generar botones de filtro por categoría
    filter_buttons_html = '<button class="filter-btn all active" onclick="filterResponses(\'all\')">All ({total})</button>'.format(total=total_count)
    filter_buttons_html += f'<button class="filter-btn call" onclick="filterByType(\'call\')">📞 Calls ({len(calls)})</button>'
    filter_buttons_html += f'<button class="filter-btn reply" onclick="filterByType(\'reply\')">📧 Email Replies ({len(emails)})</button>'

    for category in sorted_categories:
        count = real_category_counts[category]
        cat_class = get_category_class(category)
        filter_buttons_html += f'<button class="filter-btn {cat_class}" onclick="filterResponses(\'{cat_class}\')">{category} ({count})</button>'

    # HTML completo (versión abreviada por espacio)
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Response Analysis - {company_name} (Calls & Replies)</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: {colors['body_gradient']}; padding: 20px; min-height: 100vh; }}
        .container {{ max-width: 1400px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1); overflow: hidden; }}
        header {{ background: {colors['header_gradient']}; color: white; padding: 30px 40px; display: flex; justify-content: space-between; align-items: center; }}
        header h1 {{ font-size: 32px; font-weight: 600; }}
        .back-btn {{ padding: 12px 25px; background: rgba(255, 255, 255, 0.2); color: white; border: 2px solid white; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.3s ease; text-decoration: none; display: inline-block; }}
        .back-btn:hover {{ background: white; color: {colors['primary']}; }}
        .content {{ padding: 40px; }}
        .summary-stats {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 40px; }}
        .stat-card {{ background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); padding: 25px; border-radius: 10px; text-align: center; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08); }}
        .stat-card h3 {{ font-size: 14px; color: #6c757d; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 0.5px; }}
        .stat-card .stat-value {{ font-size: 36px; font-weight: 700; color: {colors['primary']}; }}
        .type-badge {{ padding: 6px 12px; border-radius: 20px; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }}
        .type-badge.call {{ background: #10b981; color: white; }}
        .type-badge.reply {{ background: #3b82f6; color: white; }}
        .response-card {{ background: white; border-radius: 10px; padding: 25px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08); border-left: 5px solid; transition: transform 0.2s ease; }}
        .response-card:hover {{ transform: translateX(5px); }}
        .response-card.not-interested {{ border-left-color: #9ca3af; }}
        .response-card.engaged-interested {{ border-left-color: #10b981; }}
        .response-card.hot-lead {{ border-left-color: #3b82f6; }}
        .response-card.no-action {{ border-left-color: #8b5cf6; }}
        .response-card.follow-up {{ border-left-color: #d4a574; }}
        .response-card.pipelined {{ border-left-color: #06b6d4; }}
        .response-card.invalid {{ border-left-color: #4b5563; }}
        .response-card.sold {{ border-left-color: #fbbf24; }}
        .response-card.existing-client {{ border-left-color: #ec4899; }}
        .response-card.neutral {{ border-left-color: #94a3b8; }}
        .response-header {{ display: flex; justify-content: space-between; align-items: start; margin-bottom: 15px; }}
        .response-who {{ font-size: 20px; font-weight: 700; color: #2c3e50; }}
        .response-email {{ font-size: 14px; color: #7f8c8d; margin-top: 5px; }}
        .response-category {{ padding: 6px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }}
        .response-category.not-interested {{ background: #9ca3af; color: white; }}
        .response-category.engaged-interested {{ background: #10b981; color: white; }}
        .response-category.hot-lead {{ background: #3b82f6; color: white; }}
        .response-category.no-action {{ background: #8b5cf6; color: white; }}
        .response-category.follow-up {{ background: #d4a574; color: white; }}
        .response-category.pipelined {{ background: #06b6d4; color: white; }}
        .response-category.invalid {{ background: #4b5563; color: white; }}
        .response-category.sold {{ background: linear-gradient(135deg, #fbbf24, #f59e0b); color: #000; font-weight: 700; }}
        .response-category.existing-client {{ background: #ec4899; color: white; }}
        .response-category.neutral {{ background: #94a3b8; color: white; }}
        .response-date {{ background: #f0f9ff; border: 1px solid #bfdbfe; padding: 8px 12px; border-radius: 6px; font-size: 13px; font-weight: 600; color: #1e40af; margin-bottom: 15px; display: inline-block; }}
        .response-details {{ margin-bottom: 15px; }}
        .response-detail-item {{ display: flex; margin-bottom: 8px; font-size: 14px; }}
        .response-detail-label {{ font-weight: 600; color: #2c3e50; min-width: 120px; }}
        .response-detail-value {{ color: #5a6c7d; }}
        .filter-buttons {{ display: flex; gap: 10px; margin-bottom: 30px; flex-wrap: wrap; }}
        .filter-btn {{ padding: 10px 20px; border: 2px solid #e9ecef; background: white; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.3s ease; }}
        .filter-btn:hover {{ transform: translateY(-2px); }}
        .filter-btn.active {{ color: white; }}
        .filter-btn.all.active {{ background: #2c3e50; border-color: #2c3e50; }}
        .filter-btn.call.active {{ background: #10b981; border-color: #10b981; color: white; }}
        .filter-btn.reply.active {{ background: #3b82f6; border-color: #3b82f6; color: white; }}
        .filter-btn.not-interested.active {{ background: #9ca3af; border-color: #9ca3af; }}
        .filter-btn.engaged-interested.active {{ background: #10b981; border-color: #10b981; }}
        .filter-btn.hot-lead.active {{ background: #3b82f6; border-color: #3b82f6; }}
        .filter-btn.no-action.active {{ background: #8b5cf6; border-color: #8b5cf6; }}
        .filter-btn.follow-up.active {{ background: #d4a574; border-color: #d4a574; }}
        .filter-btn.pipelined.active {{ background: #06b6d4; border-color: #06b6d4; }}
        .filter-btn.invalid.active {{ background: #4b5563; border-color: #4b5563; }}
        .filter-btn.sold.active {{ background: #fbbf24; border-color: #fbbf24; color: #000; }}
        .filter-btn.existing-client.active {{ background: #ec4899; border-color: #ec4899; }}
        .filter-btn.neutral.active {{ background: #94a3b8; border-color: #94a3b8; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📞📧 Response Analysis - {company_name}</h1>
            <a href="dashboard_{company_name.lower()}.html" class="back-btn">← Back to Dashboard</a>
        </header>
        <div class="content">
            <div class="summary-stats">
                <div class="stat-card"><h3>Total Responses</h3><div class="stat-value">{total_count}</div></div>
                <div class="stat-card"><h3>📞 Phone Calls</h3><div class="stat-value">{len(calls)}</div></div>
                <div class="stat-card"><h3>📧 Email Replies</h3><div class="stat-value">{len(emails)}</div></div>
            </div>
            <div class="filter-buttons">{filter_buttons_html}</div>
            <div class="responses-section">{all_cards}</div>
        </div>
    </div>
    <script>
        function filterResponses(category) {{
            const cards = document.querySelectorAll('.response-card');
            const buttons = document.querySelectorAll('.filter-btn');
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
            cards.forEach(card => {{
                card.style.display = category === 'all' || card.dataset.category === category ? 'block' : 'none';
            }});
        }}
        function filterByType(type) {{
            const cards = document.querySelectorAll('.response-card');
            const buttons = document.querySelectorAll('.filter-btn');
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
            cards.forEach(card => {{
                card.style.display = card.dataset.type === type ? 'block' : 'none';
            }});
        }}
    </script>
</body>
</html>'''
    return html

def generate_sales_html(company_name, company_data, sales_data):
    """Genera el HTML de SALES para Medicar"""
    colors = company_data['colors']
    total_sales = len(sales_data)
    total_amount = 0
    for sale in sales_data:
        try:
            amount_clean = sale.get('Amount', '0').replace('$', '').replace(',', '').strip()
            total_amount += float(amount_clean)
        except:
            pass

    def format_date_display(date_str):
        if not date_str:
            return 'Unknown'
        date_str = str(date_str).strip().split(' ')[0]
        parts = date_str.split('-')
        if len(parts) == 3:
            return f"{parts[0]}/{parts[1]}/{parts[2]}" if len(parts[0]) == 2 else f"{parts[2]}/{parts[1]}/{parts[0]}"
        return date_str

    sales_cards_html = ''
    for i, sale in enumerate(sales_data, 1):
        who = sale.get('Who', 'Unknown')
        email = sale.get('Email', '')
        message = sale.get('Message', '')
        subject = sale.get('Subject', '')
        amount = sale.get('Amount', '$0')
        action = sale.get('Action Taken', '')
        campaign = sale.get('Email Campaign Title', '')
        category = sale.get('Category', '')
        date_reply = format_date_display(sale.get('Date of Reply', ''))
        date_sent = format_date_display(sale.get('When was this email campaign sent?', ''))

        sales_cards_html += f'''
        <div class="sale-card">
            <div class="sale-header">
                <div class="sale-number">💰 Sale #{i}</div>
                <div class="sale-amount">{amount}</div>
            </div>
            <div class="sale-info">
                <div class="sale-client"><div class="client-name">{who}</div><div class="client-email">{email}</div></div>
                <div class="sale-dates">
                    <div class="date-item"><span class="date-label">Campaign Sent:</span><span class="date-value">{date_sent}</span></div>
                    <div class="date-item"><span class="date-label">Reply Received:</span><span class="date-value">{date_reply}</span></div>
                </div>
            </div>
            <div class="sale-details">
                <div class="sale-detail-item"><span class="sale-detail-label">📧 Campaign:</span><span class="sale-detail-value">{campaign}</span></div>
                {f'<div class="sale-detail-item"><span class="sale-detail-label">📝 Subject:</span><span class="sale-detail-value">{subject}</span></div>' if subject else ''}
                {f'<div class="sale-detail-item"><span class="sale-detail-label">🏷️ Category:</span><span class="sale-detail-value">{category}</span></div>' if category else ''}
                <div class="sale-detail-item"><span class="sale-detail-label">✅ Action Taken:</span><span class="sale-detail-value">{action}</span></div>
            </div>
            {f'<div class="sale-message">{message}</div>' if message else ''}
        </div>'''

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>💰 Sales - {company_name}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: {colors['body_gradient']}; padding: 20px; min-height: 100vh; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1); overflow: hidden; }}
        header {{ background: linear-gradient(135deg, #fbbf24 0%, #f59e0b 100%); color: #1a1a1a; padding: 40px; text-align: center; position: relative; }}
        header h1 {{ font-size: 42px; font-weight: 700; margin-bottom: 10px; }}
        header p {{ font-size: 18px; opacity: 0.9; }}
        .back-btn {{ position: absolute; top: 40px; left: 40px; padding: 12px 25px; background: rgba(255, 255, 255, 0.3); color: #1a1a1a; border: 2px solid #1a1a1a; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.3s ease; text-decoration: none; display: inline-block; }}
        .back-btn:hover {{ background: #1a1a1a; color: #fbbf24; }}
        .content {{ padding: 40px; }}
        .summary-stats {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 40px; }}
        .stat-card {{ background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); padding: 30px; border-radius: 12px; text-align: center; box-shadow: 0 4px 12px rgba(251, 191, 36, 0.2); border: 2px solid #fbbf24; }}
        .stat-card h3 {{ font-size: 16px; color: #92400e; margin-bottom: 15px; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }}
        .stat-card .stat-value {{ font-size: 48px; font-weight: 700; color: #92400e; }}
        .sales-grid {{ display: grid; gap: 30px; }}
        .sale-card {{ background: white; border-radius: 12px; padding: 30px; box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1); border-left: 6px solid #fbbf24; transition: all 0.3s ease; }}
        .sale-card:hover {{ transform: translateX(5px); box-shadow: 0 8px 25px rgba(251, 191, 36, 0.3); }}
        .sale-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 2px solid #fef3c7; }}
        .sale-number {{ font-size: 24px; font-weight: 700; color: #92400e; }}
        .sale-amount {{ font-size: 32px; font-weight: 700; color: #fbbf24; background: linear-gradient(135deg, #fbbf24, #f59e0b); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }}
        .sale-info {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
        .sale-client {{ background: #f8f9fa; padding: 15px; border-radius: 8px; }}
        .client-name {{ font-size: 20px; font-weight: 700; color: #2c3e50; margin-bottom: 5px; }}
        .client-email {{ font-size: 14px; color: #7f8c8d; }}
        .sale-dates {{ background: #f0f9ff; padding: 15px; border-radius: 8px; }}
        .date-item {{ display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 14px; }}
        .date-item:last-child {{ margin-bottom: 0; }}
        .date-label {{ font-weight: 600; color: #2c3e50; }}
        .date-value {{ color: #1e40af; font-weight: 600; }}
        .sale-details {{ background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 15px; }}
        .sale-detail-item {{ display: flex; margin-bottom: 10px; font-size: 15px; }}
        .sale-detail-item:last-child {{ margin-bottom: 0; }}
        .sale-detail-label {{ font-weight: 600; color: #2c3e50; min-width: 130px; }}
        .sale-detail-value {{ color: #5a6c7d; flex: 1; }}
        .sale-message {{ background: #fef3c7; padding: 20px; border-radius: 8px; font-size: 15px; line-height: 1.6; color: #2c3e50; font-style: italic; border-left: 4px solid #fbbf24; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <a href="dashboard_{company_name.lower()}.html" class="back-btn">← Back to Dashboard</a>
            <h1>💰 SALES ACHIEVED</h1>
            <p>Direct sales generated from email campaigns</p>
        </header>
        <div class="content">
            <div class="summary-stats">
                <div class="stat-card"><h3>🎯 Total Sales</h3><div class="stat-value">{total_sales}</div></div>
                <div class="stat-card"><h3>💵 Total Revenue</h3><div class="stat-value">${total_amount:,.2f}</div></div>
            </div>
            <div class="sales-grid">{sales_cards_html}</div>
        </div>
    </div>
</body>
</html>'''
    return html

def main():
    print("=" * 80)
    print("ACTUALIZANDO TODOS LOS DASHBOARDS Y REPLIES")
    print("=" * 80)
    print(f"Leyendo: {EXCEL_FILE}\n")

    for company_name, company_data in COMPANIES.items():
        print(f"[{company_name}]")

        # Leer datos de campañas (con filtro opcional para sheets compartidas)
        filter_company = company_data.get('filter_company')
        data_records = read_company_data(company_data['sheet'], filter_company)

        if not data_records:
            print(f"  ADVERTENCIA: No hay datos para {company_name}")
            continue

        total_leads = sum(int(r.get('Leads Generated', 0)) for r in data_records)
        print(f"  Total Leads: {total_leads:,}")

        # Leer replies si existen
        replies_data = []
        if company_data['replies_sheet']:
            replies_data = read_replies_data(company_data['replies_sheet'])
            print(f"  Replies: {len(replies_data)}")

            if replies_data:
                # Para iDrivio (que incluye Medicar), generar versión especial con distinción Calls/Replies
                if company_name == 'iDrivio':
                    replies_html = generate_medicar_replies_html(company_name, company_data, replies_data)
                else:
                    # Generar HTML de replies normal
                    replies_html = generate_replies_html(company_name, company_data, replies_data)

                replies_filename = f"replies_{company_name.lower().replace(' ', '')}.html"
                with open(replies_filename, 'w', encoding='utf-8') as f:
                    f.write(replies_html)
                print(f"  CREADO: {replies_filename}")

        # Leer y generar página de SALES si existe (solo Medicar por ahora)
        if company_data.get('sales_sheet'):
            sales_data = read_sales_data(company_data['sales_sheet'])
            print(f"  Sales: {len(sales_data)}")

            if sales_data:
                sales_html = generate_sales_html(company_name, company_data, sales_data)
                sales_filename = f"sales_{company_name.lower().replace(' ', '')}.html"
                with open(sales_filename, 'w', encoding='utf-8') as f:
                    f.write(sales_html)
                print(f"  CREADO: {sales_filename}")

        # Generar HTML del dashboard (usar versión mejorada para TeamFicient)
        if company_name == 'TeamFicient':
            html_content = generate_improved_dashboard_teamficient(company_name, company_data, data_records, replies_data)
        else:
            html_content = generate_dashboard_html(company_name, company_data, data_records, len(replies_data))

        # Guardar archivo
        filename = f"dashboard_{company_name.lower().replace(' ', '')}.html"
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)

        print(f"  CREADO: {filename}")
        print()

    print("=" * 80)
    print("COMPLETADO!")
    print("Todos los dashboards y paginas de replies han sido actualizados.")
    print("=" * 80)

if __name__ == '__main__':
    main()
