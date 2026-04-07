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
        'replies_sheet': None,
        'colors': {
            'body_gradient': 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
            'header_gradient': 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
            'primary': '#f59e0b',
            'secondary': '#d97706',
            'accent': '#b45309'
        }
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

def read_company_data(sheet_name):
    """Lee los datos de una compañía desde el Excel - SOLO campañas con Status = 'Sent'"""
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, skiprows=[0])
        df = df.dropna(how='all')

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

def get_category_class(category):
    """Mapea categoría a clase CSS"""
    category_lower = category.lower()
    if 'not interested' in category_lower or 'unsubscribe' in category_lower or 'negative' in category_lower:
        return 'negative'
    elif 'hot' in category_lower or 'meeting' in category_lower or 'interested' in category_lower:
        return 'hot-lead'
    elif 'engaged' in category_lower or 'positive' in category_lower:
        return 'interested'
    else:
        return 'neutral'

def generate_replies_html(company_name, company_data, replies_data):
    """Genera el HTML de replies"""

    colors = company_data['colors']

    # Calcular estadísticas
    total_replies = len(replies_data)
    category_counts = {'negative': 0, 'hot-lead': 0, 'interested': 0, 'neutral': 0}

    for reply in replies_data:
        category = reply.get('Category', 'Neutral')
        cat_class = get_category_class(category)
        category_counts[cat_class] += 1

    # Generar HTML de tarjetas de respuestas
    replies_html = ''
    for reply in replies_data:
        who = reply.get('Who', 'Unknown')
        email = reply.get('Email', '')
        message = reply.get('Reply/Calls', '')
        action = reply.get('Action Taken', '')
        subject = reply.get('Subject if email ', reply.get('Subject', ''))
        campaign = reply.get('Email Campaign', '')
        category = reply.get('Category', 'Neutral')
        cat_class = get_category_class(category)

        replies_html += f'''
        <div class="response-card {cat_class}" data-category="{cat_class}">
            <div class="response-header">
                <div>
                    <div class="response-who">{who}</div>
                    <div class="response-email">{email}</div>
                </div>
                <div class="response-category {cat_class}">{category}</div>
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

        .stat-item.negative {{
            border-left-color: #EF4444;
        }}

        .stat-item.interested {{
            border-left-color: #3B82F6;
        }}

        .stat-item.hot-lead {{
            border-left-color: #10B981;
        }}

        .stat-item.neutral {{
            border-left-color: #8B5CF6;
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

        .response-category.negative {{
            background: #EF4444;
            color: white;
        }}

        .response-category.interested {{
            background: #3B82F6;
            color: white;
        }}

        .response-category.hot-lead {{
            background: #10B981;
            color: white;
        }}

        .response-category.neutral {{
            background: #8B5CF6;
            color: white;
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

        .filter-btn.negative.active {{
            background: #EF4444;
            border-color: #EF4444;
        }}

        .filter-btn.interested.active {{
            background: #3B82F6;
            border-color: #3B82F6;
        }}

        .filter-btn.hot-lead.active {{
            background: #10B981;
            border-color: #10B981;
        }}

        .filter-btn.neutral.active {{
            background: #8B5CF6;
            border-color: #8B5CF6;
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
                    <h2>Response Statistics</h2>
                    <div class="stats-grid">
                        <div class="stat-item negative">
                            <div class="stat-label">Negative/Unsubscribe</div>
                            <div class="stat-value">{category_counts['negative']} <span class="stat-percentage">({category_counts['negative']/total_replies*100 if total_replies > 0 else 0:.0f}%)</span></div>
                        </div>
                        <div class="stat-item interested">
                            <div class="stat-label">Interested/Engaged</div>
                            <div class="stat-value">{category_counts['interested']} <span class="stat-percentage">({category_counts['interested']/total_replies*100 if total_replies > 0 else 0:.0f}%)</span></div>
                        </div>
                        <div class="stat-item hot-lead">
                            <div class="stat-label">Hot Lead/Meeting</div>
                            <div class="stat-value">{category_counts['hot-lead']} <span class="stat-percentage">({category_counts['hot-lead']/total_replies*100 if total_replies > 0 else 0:.0f}%)</span></div>
                        </div>
                        <div class="stat-item neutral">
                            <div class="stat-label">Neutral/Acknowledgment</div>
                            <div class="stat-value">{category_counts['neutral']} <span class="stat-percentage">({category_counts['neutral']/total_replies*100 if total_replies > 0 else 0:.0f}%)</span></div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="responses-section">
                <h2>All Responses ({total_replies})</h2>

                <div class="filter-buttons">
                    <button class="filter-btn all active" onclick="filterResponses('all')">All ({total_replies})</button>
                    <button class="filter-btn negative" onclick="filterResponses('negative')">Negative ({category_counts['negative']})</button>
                    <button class="filter-btn hot-lead" onclick="filterResponses('hot-lead')">Hot Lead ({category_counts['hot-lead']})</button>
                    <button class="filter-btn interested" onclick="filterResponses('interested')">Interested ({category_counts['interested']})</button>
                    <button class="filter-btn neutral" onclick="filterResponses('neutral')">Neutral ({category_counts['neutral']})</button>
                </div>

                <div id="responsesContainer">
                    {replies_html}
                </div>
            </div>
        </div>
    </div>

    <script>
        // Create response distribution chart
        const ctx = document.getElementById('responseChart');
        new Chart(ctx, {{
            type: 'doughnut',
            data: {{
                labels: ['Negative/Unsubscribe', 'Interested/Engaged', 'Hot Lead/Meeting', 'Neutral/Acknowledgment'],
                datasets: [{{
                    data: [{category_counts['negative']}, {category_counts['interested']}, {category_counts['hot-lead']}, {category_counts['neutral']}],
                    backgroundColor: [
                        '#EF4444',
                        '#3B82F6',
                        '#10B981',
                        '#8B5CF6'
                    ],
                    borderWidth: 0
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        position: 'bottom',
                        labels: {{
                            padding: 20,
                            font: {{
                                size: 13
                            }}
                        }}
                    }},
                    tooltip: {{
                        callbacks: {{
                            label: function(context) {{
                                const label = context.label || '';
                                const value = context.parsed || 0;
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                                return label + ': ' + value + ' (' + percentage + '%)';
                            }}
                        }}
                    }}
                }}
            }}
        }});

        // Filter responses
        function filterResponses(category) {{
            const cards = document.querySelectorAll('.response-card');
            const buttons = document.querySelectorAll('.filter-btn');

            // Update button states
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');

            // Filter cards
            cards.forEach(card => {{
                if (category === 'all') {{
                    card.style.display = 'block';
                }} else {{
                    if (card.dataset.category === category) {{
                        card.style.display = 'block';
                    }} else {{
                        card.style.display = 'none';
                    }}
                }}
            }});
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
                <label for="dateFilter">Date</label>
                <select id="dateFilter">
                    <option value="">All Dates</option>
                </select>
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
            const dates = [...new Set(processedData.map(d => d.date).filter(d => d && d !== 'NaT'))].sort();

            const campaignFilter = document.getElementById('campaignFilter');
            const dateFilter = document.getElementById('dateFilter');

            campaigns.forEach(campaign => {{
                const option = document.createElement('option');
                option.value = campaign;
                option.textContent = campaign;
                campaignFilter.appendChild(option);
            }});

            dates.forEach(date => {{
                const option = document.createElement('option');
                option.value = date;
                option.textContent = formatDisplayDate(date);
                dateFilter.appendChild(option);
            }});

            campaignFilter.addEventListener('change', updateDashboard);
            dateFilter.addEventListener('change', updateDashboard);
        }}

        function clearFilters() {{
            document.getElementById('campaignFilter').value = '';
            document.getElementById('dateFilter').value = '';
            updateDashboard();
        }}

        function getFilteredData() {{
            const campaignFilter = document.getElementById('campaignFilter').value;
            const dateFilter = document.getElementById('dateFilter').value;
            const processedData = processData(allData);

            return processedData.filter(item => {{
                const matchesCampaign = !campaignFilter || item.campaign === campaignFilter;
                const matchesDate = !dateFilter || item.date === dateFilter;
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

def main():
    print("=" * 80)
    print("ACTUALIZANDO TODOS LOS DASHBOARDS Y REPLIES")
    print("=" * 80)
    print(f"Leyendo: {EXCEL_FILE}\n")

    for company_name, company_data in COMPANIES.items():
        print(f"[{company_name}]")

        # Leer datos de campañas
        data_records = read_company_data(company_data['sheet'])

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
                # Generar HTML de replies
                replies_html = generate_replies_html(company_name, company_data, replies_data)
                replies_filename = f"replies_{company_name.lower().replace(' ', '')}.html"
                with open(replies_filename, 'w', encoding='utf-8') as f:
                    f.write(replies_html)
                print(f"  CREADO: {replies_filename}")

        # Generar HTML del dashboard
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
