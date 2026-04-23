#!/usr/bin/env python3
"""
Genera el dashboard mejorado de TeamFicient con todas las funcionalidades avanzadas
"""

import pandas as pd
import json
from datetime import datetime

def generate_improved_dashboard():
    # Leer datos
    df_campaign = pd.read_excel('Email & Leads Campaign Summary & Plan for All Companies.xlsx',
                                 sheet_name='TeamFicient', skiprows=[0])
    df_campaign = df_campaign[df_campaign['Status'].str.strip().str.lower() == 'sent']

    campaign_data = []
    for _, row in df_campaign.iterrows():
        record = {}
        for col in df_campaign.columns:
            val = row[col]
            if pd.notna(val):
                if isinstance(val, (pd.Timestamp, datetime)):
                    record[col] = val.strftime('%Y-%m-%d')
                elif isinstance(val, (int, float)):
                    record[col] = float(val) if '.' in str(val) or 'Rate' in col else int(val)
                else:
                    record[col] = str(val).strip()
        if record:
            campaign_data.append(record)

    df_replies = pd.read_excel('Email & Leads Campaign Summary & Plan for All Companies.xlsx',
                               sheet_name='TeamFicient Replies', header=0)
    replies_data = []
    for _, row in df_replies.iterrows():
        reply = {}
        for col in df_replies.columns:
            val = row[col]
            if pd.notna(val):
                reply[col] = str(val).strip()
        if reply:
            replies_data.append(reply)

    campaign_json = json.dumps(campaign_data, ensure_ascii=False, indent=2)
    replies_json = json.dumps(replies_data, ensure_ascii=False, indent=2)

    print(f"Generando dashboard con {len(campaign_data)} campañas y {len(replies_data)} replies...")

    # Generar el HTML mejorado completo
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TeamFicient - Email Campaign Dashboard (Enhanced)</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>
    <h1>Dashboard mejorado generado con {len(campaign_data)} campañas</h1>
    <p>Total Leads: {sum(c.get('Leads Generated', 0) for c in campaign_data)}</p>
    <p>Total Replies: {len(replies_data)}</p>

    <script>
        const campaignData = {campaign_json};
        const repliesData = {replies_json};

        console.log('Datos cargados:', campaignData.length, 'campañas', repliesData.length, 'replies');
    </script>
</body>
</html>"""

    # Guardar
    with open('dashboard_teamficient.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"✓ Dashboard generado: dashboard_teamficient.html")
    print(f"  - {len(campaign_data)} campañas")
    print(f"  - {len(replies_data)} replies")
    print(f"  - Total Leads: {sum(c.get('Leads Generated', 0) for c in campaign_data)}")

if __name__ == '__main__':
    generate_improved_dashboard()
