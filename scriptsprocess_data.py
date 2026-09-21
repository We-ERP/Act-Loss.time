import os
import pandas as pd
import numpy as np

def parse_time_to_seconds(val):
    if pd.isna(val) or val is None or str(val).strip() in ['', 'Unpaid', 'Maternity', 'Planned sick']:
        return 0
    val_str = str(val).strip()
    try:
        parts = val_str.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        pass
    return 0

def seconds_to_hhmmss(seconds):
    if pd.isna(seconds) or seconds <= 0:
        return "0:00:00"
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return f"{hours}:{minutes:02d}:{secs:02d}"

def process_pipeline():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    template_path = os.path.join(base_dir, 'templates', 'Structure.xlsx')
    output_path = os.path.join(base_dir, 'output', 'Final_Report.xlsx')

    print("🔄 جاري قراءة البيانات...")
    
    utl_path = os.path.join(data_dir, 'UTL.xlsx')
    ir_path = os.path.join(data_dir, 'IR.xlsx')
    comp_path = os.path.join(data_dir, 'Compensation.xlsx')

    utl_df = pd.read_excel(utl_path) if os.path.exists(utl_path) else pd.DataFrame()
    ir_df = pd.read_excel(ir_path) if os.path.exists(ir_path) else pd.DataFrame()
    comp_df = pd.read_excel(comp_path) if os.path.exists(comp_path) else pd.DataFrame()

    # معالجة IR
    ir_assigned = ir_df.groupby('assigned_to')['id'].count().to_dict() if 'assigned_to' in ir_df.columns else {}
    ir_closed = ir_df[ir_df['ticket_status'] == 'Reached'].groupby('assigned_to')['id'].count().to_dict() if 'ticket_status' in ir_df.columns and 'assigned_to' in ir_df.columns else {}

    # معالجة Comp
    comp_dict = {}
    if 'ID' in comp_df.columns and 'Code Time' in comp_df.columns:
        comp_df['ID'] = comp_df['ID'].astype(str).str.strip()
        comp_dict = comp_df.set_index('ID')['Code Time'].to_dict()

    # معالجة UTL
    if 'Login ID' in utl_df.columns:
        utl_df['Login ID'] = utl_df['Login ID'].astype(str).str.strip()

    struct_df = pd.read_excel(template_path)

    for idx, row in struct_df.iterrows():
        login_id = str(row['Login ID']).strip() if pd.notna(row['Login ID']) else ""
        tts_user = str(row['TTS User']).strip() if pd.notna(row['TTS User']) else ""
        status = row.get('Status', 'Active')

        if status != 'Active':
            struct_df.at[idx, 'Loss Time'] = status
            continue

        if not utl_df.empty and 'Talk Time' in utl_df.columns:
            agent_utl = utl_df[utl_df['Login ID'] == login_id]
            if not agent_utl.empty:
                talk_sec = parse_time_to_seconds(agent_utl['Talk Time'].values[0])
                struct_df.at[idx, 'Talk Time'] = seconds_to_hhmmss(talk_sec)

        struct_df.at[idx, 'Assigning Tkts'] = ir_assigned.get(tts_user, row.get('Assigning Tkts', 0))
        struct_df.at[idx, 'TKT'] = ir_closed.get(tts_user, row.get('TKT', 0))

        comp_val = comp_dict.get(login_id, row.get('Comp', "0:00:00"))
        struct_df.at[idx, 'Comp'] = comp_val

        tele_sch_sec = parse_time_to_seconds(row.get('Tele-SCH', '7:12:00'))
        talk_time_sec = parse_time_to_seconds(struct_df.at[idx, 'Talk Time'])
        comp_sec = parse_time_to_seconds(comp_val)

        loss_time_sec = max(0, tele_sch_sec - (talk_time_sec + comp_sec))
        struct_df.at[idx, 'Loss Time'] = seconds_to_hhmmss(loss_time_sec)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    struct_df.to_excel(output_path, index=False)
    print(f"✅ تم حفظ التقرير المحدث في: {output_path}")

if __name__ == '__main__':
    process_pipeline()