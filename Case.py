import streamlit as st
import pandas as pd
import datetime
import gspread
import json
from google.oauth2.service_account import Credentials

# ------------ CONFIG -------------
SHEET_NAME = "Project_List"
WORKSHEET_NAME = "Sheet1"
LOCK_SHEET_NAME = "Lock"

USERS = {
    "sam@kipotec.com.tw":   {"password": "Kipo-0926969586$$$",    "role": "requestor", "name": "Sam",    "priority": 4},
    "sale1@kipotec.com.tw": {"password": "Kipo-0917369466$$$",    "role": "requestor", "name": "Vivian", "priority": 3},
    "sale2@kipotec.com.tw": {"password": "Kipo-0905038111$$$",    "role": "requestor", "name": "Lillian","priority": 2},
    "sale5@kipotec.com.tw": {"password": "Kipo-0925698417$$$",    "role": "requestor", "name": "Wendy",  "priority": 1},
    "bruce@kipotec.com.tw": {"password": "Kipo-0935300679$$$@@@", "role": "approver",  "name": "Bruce",  "priority": 0},
}

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# ---------- Google Sheet ----------
def get_gc():
    service_account_info = json.loads(st.secrets["GOOGLE_CLOUD_KEY"])
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    return gspread.authorize(creds)

def ensure_sheet_and_headers(ws, expected_headers):
    values = ws.get_all_values()
    if not values:
        ws.append_row(expected_headers)
        return
    headers = values[0]
    if headers != expected_headers:
        missing = [h for h in expected_headers if h not in headers]
        if missing:
            ws.update('A1', [headers + missing])

def open_main_ws():
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=20)
    ensure_sheet_and_headers(ws, ["Client","Project","Cooling","Department","Number","Project_ID",
                                  "Created_Time","Status","Note","Applicant","Approver"])
    return ws

def open_lock_ws():
    gc = get_gc()
    sh = gc.open(SHEET_NAME)
    try:
        ws_lock = sh.worksheet(LOCK_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws_lock = sh.add_worksheet(title=LOCK_SHEET_NAME, rows=10, cols=2)
    ensure_sheet_and_headers(ws_lock, ["User","Locked_Time"])
    return ws_lock

def load_sheet_df():
    ws = open_main_ws()
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    for col in ["Client","Project","Cooling","Department","Number","Project_ID",
                "Created_Time","Status","Note","Applicant","Approver"]:
        if col not in df.columns:
            df[col] = ""
    return df, ws

def load_lock_df():
    ws_lock = open_lock_ws()
    data = ws_lock.get_all_records()
    df_lock = pd.DataFrame(data)
    if "User" not in df_lock.columns: df_lock["User"] = ""
    if "Locked_Time" not in df_lock.columns: df_lock["Locked_Time"] = ""
    return df_lock, ws_lock

def append_row_by_headers(ws, row_dict):
    headers = ws.row_values(1)
    row = [str(row_dict.get(h, "")) for h in headers]
    ws.append_row(row, value_input_option="RAW")

# ---------- Lock 機制 ----------
def acquire_lock(username: str) -> (bool, str):
    df_lock, ws_lock = load_lock_df()
    active = df_lock[(df_lock["User"] != "")]
    if active.empty:
        ws_lock.append_row([username, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        return True, ""
    current_user = active.iloc[0]["User"]
    if current_user == username:
        return True, ""
    return False, USERS.get(current_user, {"name": current_user})["name"]

def release_lock(username: str):
    df_lock, ws_lock = load_lock_df()
    for i, u in enumerate(df_lock.get("User", [])):
        if u == username:
            ws_lock.update_cell(i + 2, 1, "")
            ws_lock.update_cell(i + 2, 2, "")

# ---------- 號碼 ----------
def next_number_for_client(df: pd.DataFrame, client_code: str) -> str:
    if df.empty:
        return "001"
    client_mask = df["Client"].astype(str).str.extract(r"\((.*?)\)")[0] == client_code
    nums = pd.to_numeric(df[client_mask]["Number"], errors="coerce").dropna().astype(int)
    if nums.empty:
        return "001"
    return str(nums.max() + 1).zfill(3)

def find_reserved_row(df: pd.DataFrame, username: str):
    display_name = USERS[username]["name"]
    mask = (df["Applicant"] == display_name) & (df["Status"] == "預留中")
    if mask.any():
        idx = df.index[mask][0]
        return idx, df.loc[idx]
    return None, None

def find_row_by_project_id(df: pd.DataFrame, project_id: str):
    mask = (df["Project_ID"] == project_id)
    if mask.any():
        idx = df.index[mask][0]
        return idx, df.loc[idx]
    return None, None

# ---------- UI ----------
st.set_page_config(page_title="Project ID System", layout="centered")
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = None

def login():
    if not st.session_state.logged_in:
        st.title("專案建立系統")
        username = st.text_input("帳號 (email)")
        password = st.text_input("密碼", type="password")
        if st.button("登入"):
            if username in USERS and USERS[username]["password"] == password:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = USERS[username]["role"]
                st.rerun()
            else:
                st.error("帳號或密碼錯誤")

def main_page():
    username = st.session_state.username
    role = st.session_state.role
    display_name = USERS[username]["name"]
    st.write(f"已登入：{display_name}（{role}）")

    if st.button("登出"):
        release_lock(username)
        st.session_state.logged_in = False
        st.rerun()

    clients = {"01":"仁寶","02":"廣達","03":"緯創","04":"華勤","05":"光寶","06":"技嘉","07":"智邦","00":"其他"}
    project_types = {"S1":"Server","N1":"NB","M1":"MINI PC","A1":"AIO","C1":"車用","00":"其他"}
    coolings = {"A":"氣冷","L":"水冷"}
    departments = {"F":"風扇部門","N":"筆電模組部門","S":"伺服器模組部門"}

    if role == "requestor":
        df, ws = load_sheet_df()
        reserved_idx, reserved_row = find_reserved_row(df, username)

        if reserved_idx is not None:
            st.subheader("上次未完成的專案，可修改後送出")
            client_choice = st.selectbox("客戶端", options=list(clients.keys()),
                                         index=list(clients.keys()).index(reserved_row["Client"][1:3]),
                                         format_func=lambda k: f"({k}){clients[k]}")
            project_choice = st.selectbox("專案類型", options=list(project_types.keys()),
                                          index=list(project_types.keys()).index(reserved_row["Project"][1:3]),
                                          format_func=lambda k: f"({k}){project_types[k]}")
            cooling_choice = st.selectbox("散熱方案", options=list(coolings.keys()),
                                          index=list(coolings.keys()).index(reserved_row["Cooling"][1:2]),
                                          format_func=lambda k: f"({k}){coolings[k]}")
            dept_choice = st.selectbox("部門代碼", options=list(departments.keys()),
                                       index=list(departments.keys()).index(reserved_row["Department"][1:2]),
                                       format_func=lambda k: f"({k}){departments[k]}")
            note = st.text_area("備註", value=reserved_row["Note"])

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("送出"):
                    row_num = reserved_idx + 2
                    ws.update_cell(row_num, df.columns.get_loc("Client")+1, f"({client_choice}){clients[client_choice]}")
                    ws.update_cell(row_num, df.columns.get_loc("Project")+1, f"({project_choice}){project_types[project_choice]}")
                    ws.update_cell(row_num, df.columns.get_loc("Cooling")+1, f"({cooling_choice}){coolings[cooling_choice]}")
                    ws.update_cell(row_num, df.columns.get_loc("Department")+1, f"({dept_choice}){departments[dept_choice]}")
                    ws.update_cell(row_num, df.columns.get_loc("Note")+1, note)
                    ws.update_cell(row_num, df.columns.get_loc("Status")+1, "簽核中")
                    ws.update_cell(row_num, df.columns.get_loc("Created_Time")+1,
                                   datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    release_lock(username)
                    st.success("專案已送出")
                    st.rerun()
            with col_b:
                if st.button("取消"):
                    ws.update_cell(reserved_idx + 2, df.columns.get_loc("Status")+1, "已取消")
                    release_lock(username)
                    st.info("已取消專案")
                    st.rerun()
            return

        st.subheader("建立專案編號")
        client_choice = st.selectbox("客戶端", options=list(clients.keys()), format_func=lambda k: f"({k}){clients[k]}")
        project_choice = st.selectbox("專案類型", options=list(project_types.keys()), format_func=lambda k: f"({k}){project_types[k]}")
        cooling_choice = st.selectbox("散熱方案", options=list(coolings.keys()), format_func=lambda k: f"({k}){coolings[k]}")
        dept_choice = st.selectbox("部門代碼", options=list(departments.keys()), format_func=lambda k: f"({k}){departments[k]}")
        note = st.text_area("備註")

        if st.button("生成"):
            lock_acquired, holder_name = acquire_lock(username)
            if not lock_acquired:
                st.warning(f"目前由 {holder_name} 佔用，請稍後")
                return
            df_latest, ws_latest = load_sheet_df()
            number_str = next_number_for_client(df_latest, client_choice)
            project_id = f"{client_choice}-{project_choice}-{cooling_choice}-{dept_choice}-{number_str}"
            row_dict = {
                "Client": f"({client_choice}){clients[client_choice]}",
                "Project": f"({project_choice}){project_types[project_choice]}",
                "Cooling": f"({cooling_choice}){coolings[cooling_choice]}",
                "Department": f"({dept_choice}){departments[dept_choice]}",
                "Number": number_str,
                "Project_ID": project_id,
                "Created_Time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Status": "預留中",
                "Note": note,
                "Applicant": display_name,
                "Approver": ""
            }
            append_row_by_headers(ws_latest, row_dict)
            st.success(f"已生成：{project_id}")
            st.rerun()

    elif role == "approver":
        st.subheader("待審核專案")
        df, ws = load_sheet_df()
        df_pending = df[df["Status"] == "簽核中"]
        if df_pending.empty:
            st.info("無待審核專案")
        else:
            st.dataframe(df_pending)
            selected = st.multiselect("選擇專案", df_pending["Project_ID"].tolist())
            if st.button("批准"):
                for pid in selected:
                    idx, _ = find_row_by_project_id(df, pid)
                    if idx is not None:
                        ws.update_cell(idx+2, df.columns.get_loc("Status")+1, "批准")
                        ws.update_cell(idx+2, df.columns.get_loc("Approver")+1, display_name)
                st.success(f"已批准 {len(selected)} 個專案")
                st.rerun()
            if st.button("駁回"):
                for pid in selected:
                    idx, _ = find_row_by_project_id(df, pid)
                    if idx is not None:
                        # 改成 "預留中" 讓使用者能再次編輯送出
                        ws.update_cell(idx+2, df.columns.get_loc("Status")+1, "預留中")
                        ws.update_cell(idx+2, df.columns.get_loc("Approver")+1, display_name)
                st.warning(f"已駁回 {len(selected)} 個專案，狀態已改回預留中")
                st.rerun()

def main():
    if not st.session_state.logged_in:
        login()
    else:
        main_page()
