import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1.base_query import FieldFilter
import bcrypt
import os

def clean_private_key(pk):
    if not isinstance(pk, str): return pk
    pk = pk.strip()
    
    # ตรวจสอบเบื้องต้นว่าก๊อปปี้มาผิดหรือไม่
    if pk.startswith('{'):
        st.error("⚠️ คำเตือน: ข้อมูลในช่อง private_key ขึ้นต้นด้วย '{' เป็นไปได้ว่าคุณก๊อปปี้ไฟล์ JSON มาวางทั้งไฟล์ แทนที่จะเป็นแค่รหัสกุญแจครับ")
    
    # ถ้ามีการก๊อปปี้เครื่องหมายคำพูดติดมาด้วย ให้เอาออก
    if (pk.startswith('"') and pk.endswith('"')) or (pk.startswith("'") and pk.endswith("'")):
        pk = pk[1:-1]
    
    # จัดการเรื่องขึ้นบรรทัดใหม่ (รองรับทั้งแบบ \n จริงๆ และแบบพิมพ์ตัวอักษร \ n)
    return pk.replace("\\n", "\n").replace("\\\\n", "\n")

def init_firebase():
    # ลบการเชื่อมต่อเดิมทิ้งเพื่อบังคับโหลดค่าใหม่จาก Secrets (ป้องกันการจำค่าเก่าที่ผิด)
    for app_name in list(firebase_admin._apps.keys()):
        firebase_admin.delete_app(firebase_admin._apps[app_name])
        
    try:
        # 1. ลองโหลดจากไฟล์ credentials.json ก่อน
        if os.path.exists("credentials.json"):
            cred = credentials.Certificate("credentials.json")
        # 2. ถ้าไม่มีไฟล์ ให้โหลดจาก st.secrets
        elif "gspread_credentials" in st.secrets:
            source = st.secrets["gspread_credentials"]
            cert_dict = dict(source)
            cert_dict["private_key"] = clean_private_key(cert_dict.get("private_key", ""))
            cred = credentials.Certificate(cert_dict)
        else:
            # ลองโหลดจาก secrets โดยตรง (กรณีไม่ได้ครอบด้วย gspread_credentials)
            needed_keys = ["type", "project_id", "private_key_id", "private_key", "client_email", "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url", "client_x509_cert_url"]
            cert_dict = {k: st.secrets[k] for k in needed_keys if k in st.secrets}
            
            if "private_key" in cert_dict:
                cert_dict["private_key"] = clean_private_key(cert_dict["private_key"])
                cred = credentials.Certificate(cert_dict)
            else:
                st.error("❌ ไม่พบข้อมูลการเชื่อมต่อ Firebase ใน Secrets")
                return

        firebase_admin.initialize_app(cred)
        # แสดงข้อมูลเพื่อตรวจสอบแบบชัดเจนที่ด้านบนหน้าจอ
        st.warning(f"🔧 Firebase Debug: Project={firebase_admin.get_app().project_id} | Email={firebase_admin.get_app().credential.service_account_email[:30]}...")
    except Exception as e:
        st.error(f"❌ โหลด Firebase ไม่สำเร็จ: {e}")

def get_db():
    try:
        return firestore.client()
    except Exception as e:
        st.error(f"Error connecting to Firestore: {e}")
        return None

def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password, hashed):
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def auth_css():
    st.markdown("""
    <style>
        [data-testid="stAppViewContainer"] {
            background-color: #1a4980;
        }
        [data-testid="stHeader"] {
            background-color: rgba(0,0,0,0);
        }
        [data-testid="stForm"] {
            background-color: white !important;
            border-radius: 10px !important;
            padding: 20px !important;
            box-shadow: 0 4px 10px rgba(0,0,0,0.2) !important;
            border: none !important;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: pre-wrap;
            background-color: #f0f2f6;
            border-radius: 5px 5px 0px 0px;
            color: #31333F;
        }
        .stTabs [aria-selected="true"] {
            background-color: white;
            color: #1a4980;
            border-top: 3px solid #1a4980;
        }
    </style>
    """, unsafe_allow_html=True)

def admin_css():
    st.markdown("""
    <style>
        [data-testid="stAppViewContainer"] {
            background-color: #2d1b4e;
        }
        [data-testid="stHeader"] {
            background-color: rgba(0,0,0,0);
        }
        [data-testid="stVerticalBlockBorderWrapper"] {
            background-color: rgba(255,255,255,0.05) !important;
            border-color: rgba(255,255,255,0.2) !important;
        }
    </style>
    """, unsafe_allow_html=True)

def check_login():
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
    
    if st.session_state['logged_in']:
        # If logged in as admin, show admin dashboard instead of main app?
        if st.session_state.get('role') == 'admin':
            admin_dashboard()
            return False # Return False so main app doesn't load
        return True

    init_firebase()
    db = get_db()
    
    auth_css()
    
    # UI Structure
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style="text-align: center; margin-bottom: 20px;">
            <div style="background-color: white; width: 80px; height: 80px; border-radius: 50%; display: inline-flex; justify-content: center; align-items: center; margin-bottom: 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.2);">
                <span style="font-size: 40px; color: #1a4980;">🏛️</span>
            </div>
            <h3 style="color: white; font-weight: bold; margin-bottom: 5px;">ระบบทะเบียนกิจการที่เป็น<br>อันตรายต่อสุขภาพ</h3>
            <p style="color: #e2e8f0; font-size: 14px;">องค์การบริหารส่วนตำบลดอยงาม อำเภอพาน จังหวัดเชียงราย</p>
        </div>
        """, unsafe_allow_html=True)
        
        tab1, tab2, tab3 = st.tabs(["เข้าสู่ระบบ", "ลงทะเบียนใหม่", "ผู้ดูแลระบบ"])
        
        with tab1:
            with st.form("login_form"):
                st.markdown("<p style='color: #1a4980; font-weight: bold; margin-bottom: 5px;'>ชื่อผู้ใช้</p>", unsafe_allow_html=True)
                username = st.text_input("ชื่อผู้ใช้", label_visibility="collapsed", placeholder="ชื่อผู้ใช้งาน")
                st.markdown("<p style='color: #1a4980; font-weight: bold; margin-bottom: 5px;'>รหัสผ่าน</p>", unsafe_allow_html=True)
                password = st.text_input("รหัสผ่าน", type="password", label_visibility="collapsed", placeholder="รหัสผ่าน")
                
                submit = st.form_submit_button("เข้าสู่ระบบ", use_container_width=True, type="primary")
                
                if submit:
                    if not username or not password:
                        st.error("กรุณากรอกชื่อผู้ใช้และรหัสผ่าน")
                    else:
                        doc_ref = db.collection("users").document(username)
                        doc = doc_ref.get()
                        
                        if doc.exists:
                            user_data = doc.to_dict()
                            if check_password(password, user_data.get('password', '')):
                                if user_data.get('status') == 'approved':
                                    st.session_state['logged_in'] = True
                                    st.session_state['username'] = username
                                    st.session_state['name'] = user_data.get('name')
                                    st.session_state['role'] = 'staff'
                                    st.success("เข้าสู่ระบบสำเร็จ!")
                                    st.rerun()
                                elif user_data.get('status') == 'pending':
                                    st.warning("บัญชีของคุณอยู่ระหว่างรอการอนุมัติจากผู้ดูแลระบบ")
                                else:
                                    st.error("บัญชีของคุณถูกระงับการใช้งาน")
                            else:
                                st.error("รหัสผ่านไม่ถูกต้อง")
                        else:
                            st.error("ไม่พบชื่อผู้ใช้งานนี้")
                            
        with tab2:
            with st.form("register_form"):
                new_username = st.text_input("ชื่อผู้ใช้ (Username)")
                new_password = st.text_input("รหัสผ่าน (Password)", type="password")
                new_name = st.text_input("ชื่อ-นามสกุล (Full Name)")
                new_position = st.text_input("ตำแหน่ง (Position)", placeholder="เช่น นักวิชาการสาธารณสุข")
                
                reg_submit = st.form_submit_button("ลงทะเบียน", use_container_width=True)
                
                if reg_submit:
                    if not new_username or not new_password or not new_name:
                        st.error("กรุณากรอกข้อมูลให้ครบถ้วน")
                    else:
                        doc_ref = db.collection("users").document(new_username)
                        if doc_ref.get().exists:
                            st.error("ชื่อผู้ใช้นี้มีคนใช้แล้ว กรุณาเลือกชื่ออื่น")
                        else:
                            try:
                                doc_ref.set({
                                    "username": new_username,
                                    "password": hash_password(new_password),
                                    "name": new_name,
                                    "position": new_position,
                                    "role": "staff",
                                    "status": "pending"
                                })
                                st.success("ลงทะเบียนสำเร็จ! กรุณารอผู้ดูแลระบบอนุมัติบัญชีของคุณ")
                            except Exception as e:
                                st.error(f"เกิดข้อผิดพลาด: {e}")
                                
        with tab3:
            with st.form("admin_login_form"):
                st.info("สำหรับผู้ดูแลระบบเท่านั้น (Admin)")
                st.markdown("<p style='color: #1a4980; font-weight: bold; margin-bottom: 5px;'>รหัสผ่านแอดมิน</p>", unsafe_allow_html=True)
                admin_password = st.text_input("รหัสผ่านแอดมิน", type="password", label_visibility="collapsed", placeholder="รหัสผ่านผู้ดูแลระบบ")
                
                admin_submit = st.form_submit_button("เข้าสู่ระบบผู้ดูแลระบบ", use_container_width=True, type="primary")
                
                if admin_submit:
                    # Check password from secrets
                    if 'admin_password' in st.secrets:
                        correct_pw = st.secrets['admin_password']
                    else:
                        correct_pw = 'admin1234' # Fallback default
                        
                    if admin_password == correct_pw:
                        st.session_state['logged_in'] = True
                        st.session_state['role'] = 'admin'
                        st.success("เข้าสู่ระบบแอดมินสำเร็จ!")
                        st.rerun()
                    else:
                        st.error("รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง")
                        
        st.markdown('<br>', unsafe_allow_html=True)
    return False

def admin_dashboard():
    admin_css()
    db = get_db()
    
    col1, col2 = st.columns([4, 1])
    with col1:
        st.markdown("<h2 style='color: white;'>🔐 แผงควบคุมผู้ดูแลระบบ</h2>", unsafe_allow_html=True)
    with col2:
        if st.button("ออกจากระบบ", use_container_width=True):
            st.session_state.clear()
            st.rerun()
            
    st.markdown("<hr style='border-color: rgba(255,255,255,0.2);'>", unsafe_allow_html=True)
    
    st.markdown("<h3 style='color: white;'>⏳ คำขอลงทะเบียนรอการอนุมัติ</h3>", unsafe_allow_html=True)
    
    pending_users = list(db.collection("users").where(filter=FieldFilter("status", "==", "pending")).stream())
    
    if not pending_users:
        st.markdown("<p style='text-align: center; color: #a78bfa !important;'>ไม่มีคำขอรอการอนุมัติ</p>", unsafe_allow_html=True)
    else:
        for u in pending_users:
            u_data = u.to_dict()
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.markdown(f"<p style='color: white;'><b>{u_data.get('name')}</b> ({u.id}) — {u_data.get('position', '-')}</p>", unsafe_allow_html=True)
                with c2:
                    c2_1, c2_2 = st.columns(2)
                    if c2_1.button("✅ อนุมัติ", key=f"app_{u.id}", type="primary"):
                        db.collection("users").document(u.id).update({"status": "approved"})
                        st.rerun()
                    if c2_2.button("❌ ปฏิเสธ", key=f"rej_{u.id}"):
                        db.collection("users").document(u.id).delete()
                        st.rerun()
                        
    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("<h3 style='color: white;'>✅ เจ้าหน้าที่ที่ได้รับการอนุมัติแล้ว</h3>", unsafe_allow_html=True)
    
    approved_users = list(db.collection("users").where(filter=FieldFilter("status", "==", "approved")).stream())
    
    if not approved_users:
        st.markdown("<p style='text-align: center; color: #a78bfa !important;'>ไม่มีเจ้าหน้าที่ในระบบ</p>", unsafe_allow_html=True)
    else:
        for u in approved_users:
            u_data = u.to_dict()
            with st.container(border=True):
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.markdown(f"<p style='color: white;'><span style='color: #34d399;'>●</span> <b>{u_data.get('name')}</b> ({u.id}) — {u_data.get('position', '-')}</p>", unsafe_allow_html=True)
                with c2:
                    if st.button("ยกเลิกสิทธิ์", key=f"rev_{u.id}"):
                        db.collection("users").document(u.id).update({"status": "revoked"})
                        st.rerun()
