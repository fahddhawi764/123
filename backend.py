import os
import sqlite3
from datetime import datetime, timedelta
import pandas as pd

# إعداد المسارات
script_dir = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(script_dir, 'document_management.db')
ATTACHMENTS_DIR = os.path.join(script_dir, 'attachments')

if not os.path.exists(ATTACHMENTS_DIR):
    os.makedirs(ATTACHMENTS_DIR)

# --- دوال تحويل التاريخ ---
def convert_date_to_db_format(date_str_ddmmyyyy):
    if not date_str_ddmmyyyy:
        return None
    try:
        return datetime.strptime(date_str_ddmmyyyy, "%d-%m-%Y").strftime("%Y-%m-%d")
    except ValueError:
        raise ValueError("تنسيق تاريخ غير صحيح. يرجى استخدام DD-MM-YYYY.")

def convert_date_from_db_format(date_str_yyyymmdd):
    if not date_str_yyyymmdd:
        return ""
    try:
        return datetime.strptime(date_str_yyyymmdd, "%Y-%m-%d").strftime("%d-%m-%Y")
    except ValueError:
        return ""

# --- إنشاء قاعدة البيانات ---
def create_database():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                number TEXT NOT NULL UNIQUE,
                date TEXT NOT NULL,
                expiry_date TEXT,
                issuer TEXT NOT NULL,
                employee_id INTEGER,
                category TEXT,
                tags TEXT,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE SET NULL
            )
        ''')
        cursor.execute("PRAGMA table_info(documents)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'category' not in columns:
            cursor.execute("ALTER TABLE documents ADD COLUMN category TEXT")
        if 'tags' not in columns:
            cursor.execute("ALTER TABLE documents ADD COLUMN tags TEXT")


        cursor.execute('''
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                employee_number TEXT NOT NULL UNIQUE,
                department TEXT,
                contact_info TEXT,
                hire_date TEXT
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS audit_log (
                log_id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                user_action TEXT NOT NULL,
                details TEXT
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                filename TEXT NOT NULL,
                filepath TEXT NOT NULL,
                upload_date TEXT NOT NULL,
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE
            )
        ''')

        # جدول الرواتب الجديد
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS salaries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL,
                basic_salary REAL NOT NULL, -- هذا هو الراتب الشهري
                allowances REAL NOT NULL,
                deductions REAL NOT NULL,
                net_salary REAL NOT NULL,
                payment_method TEXT NOT NULL,
                payment_date TEXT NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
            )
        ''')

        conn.commit()

# --- سجل التدقيق ---
def log_audit_event(action, details):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT INTO audit_log (timestamp, user_action, details) VALUES (?, ?, ?)",
                       (timestamp, action, details))
        conn.commit()

# --- CRUD: المستندات ---
def add_document(name, number, date_ddmmyyyy, expiry_date_ddmmyyyy, issuer, employee_id, category, tags):
    if not name or not number or not date_ddmmyyyy or not issuer:
        raise ValueError("يرجى تعبئة جميع الحقول المطلوبة: الاسم، الرقم، تاريخ الإصدار، الجهة المصدرة.")

    try:
        date_db = convert_date_to_db_format(date_ddmmyyyy)
        expiry_date_db = convert_date_to_db_format(expiry_date_ddmmyyyy) if expiry_date_ddmmyyyy else None

        if date_db is None:
            raise ValueError("تنسيق تاريخ الإصدار غير صحيح. يرجى استخدام DD-MM-YYYY.")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO documents (name, number, date, expiry_date, issuer, employee_id, category, tags) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                           (name, number, date_db, expiry_date_db, issuer, employee_id, category, tags))
            doc_id = cursor.lastrowid
            conn.commit()
            log_audit_event("إضافة مستند", f"تمت إضافة المستند: {name} ({number})")
            return doc_id
    except sqlite3.IntegrityError:
        raise ValueError("⚠ رقم المستند موجود مسبقاً. يرجى إدخال رقم فريد.")
    except ValueError as e:
        raise e
    except Exception as e:
        raise Exception(f"❌ حدث خطأ غير متوقع أثناء إضافة المستند: {str(e)}")

def update_document(doc_id, name, number, date_ddmmyyyy, expiry_date_ddmmyyyy, issuer, employee_id, category, tags):
    if not name or not number or not date_ddmmyyyy or not issuer:
        raise ValueError("يرجى تعبئة جميع الحقول المطلوبة: الاسم، الرقم، تاريخ الإصدار، الجهة المصدرة.")

    try:
        date_db = convert_date_to_db_format(date_ddmmyyyy)
        expiry_date_db = convert_date_to_db_format(expiry_date_ddmmyyyy) if expiry_date_ddmmyyyy else None

        if date_db is None:
            raise ValueError("تنسيق تاريخ الإصدار غير صحيح. يرجى استخدام DD-MM-YYYY.")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE documents SET name=?, number=?, date=?, expiry_date=?, issuer=?, employee_id=?, category=?, tags=? WHERE id=?",
                           (name, number, date_db, expiry_date_db, issuer, employee_id, category, tags, doc_id))
            conn.commit()
            if cursor.rowcount == 0:
                raise ValueError(f"لم يتم العثور على مستند بالرقم التعريفي {doc_id} للتعديل.")
            log_audit_event("تحديث مستند", f"تم تحديث المستند ID: {doc_id} إلى: {name} ({number})")
    except sqlite3.IntegrityError:
        raise ValueError("⚠ رقم المستند موجود مسبقاً لمستند آخر. يرجى إدخال رقم فريد.")
    except ValueError as e:
        raise e
    except Exception as e:
        raise Exception(f"❌ حدث خطأ غير متوقع أثناء تحديث المستند: {str(e)}")

def delete_document(doc_id):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name, number FROM documents WHERE id=?", (doc_id,))
        doc_info = cursor.fetchone()
        if doc_info:
            cursor.execute("SELECT filepath FROM attachments WHERE document_id = ?", (doc_id,))
            attachments = cursor.fetchall()
            for att_path in attachments:
                if os.path.exists(att_path[0]):
                    try:
                        os.remove(att_path[0])
                        log_audit_event("حذف مرفق", f"تم حذف الملف المرفق من القرص: {att_path[0]}")
                    except OSError as e:
                        print(f"خطأ في حذف الملف {att_path[0]}: {e}")
            
            cursor.execute("DELETE FROM documents WHERE id=?", (doc_id,))
            conn.commit()
            log_audit_event("حذف مستند", f"تم حذف المستند: {doc_info[0]} ({doc_info[1]}) ID: {doc_id}")
        else:
            raise ValueError(f"لم يتم العثور على مستند بالرقم التعريفي {doc_id} للحذف.")

# --- CRUD: الموظفين ---
def add_employee(name, employee_number, department, contact_info, hire_date_ddmmyyyy):
    if not name or not employee_number or not hire_date_ddmmyyyy:
        raise ValueError("يرجى تعبئة جميع الحقول المطلوبة للموظف: الاسم، رقم الموظف، تاريخ التعيين.")
    try:
        hire_date_db = convert_date_to_db_format(hire_date_ddmmyyyy)
        if hire_date_db is None:
            raise ValueError("تنسيق تاريخ التعيين غير صحيح. يرجى استخدام DD-MM-YYYY.")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO employees (name, employee_number, department, contact_info, hire_date) VALUES (?, ?, ?, ?, ?)",
                           (name, employee_number, department, contact_info, hire_date_db))
            conn.commit()
            log_audit_event("إضافة موظف", f"تمت إضافة الموظف: {name} ({employee_number})")
    except sqlite3.IntegrityError:
        raise ValueError("رقم الموظف موجود مسبقاً. يرجى إدخال رقم فريد.")
    except ValueError as e:
        raise e
    except Exception as e:
        raise Exception(f"❌ حدث خطأ غير متوقع أثناء إضافة الموظف: {str(e)}")

def update_employee(emp_id, name, employee_number, department, contact_info, hire_date_ddmmyyyy):
    if not name or not employee_number or not hire_date_ddmmyyyy:
        raise ValueError("يرجى تعبئة جميع الحقول المطلوبة للموظف: الاسم، رقم الموظف، تاريخ التعيين.")
    try:
        hire_date_db = convert_date_to_db_format(hire_date_ddmmyyyy)
        if hire_date_db is None:
            raise ValueError("تنسيق تاريخ التعيين غير صحيح. يرجى استخدام DD-MM-YYYY.")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE employees SET name=?, employee_number=?, department=?, contact_info=?, hire_date=? WHERE id=?",
                           (name, employee_number, department, contact_info, hire_date_db, emp_id))
            conn.commit()
            if cursor.rowcount == 0:
                raise ValueError(f"لم يتم العثور على موظف بالرقم التعريفي {emp_id} للتعديل.")
            log_audit_event("تحديث موظف", f"تم تحديث الموظف ID: {emp_id} إلى: {name} ({employee_number})")
    except sqlite3.IntegrityError:
        raise ValueError("رقم الموظف موجود مسبقاً لموظف آخر. يرجى إدخال رقم فريد.")
    except ValueError as e:
        raise e
    except Exception as e:
        raise Exception(f"❌ حدث خطأ غير متوقع أثناء تحديث بيانات الموظف: {str(e)}")

def delete_employee(emp_id):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name, employee_number FROM employees WHERE id=?", (emp_id,))
        emp_info = cursor.fetchone()
        if emp_info:
            cursor.execute("UPDATE documents SET employee_id = NULL WHERE employee_id = ?", (emp_id,))
            cursor.execute("DELETE FROM employees WHERE id=?", (emp_id,))
            conn.commit()
            log_audit_event("حذف موظف", f"تم حذف الموظف: {emp_info[0]} ({emp_info[1]}) ID: {emp_id}")
        else:
            raise ValueError(f"لم يتم العثور على موظف بالرقم التعريفي {emp_id} للحذف.")

# --- دوال إدارة المرفقات ---
def add_attachment(document_id, original_filepath):
    filename = os.path.basename(original_filepath)
    unique_filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}"
    destination_filepath = os.path.join(ATTACHMENTS_DIR, unique_filename)

    try:
        import shutil
        shutil.copy(original_filepath, destination_filepath)
        
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            upload_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute("INSERT INTO attachments (document_id, filename, filepath, upload_date) VALUES (?, ?, ?, ?)",
                           (document_id, filename, destination_filepath, upload_date))
            conn.commit()
            log_audit_event("إضافة مرفق", f"تم إرفاق الملف {filename} للمستند ID: {document_id}")
            return destination_filepath
    except Exception as e:
        raise Exception(f"❌ حدث خطأ أثناء إرفاق الملف: {str(e)}")

def get_attachments_for_document(document_id):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, filename, filepath, upload_date FROM attachments WHERE document_id = ?", (document_id,))
        return cursor.fetchall()

def delete_attachment(attachment_id, filepath):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))
        conn.commit()
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                log_audit_event("حذف مرفق", f"تم حذف الملف المرفق من القرص: {filepath}")
            except OSError as e:
                print(f"خطأ في حذف الملف من القرص {filepath}: {e}")
        log_audit_event("حذف مرفق", f"تم حذف المرفق ID: {attachment_id} والملف: {filepath}")

# --- دوال مساعدة عامة ---
def fetch_employee_id_name():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM employees ORDER BY name")
        rows = cursor.fetchall()
    return rows

def fetch_all_employees():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, employee_number, department, contact_info, hire_date FROM employees")
        rows = cursor.fetchall()
    return rows

def fetch_audit_log():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT timestamp, user_action, details FROM audit_log ORDER BY timestamp DESC")
        rows = cursor.fetchall()
    return rows

def get_all_categories():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT category FROM documents WHERE category IS NOT NULL AND category != '' ORDER BY category")
        return [row[0] for row in cursor.fetchall()]

# New function to get all unique departments
def get_all_departments():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT department FROM employees WHERE department IS NOT NULL AND department != '' ORDER BY department")
        return [row[0] for row in cursor.fetchall()]

def calculate_remaining_time(expiry_date_str_yyyymmdd):
    if not expiry_date_str_yyyymmdd:
        return "N/A"
    try:
        expiry_date = datetime.strptime(expiry_date_str_yyyymmdd, "%Y-%m-%d").date()
        today = datetime.today().date()
        
        if expiry_date < today:
            return "منتهية الصلاحية"

        delta_days = (expiry_date - today).days

        years = delta_days // 365
        remaining_days_after_years = delta_days % 365
        
        months = remaining_days_after_years // 30
        days = remaining_days_after_years % 30

        return f"{years} سنة, {months} شهر, {days} يوم"
    except Exception as e:
        return f"خطأ في الحساب: {e}"

def fetch_all_documents_for_export():
    """يجلب جميع بيانات المستندات من قاعدة البيانات للتصدير."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, number, date, expiry_date, issuer, category, tags FROM documents")
        return cursor.fetchall()

# --- دوال الرواتب الجديدة ---
def calculate_net_salary(basic_salary, allowances, deductions):
    try:
        basic_salary = float(basic_salary)
        allowances = float(allowances)
        deductions = float(deductions)
        net = basic_salary + allowances - deductions
        return round(net, 2)
    except ValueError:
        return 0.0 # أو يمكن رفع استثناء أو إرجاع قيمة خطأ

def add_salary(employee_id, basic_salary, allowances, deductions, payment_method, payment_date_ddmmyyyy):
    if not all([employee_id, basic_salary, allowances, deductions, payment_method, payment_date_ddmmyyyy]):
        raise ValueError("يرجى تعبئة جميع حقول الراتب المطلوبة.")
    
    try:
        net_salary = calculate_net_salary(basic_salary, allowances, deductions)
        payment_date_db = convert_date_to_db_format(payment_date_ddmmyyyy)

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO salaries (employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date) VALUES (?, ?, ?, ?, ?, ?, ?)",
                           (employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date_db))
            conn.commit()
            log_audit_event("إضافة راتب", f"تمت إضافة راتب للموظف ID: {employee_id}، صافي: {net_salary}")
    except ValueError as e:
        raise e
    except Exception as e:
        raise Exception(f"❌ حدث خطأ غير متوقع أثناء إضافة الراتب: {str(e)}")

def update_salary(salary_id, employee_id, basic_salary, allowances, deductions, payment_method, payment_date_ddmmyyyy):
    if not all([salary_id, employee_id, basic_salary, allowances, deductions, payment_method, payment_date_ddmmyyyy]):
        raise ValueError("يرجى تعبئة جميع حقول الراتب المطلوبة للتعديل.")
    
    try:
        net_salary = calculate_net_salary(basic_salary, allowances, deductions)
        payment_date_db = convert_date_to_db_format(payment_date_ddmmyyyy)

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE salaries SET employee_id=?, basic_salary=?, allowances=?, deductions=?, net_salary=?, payment_method=?, payment_date=? WHERE id=?",
                           (employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date_db, salary_id))
            conn.commit()
            if cursor.rowcount == 0:
                raise ValueError(f"لم يتم العثور على راتب بالرقم التعريفي {salary_id} للتعديل.")
            log_audit_event("تحديث راتب", f"تم تحديث راتب ID: {salary_id} للموظف ID: {employee_id}، صافي: {net_salary}")
    except ValueError as e:
        raise e
    except Exception as e:
        raise Exception(f"❌ حدث خطأ غير متوقع أثناء تحديث الراتب: {str(e)}")

def delete_salary(salary_id):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT employee_id, net_salary FROM salaries WHERE id=?", (salary_id,))
        salary_info = cursor.fetchone()
        if salary_info:
            cursor.execute("DELETE FROM salaries WHERE id=?", (salary_id,))
            conn.commit()
            log_audit_event("حذف راتب", f"تم حذف راتب ID: {salary_id} للموظف ID: {salary_info[0]}، صافي: {salary_info[1]}")
        else:
            raise ValueError(f"لم يتم العثور على راتب بالرقم التعريفي {salary_id} للحذف.")

def fetch_all_salaries(department_filter=None):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        query = """
            SELECT s.id, e.name, e.department, s.basic_salary, s.allowances, s.deductions, s.net_salary, s.payment_method, s.payment_date, s.employee_id
            FROM salaries s
            JOIN employees e ON s.employee_id = e.id
        """
        params = []
        if department_filter and department_filter != "الكل":
            query += " WHERE e.department = ?"
            params.append(department_filter)
        
        query += " ORDER BY s.payment_date DESC"
        
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    return rows

def fetch_all_salaries_for_export():
    """يجلب جميع بيانات الرواتب من قاعدة البيانات للتصدير، بما في ذلك الراتب السنوي."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT s.id, e.name, e.department, s.basic_salary, s.allowances, s.deductions, s.net_salary, s.payment_method, s.payment_date
            FROM salaries s
            JOIN employees e ON s.employee_id = e.id
            ORDER BY s.payment_date DESC
        """)
        rows = cursor.fetchall()
    
    # إضافة الراتب السنوي إلى كل صف
    exported_data = []
    for row in rows:
        # row: id, employee_name, department, basic_salary (monthly), allowances, deductions, net_salary, payment_method, payment_date
        monthly_basic_salary = row[3]
        annual_basic_salary = monthly_basic_salary * 12
        
        # تحويل تاريخ الدفع للعرض في Excel
        payment_date_ddmmyyyy = convert_date_from_db_format(row[8])
        
        exported_data.append(row[0:3] + (monthly_basic_salary, annual_basic_salary) + row[4:8] + (payment_date_ddmmyyyy,))
    return exported_data

def get_last_employee_salary(employee_id):
    """
    يجلب آخر راتب أساسي وبدلات وخصومات لموظف معين.
    يعيد (basic_salary, allowances, deductions) أو None إذا لم يتم العثور على سجل.
    """
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT basic_salary, allowances, deductions
            FROM salaries
            WHERE employee_id = ?
            ORDER BY payment_date DESC
            LIMIT 1
        """, (employee_id,))
        result = cursor.fetchone()
    return result

def salary_exists_for_month(employee_id, year, month):
    """
    يتحقق مما إذا كان هناك سجل راتب لموظف معين في شهر وسنة محددين.
    """
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        # تنسيق الشهر والسنة للمقارنة مع payment_date المخزن كـYYYY-MM-DD
        # نستخدم LIKE لأنه قد يكون هناك اختلاف في اليوم
        month_str_db_format = f"{year}-{month:02d}-%"
        cursor.execute("""
            SELECT 1 FROM salaries
            WHERE employee_id = ? AND payment_date LIKE ?
            LIMIT 1
        """, (employee_id, month_str_db_format))
        return cursor.fetchone() is not None

def fetch_employee_salary_history(employee_id):
    """
    يجلب جميع سجلات الرواتب لموظف معين، مرتبة تنازليًا حسب تاريخ الدفع.
    """
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date
            FROM salaries
            WHERE employee_id = ?
            ORDER BY payment_date DESC
        """, (employee_id,))
        rows = cursor.fetchall()
    
    history_data = []
    for row in rows:
        salary_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date_db = row
        payment_date_ddmmyyyy = convert_date_from_db_format(payment_date_db)
        annual_basic_salary = basic_salary * 12
        history_data.append((salary_id, basic_salary, annual_basic_salary, allowances, deductions, net_salary, payment_method, payment_date_ddmmyyyy))
    return history_data