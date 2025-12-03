import json
import time
import win32com.client as win32
import os
import random
from datetime import datetime, timedelta

# ------------------------------
# Rutas absolutas
# ------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
HTML_DIR = os.path.join(BASE_DIR, "html")

data_path = os.path.join(DATA_DIR, "data.json")
subject_path = os.path.join(DATA_DIR, "subject.json")

# ------------------------------
# Cargar data JSON
# ------------------------------
with open(data_path, "r", encoding="utf-8") as file:
    data = json.load(file)

asesor = data["asesor"]
correo_asesor = data["correoAsesor"]
numero_asesor = data["numeroAsesor"]
clientes = data["datosClientes"]

# ------------------------------
# Cargar subjects
# ------------------------------
with open(subject_path, "r", encoding="utf-8") as file:
    subjects = json.load(file)["subjects"]  # debe ser lista: ["Asunto 1", "Asunto 2", ...]

# ------------------------------
# Cargar todas las plantillas HTML
# ------------------------------
html_templates = []
for filename in sorted(os.listdir(HTML_DIR)):
    if filename.endswith(".html"):
        path = os.path.join(HTML_DIR, filename)
        with open(path, "r", encoding="utf-8") as file:
            html_templates.append(file.read())

# ------------------------------
# Crear instancia de Outlook
# ------------------------------
outlook = win32.Dispatch('Outlook.Application')

# ------------------------------
# Enviar correos
# ------------------------------
for i, cliente in enumerate(clientes):
    nombre_cliente = cliente["nombre"]
    correo_cliente = cliente["correo"]

    # Selección cíclica de plantilla y asunto
    html_template = random.choice(html_templates)
    subject = random.choice(subjects)

    # Reemplazar variables dentro del HTML
    html_content = (
        html_template
        .replace("{{ cliente }}", nombre_cliente)
        .replace("{{ asesor }}", asesor)
        .replace("{{ correoAsesor }}", correo_asesor)
        .replace("{{ numeroAsesor }}", numero_asesor)
    )

    # Crear correo
    mail = outlook.CreateItem(0)
    mail.To = correo_cliente
    mail.Subject = subject
    mail.HTMLBody = html_content 

    # Enviar correo
    mail.Send()
    print(f"✔ Correo enviado a: {nombre_cliente} -> {correo_cliente}")
    
    
    # ----------------------------------------
    # 🔍 Verificación de hora límite (5:30 PM)
    # ----------------------------------------
    hora_actual = datetime.now()  # hora local del equipo (Colombia)
    hora_limite = hora_actual.replace(hour=17, minute=30, second=0, microsecond=0)

    if hora_actual > hora_limite:
        print("⛔ Envío detenido: Se alcanzó el límite horario (5:30 pm).")
        print(f"🛑 Último correo enviado fue: {nombre_cliente} -> {correo_cliente}")
        break

    # Espera para evitar bloqueos
    
    # ------------------------------
    # tiempo de espera para pruebas
    # time.sleep(3)
    # ------------------------------
    
    # ------------------------------
    # tiempo de espera para producción
    
    wait_time = random.uniform(120, 300)  # segundos
    print(f"⏱ Esperando {wait_time/60:.2f} minutos antes del siguiente envío...")
    time.sleep(wait_time)
    # ------------------------------

print("\n🎉 Todos los correos fueron enviados exitosamente.")
