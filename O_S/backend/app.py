from flask import Flask, render_template, request, redirect, url_for
from openpyxl import Workbook, load_workbook
import os
from functools import wraps
from flask import request, Response
import gspread
from google.oauth2.service_account import Credentials


# --- CONFIGURA TU USUARIO Y CONTRASEÑA AQUÍ ---

USERNAME = "COSTO2SAS"
PASSWORD = "1808"


def check_auth(username, password):
    return username == USERNAME and password == PASSWORD


def authenticate():
    return Response(
        "No autorizado.\n" "Debes iniciar sesión para acceder.",
        401,
        {"WWW-Authenticate": 'Basic realm="Login Required"'},
    )


def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)

    return decorated

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "..", "fronthead")

app = Flask(__name__, template_folder="../fronthead")

#Numero de formulario
'''
def obtener_numero_orden():
    archivo = "contador.txt"
    if not os.path.exists(archivo):
        with open(archivo, "w") as f:
            f.write("1")
        return 1

    with open(archivo, "r") as f:
        return int(f.read().strip())

def incrementar_numero_orden():
    numero = obtener_numero_orden() + 1
    with open("contador.txt", "w") as f:
        f.write(str(numero))
'''

# configurar googlesheets

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
import json
CREDS_FILE = json.loads(os.getenv("CREDS_JSON"))
SPREADSHEET_ID = "1iL1FLRb42b-mBZrSc2VVoC7NQJoQ8f420yrOaugrpKo"


def guardar_google_sheets(num_orden, nombre, cedula, empresa, vehiculo, telefono, items):

    # autenticacion google sheets
    creds = Credentials.from_service_account_info(CREDS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)

    # abrir hoja
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1

    # guardar items
    for index, item in enumerate(items, start=1):
        fila = [
            num_orden,
            index,
            nombre,
            cedula,
            empresa,
            vehiculo,
            telefono,
            item["Marca"],
            item["Referencia"],
            item["Serie"],
            item["Servicio"],
        ]
        sheet.append_row(fila)

def obtener_numero_orden():
    """
    Obtiene el siguiente número de orden leyendo 
    el último valor registrado en Google Sheets.
    """
    creds = Credentials.from_service_account_info(CREDS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)

    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    columna_orden = sheet.col_values(1)  # columna A completa

    # Si no hay registros, empezamos en 1
    if len(columna_orden) <= 1:
        return 1

    try:
        ultimo = int(columna_orden[-1])
    except:
        ultimo = 0

    return ultimo + 1


def guardar_excel(num_orden, nombre, cedula, empresa, vehiculo, telefono, items):

    excel_file = "registros.xlsx"

    if not os.path.exists(excel_file):
        # Crear un nuevo archivo Excel y agregar encabezados
        wb = Workbook()
        ws = wb.active
        ws.title = "Registros"
        ws.append(
            [
                "Num_Orden",
                "Item",
                "Nombre",
                "Cedula",
                "Empresa",
                "Vehiculo",
                "Telefono",
                "Marca",
                "Referencia",
                "Serie",
                "Servicio",
            ]
        )
    else:
        # Abrir archivo existente y agregar el registro
        wb = load_workbook(excel_file)
        ws = wb.active

    # guardar cada item como fila
    for index, item in enumerate(items, start=1):
        ws.append([
                num_orden,
                index,
                nombre,
                cedula,
                empresa,
                vehiculo,
                telefono,
                item["Marca"],
                item["Referencia"],
                item["Serie"],
                item["Servicio"],
            ])

    wb.save(excel_file)


# ruta principal - formualrio


@app.route("/")
@require_auth
def index():
    num_orden = obtener_numero_orden()
    return render_template("index.html", num_orden=num_orden)


# ruta formulario


@app.route("/register", methods=["POST"])
@require_auth
def register():

    num_orden = obtener_numero_orden()
    nombre = request.form["Nombre"]
    cedula = request.form["Cedula"]
    empresa = request.form["Empresa"]
    vehiculo = request.form["Vehiculo"]
    telefono = request.form["Telefono"]

    # datos tabla

    marca = request.form.getlist("Marca[]")
    referencia = request.form.getlist("Referencias[]")
    serie = request.form.getlist("Serie[]")
    servicio = request.form.getlist("Tipos[]")

    # combinar items

    items = []
    for s, c, v, st in zip(marca, referencia, serie, servicio):
        items.append({"Marca": s, "Referencia": c, "Serie": v, "Servicio": st})

    # guardar excel
    guardar_excel(num_orden, nombre, cedula, empresa, vehiculo, telefono, items)
    # Guardar en Google Sheets
    # guardar_google_sheets(num_orden, nombre, cedula, empresa, vehiculo, telefono, items)

    incrementar_numero_orden()

    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)


