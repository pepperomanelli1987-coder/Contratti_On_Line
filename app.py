from flask import Flask, render_template, request, send_file
from docx import Document
import os
import datetime
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)

# Percorso base del progetto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODELLO = os.path.join(BASE_DIR, "modello_contratto_lampade.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "Contratti")

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# Connessione al database PostgreSQL (Render)
def get_db():
    return psycopg2.connect(
        host=os.environ.get("DB_HOST"),
        database=os.environ.get("DB_NAME"),
        user=os.environ.get("DB_USER"),
        password=os.environ.get("DB_PASS")
    )

def genera_id_progressivo():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT MAX(id) FROM contratti;")
    risultato = cur.fetchone()[0]
    conn.close()

    if risultato is None:
        return "0001"
    else:
        return str(risultato + 1).zfill(4)

def salva_contratto_db(id_contratto, nome, cf, percorso):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO contratti (id, nome_completo, cf, data_contratto, percorso_file) VALUES (%s, %s, %s, %s, %s)",
        (int(id_contratto), nome, cf, datetime.date.today(), percorso)
    )
    conn.commit()
    conn.close()

def compila_contratto(dati, percorso_output):
    doc = Document(MODELLO)

    segnaposto = {
        "{{ID_CONTRATTO}}": dati["id_contratto"],
        "{{NOME_COMPLETO}}": dati["nome_completo"],
        "{{LUOGO_NASCITA}}": dati["luogo_nascita"],
        "{{DATA_NASCITA}}": dati["data_nascita"],
        "{{CF}}": dati["cf"],
        "{{COMUNE_RESIDENZA}}": dati["comune_residenza"],
        "{{VIA}}": dati["via"],
        "{{TELEFONO}}": dati["telefono"],
        "{{EMAIL}}": dati["email"],
        "{{METODO_PAGAMENTO}}": dati["metodo_pagamento"],
        "{{NUM_LAMPADE}}": dati["num_lampade"],
        "{{ELENCO_DEFUNTI}}": dati["elenco_defunti"],
        "{{DATA_CONTRATTO}}": dati["data_contratto"],
    }

    for p in doc.paragraphs:
        for key, val in segnaposto.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in segnaposto.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, val)

    doc.save(percorso_output)

@app.route("/", methods=["GET"])
def form():
    return render_template("form.html")

@app.route("/genera", methods=["POST"])
def genera():
    id_progressivo = genera_id_progressivo()
    data_oggi = datetime.date.today().strftime("%d/%m/%Y")

    dati = {
        "id_contratto": id_progressivo,
        "nome_completo": request.form["nome_completo"].strip(),
        "luogo_nascita": request.form["luogo_nascita"].strip(),
        "data_nascita": request.form["data_nascita"].strip(),
        "cf": request.form["cf"].strip(),
        "comune_residenza": request.form["comune_residenza"].strip(),
        "via": request.form["via"].strip(),
        "telefono": request.form.get("telefono", "").strip(),
        "email": request.form.get("email", "").strip(),
        "metodo_pagamento": request.form["metodo_pagamento"].strip(),
        "num_lampade": request.form["num_lampade"].strip(),
        "elenco_defunti": request.form.get("elenco_defunti", "").strip(),
        "data_contratto": data_oggi,
    }

    nome_cartella = f"{id_progressivo}_{dati['nome_completo'].replace(' ', '_')}_{dati['cf']}"
    cartella_cliente = os.path.join(OUTPUT_DIR, nome_cartella)
    os.makedirs(cartella_cliente, exist_ok=True)

    nome_file = f"Contratto_{dati['nome_completo'].replace(' ', '_')}.docx"
    percorso_output = os.path.join(cartella_cliente, nome_file)

    compila_contratto(dati, percorso_output)

    salva_contratto_db(id_progressivo, dati["nome_completo"], dati["cf"], percorso_output)

    return send_file(percorso_output, as_attachment=True)

if __name__ == "__main__":
    app.run()
