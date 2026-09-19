from flask import Flask, render_template, request, send_file, redirect, url_for
import sqlite3
import os
from datetime import datetime
from docx import Document

app = Flask(__name__)

DB_PATH = "contratti.db"
MODELLO_PATH = "modello_contratto_lampade.docx"
OUTPUT_DIR = "contratti_generati"

# Creazione cartella output
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# Creazione database + tabella config
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS contratti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_contratto INTEGER,
            nome_completo TEXT,
            luogo_nascita TEXT,
            data_nascita TEXT,
            cf TEXT,
            comune_residenza TEXT,
            via TEXT,
            telefono TEXT,
            email TEXT,
            metodo_pagamento TEXT,
            num_lampade INTEGER,
            elenco_defunti TEXT,
            data_contratto TEXT,
            file_path TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS config (
            chiave TEXT PRIMARY KEY,
            valore INTEGER
        )
    """)

    # Se non esiste ultimo_id, lo inizializziamo a 235
    c.execute("SELECT valore FROM config WHERE chiave='ultimo_id'")
    row = c.fetchone()
    if row is None:
        c.execute("INSERT INTO config (chiave, valore) VALUES ('ultimo_id', 235)")

    conn.commit()
    conn.close()

init_db()

@app.route("/")
def index():
    success = request.args.get("success")
    id_generato = request.args.get("id")
    return render_template("index.html", success=success, id_generato=id_generato)

@app.route("/genera", methods=["POST"])
def genera():
    # Dati dal form
    nome_completo = request.form.get("nome_completo")
    luogo_nascita = request.form.get("luogo_nascita")
    data_nascita = request.form.get("data_nascita")
    cf = request.form.get("cf")
    comune_residenza = request.form.get("comune_residenza")
    via = request.form.get("via")
    telefono = request.form.get("telefono")
    email = request.form.get("email")
    metodo_pagamento = request.form.get("metodo_pagamento")
    num_lampade = request.form.get("num_lampade")
    elenco_defunti = request.form.get("elenco_defunti")

    # Data automatica
    data_contratto = datetime.now().strftime("%d/%m/%Y")

    # Recupero numerazione
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute("SELECT valore FROM config WHERE chiave='ultimo_id'")
    ultimo_id = c.fetchone()[0]

    id_contratto = ultimo_id + 1

    # Aggiorna numerazione
    c.execute("UPDATE config SET valore=? WHERE chiave='ultimo_id'", (id_contratto,))

    # Inserimento nel DB
    c.execute("""
        INSERT INTO contratti (
            id_contratto, nome_completo, luogo_nascita, data_nascita, cf,
            comune_residenza, via, telefono, email,
            metodo_pagamento, num_lampade, elenco_defunti,
            data_contratto, file_path
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        id_contratto, nome_completo, luogo_nascita, data_nascita, cf,
        comune_residenza, via, telefono, email,
        metodo_pagamento, num_lampade, elenco_defunti,
        data_contratto, ""
    ))

    conn.commit()
    conn.close()

    # Compilazione DOCX
    doc = Document(MODELLO_PATH)

    segnaposto = {
        "{{ID_CONTRATTO}}": str(id_contratto),
        "{{NOME_COMPLETO}}": nome_completo,
        "{{LUOGO_NASCITA}}": luogo_nascita,
        "{{DATA_NASCITA}}": data_nascita,
        "{{CF}}": cf,
        "{{COMUNE_RESIDENZA}}": comune_residenza,
        "{{VIA}}": via,
        "{{TELEFONO}}": telefono,
        "{{EMAIL}}": email,
        "{{METODO_PAGAMENTO}}": metodo_pagamento,
        "{{NUM_LAMPADE}}": num_lampade,
        "{{ELENCO_DEFUNTI}}": elenco_defunti,
        "{{DATA_CONTRATTO}}": data_contratto
    }

    for p in doc.paragraphs:
        for key, value in segnaposto.items():
            if key in p.text:
                p.text = p.text.replace(key, value)

    # Salvataggio file
    output_path = f"{OUTPUT_DIR}/contratto_{id_contratto}.docx"
    doc.save(output_path)

    # Aggiorna DB con percorso file
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE contratti SET file_path = ? WHERE id_contratto = ?", (output_path, id_contratto))
    conn.commit()
    conn.close()

    # Download del file
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
