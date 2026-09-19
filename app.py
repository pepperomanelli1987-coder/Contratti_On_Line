from flask import Flask, render_template, request, send_file
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

# Creazione automatica del database
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS contratti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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
    conn.commit()
    conn.close()

init_db()

@app.route("/")
def index():
    return render_template("index.html")

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

    # Inserimento nel DB
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT INTO contratti (
            nome_completo, luogo_nascita, data_nascita, cf,
            comune_residenza, via, telefono, email,
            metodo_pagamento, num_lampade, elenco_defunti,
            data_contratto, file_path
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        nome_completo, luogo_nascita, data_nascita, cf,
        comune_residenza, via, telefono, email,
        metodo_pagamento, num_lampade, elenco_defunti,
        data_contratto, ""
    ))
    conn.commit()

    # ID contratto generato
    contratto_id = c.lastrowid

    conn.close()

    # Compilazione DOCX
    doc = Document(MODELLO_PATH)

    segnaposto = {
        "{{ID_CONTRATTO}}": str(contratto_id),
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
    output_path = f"{OUTPUT_DIR}/contratto_{contratto_id}.docx"
    doc.save(output_path)

    # Aggiorna DB con percorso file
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE contratti SET file_path = ? WHERE id = ?", (output_path, contratto_id))
    conn.commit()
    conn.close()

    # Download del file
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
