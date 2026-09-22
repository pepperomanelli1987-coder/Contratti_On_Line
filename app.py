from flask import Flask, render_template, request, redirect, session, send_file
import sqlite3
import os
from docx import Document
from datetime import datetime

app = Flask(__name__)
app.secret_key = "supersegreto123"  # Cambiala se vuoi

PASSWORD = "mgservizi"  # Password di accesso

DB_PATH = "contratti.db"
OUTPUT_DIR = "contratti_generati"
MODELLO = "modello_contratto_lampade.docx"

# Assicura che la cartella dei contratti esista
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# ---------------------------
# LOGIN
# ---------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if request.form.get("password") == PASSWORD:
            session["logged_in"] = True
            return redirect("/")
        else:
            return render_template("login.html", error=True)

    return render_template("login.html", error=False)


@app.route("/logout")
def logout():
    session["logged_in"] = False
    return redirect("/login")


# ---------------------------
# HOME PROTETTA
# ---------------------------

@app.route("/")
def index():
    if not session.get("logged_in"):
        return redirect("/login")
    return render_template("index.html")


# ---------------------------
# GENERA CONTRATTO
# ---------------------------

@app.route("/genera", methods=["POST"])
def genera():
    if not session.get("logged_in"):
        return redirect("/login")

    nome_completo = request.form["nome_completo"]
    luogo_nascita = request.form["luogo_nascita"]
    data_nascita = request.form["data_nascita"]
    cf = request.form["cf"]
    comune_residenza = request.form["comune_residenza"]
    via = request.form["via"]
    telefono = request.form["telefono"]
    email = request.form["email"]
    metodo_pagamento = request.form["metodo_pagamento"]
    num_lampade = request.form["num_lampade"]
    elenco_defunti = request.form["elenco_defunti"]

    data_contratto = datetime.now().strftime("%d/%m/%Y")

    # Connessione DB
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Crea tabella se non esiste
    c.execute("""
        CREATE TABLE IF NOT EXISTS contratti (
            id_contratto INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_completo TEXT,
            file_path TEXT
        )
    """)

    # Inserisci nuovo contratto
    c.execute("INSERT INTO contratti (nome_completo) VALUES (?)", (nome_completo,))
    conn.commit()

    id_contratto = c.lastrowid

    # Normalizza nome file
    nome_file = nome_completo.replace(" ", "_").replace("/", "_")

    output_path = f"{OUTPUT_DIR}/contratto_{id_contratto}_{nome_file}.docx"

    # Carica modello
    doc = Document(MODELLO)

    # Sostituzioni nel modello (MAIUSCOLO come nel DOCX)
    for p in doc.paragraphs:
        p.text = p.text.replace("{{ID_CONTRATTO}}", str(id_contratto))
        p.text = p.text.replace("{{NOME_COMPLETO}}", nome_completo)
        p.text = p.text.replace("{{LUOGO_NASCITA}}", luogo_nascita)
        p.text = p.text.replace("{{DATA_NASCITA}}", data_nascita)
        p.text = p.text.replace("{{CF}}", cf)
        p.text = p.text.replace("{{COMUNE_RESIDENZA}}", comune_residenza)
        p.text = p.text.replace("{{VIA}}", via)
        p.text = p.text.replace("{{TELEFONO}}", telefono)
        p.text = p.text.replace("{{EMAIL}}", email)
        p.text = p.text.replace("{{METODO_PAGAMENTO}}", metodo_pagamento)
        p.text = p.text.replace("{{NUM_LAMPADE}}", num_lampade)
        p.text = p.text.replace("{{ELENCO_DEFUNTI}}", elenco_defunti)
        p.text = p.text.replace("{{DATA_CONTRATTO}}", data_contratto)

    # Salva contratto
    doc.save(output_path)

    # Aggiorna DB con percorso file
    c.execute("UPDATE contratti SET file_path = ? WHERE id_contratto = ?", (output_path, id_contratto))
    conn.commit()
    conn.close()

    return send_file(output_path, as_attachment=True)


# ---------------------------
# AVVIO SERVER
# ---------------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
