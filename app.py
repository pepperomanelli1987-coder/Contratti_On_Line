from flask import Flask, render_template, request
import sqlite3
import os
from datetime import datetime

app = Flask(__name__)

DB_PATH = "contratti.db"

# Creazione automatica del database e della tabella
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS contratti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_cliente TEXT,
            indirizzo TEXT,
            numero_lampade INTEGER,
            note TEXT,
            data TEXT
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
    nome = request.form.get("nome")
    indirizzo = request.form.get("indirizzo")
    lampade = request.form.get("lampade")
    note = request.form.get("note")

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT INTO contratti (nome_cliente, indirizzo, numero_lampade, note, data) VALUES (?, ?, ?, ?, ?)",
              (nome, indirizzo, lampade, note, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()

    return "Contratto generato correttamente!"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
