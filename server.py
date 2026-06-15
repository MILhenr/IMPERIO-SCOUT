import os, uuid, json
import urllib.request, urllib.parse, urllib.error
from datetime import datetime
from flask import Flask, request, jsonify, send_file
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)
DATABASE_URL = os.environ.get("DATABASE_URL")
ANALISE_URL = os.environ.get("ANALISE_URL", "")
ANALISE_BOT_SECRET = os.environ.get("ANALISE_BOT_SECRET", "scoutbot_secret_2024")
LIGA_PARA_CAT = {"lnf": "Nacional", "base": "Base", "outros": "Outros"}

def get_db():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor, sslmode="require")

def init_db():
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute('''CREATE TABLE IF NOT EXISTS times (
                id TEXT PRIMARY KEY,
                nome TEXT NOT NULL,
                emoji TEXT DEFAULT '⚽',
                liga TEXT DEFAULT 'lnf',
                jogos INTEGER DEFAULT 0,
                criado_em TEXT,
                import_id TEXT
            )''')
            c.execute('''CREATE TABLE IF NOT EXISTS jogadores (
                id TEXT PRIMARY KEY,
                time_id TEXT REFERENCES times(id) ON DELETE CASCADE,
                num TEXT DEFAULT '—',
                nome TEXT NOT NULL,
                gols INTEGER DEFAULT 0,
                passes INTEGER DEFAULT 0,
                criado_em TEXT
            )''')
            c.execute('''ALTER TABLE times ADD COLUMN IF NOT EXISTS import_id TEXT''')
        conn.commit()
    print("✅ DB OK")

init_db()

def sync_analise(jogador, time):
    if not ANALISE_URL:
        return
    gols = jogador.get("gols", 0)
    passes = jogador.get("passes", 0)
    if gols == 0 and passes == 0:
        return
    cat = LIGA_PARA_CAT.get(time.get("liga", ""), "Outros")
    comp = time.get("nome", "")
    clube = time.get("nome", "")
    payload = {
        "nome": jogador["nome"],
        "idade": 0,
        "posicao": "",
        "modalidade": "Futsal",
        "clube": clube,
        "pe": "Direito",
        "disponivel": True,
        "cat1": cat,
        "comp1": comp,
        "stats_gols1": str(gols),
        "stats_assists1": str(passes),
        "stats_jogos1": str(time.get("jogos", 0)),
        "stats_gols": str(gols),
        "stats_assists": str(passes),
        "stats_jogos": str(time.get("jogos", 0)),
        "status": "aprovado",
        "foto": None
    }
    try:
        busca_url = f"{ANALISE_URL}/api/bot/atletas?nome={urllib.parse.quote(jogador['nome'])}&clube={urllib.parse.quote(clube)}"
        req = urllib.request.Request(busca_url, headers={"X-Bot-Secret": ANALISE_BOT_SECRET})
        with urllib.request.urlopen(req, timeout=8) as r:
            encontrados = json.loads(r.read())
        if encontrados:
            aid = encontrados[0]["id"]
            update_payload = {
                "nome": jogador["nome"],
                "idade": int(encontrados[0].get("idade") or 0),
                "posicao": encontrados[0].get("posicao", ""),
                "modalidade": "Futsal",
                "clube": clube,
                "pe": encontrados[0].get("pe", "Direito"),
                "disponivel": True,
                "cat1": cat,
                "comp1": comp,
                "stats_gols1": str(gols),
                "stats_assists1": str(passes),
                "stats_jogos1": str(time.get("jogos", 0)),
                "stats_gols": str(gols),
                "stats_assists": str(passes),
                "stats_jogos": str(time.get("jogos", 0)),
                "status": "aprovado",
                "foto": encontrados[0].get("foto")
            }
            req2 = urllib.request.Request(
                f"{ANALISE_URL}/api/sync/atleta/{aid}",
                data=json.dumps(update_payload).encode(),
                headers={"Content-Type": "application/json", "X-Bot-Secret": ANALISE_BOT_SECRET},
                method="PUT"
            )
            urllib.request.urlopen(req2, timeout=8)
        else:
            req2 = urllib.request.Request(
                f"{ANALISE_URL}/api/sync/atleta",
                data=json.dumps(payload).encode(),
                headers={"Content-Type": "application/json", "X-Bot-Secret": ANALISE_BOT_SECRET},
                method="POST"
            )
            urllib.request.urlopen(req2, timeout=8)
    except Exception as e:
        print(f"⚠️ Sync Analise.io falhou: {e}")


@app.route('/')
def index():
    return send_file('scout_stats.html')

# ── TIMES ──────────────────────────────────────────────────────
@app.route('/api/times', methods=['GET'])
def listar_times():
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("SELECT * FROM times ORDER BY criado_em DESC")
            times = [dict(r) for r in c.fetchall()]
            for t in times:
                c.execute("SELECT * FROM jogadores WHERE time_id=%s ORDER BY nome", (t['id'],))
                t['jogadores'] = [dict(j) for j in c.fetchall()]
    return jsonify(times)

@app.route('/api/times', methods=['POST'])
def criar_time():
    d = request.json
    tid = str(uuid.uuid4())[:8]
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("INSERT INTO times (id,nome,emoji,liga,jogos,criado_em,import_id) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (tid, d['nome'], d.get('emoji','⚽'), d.get('liga','lnf'), 0, datetime.now().isoformat(), None))
        conn.commit()
    return jsonify({"ok": True, "id": tid})

@app.route('/api/times/<tid>', methods=['PUT'])
def editar_time(tid):
    d = request.json
    with get_db() as conn:
        with conn.cursor() as c:
            if 'jogos' in d:
                c.execute("UPDATE times SET jogos=%s WHERE id=%s", (d['jogos'], tid))
            if 'nome' in d:
                c.execute("UPDATE times SET nome=%s, emoji=%s, liga=%s WHERE id=%s",
                    (d['nome'], d.get('emoji','⚽'), d.get('liga','lnf'), tid))
        conn.commit()
    return jsonify({"ok": True})

@app.route('/api/times/<tid>', methods=['DELETE'])
def deletar_time(tid):
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("DELETE FROM times WHERE id=%s", (tid,))
        conn.commit()
    return jsonify({"ok": True})

# ── JOGADORES ──────────────────────────────────────────────────
@app.route('/api/times/<tid>/jogadores', methods=['POST'])
def criar_jogador(tid):
    d = request.json
    jid = str(uuid.uuid4())[:8]
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("INSERT INTO jogadores (id,time_id,num,nome,gols,passes,criado_em) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (jid, tid, d.get('num','—'), d['nome'], d.get('gols',0), d.get('passes',0), datetime.now().isoformat()))
        conn.commit()
    return jsonify({"ok": True, "id": jid})

@app.route('/api/jogadores/<jid>', methods=['PUT'])
def editar_jogador(jid):
    d = request.json
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("UPDATE jogadores SET gols=%s, passes=%s WHERE id=%s",
                (d['gols'], d['passes'], jid))
            c.execute("SELECT * FROM jogadores WHERE id=%s", (jid,))
            jogador = dict(c.fetchone())
            c.execute("SELECT * FROM times WHERE id=%s", (jogador['time_id'],))
            time = dict(c.fetchone())
        conn.commit()
    try:
        sync_analise(jogador, time)
    except Exception as e:
        print(f"⚠️ Sync error: {e}")
    return jsonify({"ok": True})

@app.route('/api/jogadores/<jid>', methods=['DELETE'])
def deletar_jogador(jid):
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("DELETE FROM jogadores WHERE id=%s", (jid,))
        conn.commit()
    return jsonify({"ok": True})

# ── BACKUP ────────────────────────────────────────────────────
@app.route('/api/backup', methods=['GET'])
def backup():
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("SELECT * FROM times ORDER BY criado_em")
            times = [dict(r) for r in c.fetchall()]
            for t in times:
                c.execute("SELECT * FROM jogadores WHERE time_id=%s", (t['id'],))
                t['jogadores'] = [dict(j) for j in c.fetchall()]
    return jsonify({"backup": datetime.now().isoformat(), "times": times})

# ── MIGRAÇÃO ─────────────────────────────────────────────────
@app.route('/api/migrar', methods=['POST'])
def migrar():
    dados = request.json
    count_t = 0; count_j = 0
    import_id = str(uuid.uuid4())[:8]
    with get_db() as conn:
        with conn.cursor() as c:
            for t in dados:
                tid = str(uuid.uuid4())[:8]
                c.execute("INSERT INTO times (id,nome,emoji,liga,jogos,criado_em,import_id) VALUES (%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                    (tid, t['nome'], t.get('emoji','⚽'), t.get('liga','lnf'), t.get('jogos',0), datetime.now().isoformat(), import_id))
                count_t += 1
                for j in t.get('jogadores',[]):
                    jid = str(uuid.uuid4())[:8]
                    c.execute("INSERT INTO jogadores (id,time_id,num,nome,gols,passes,criado_em) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                        (jid, tid, j.get('num','—'), j['nome'], j.get('gols',0), j.get('passes',0), datetime.now().isoformat()))
                    count_j += 1
        conn.commit()
    return jsonify({"ok": True, "times": count_t, "jogadores": count_j, "import_id": import_id})

# ── LISTAR IMPORTS ────────────────────────────────────────────
@app.route('/api/imports', methods=['GET'])
def listar_imports():
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("""
                SELECT import_id, COUNT(*) as total_times, MIN(criado_em) as criado_em
                FROM times
                WHERE import_id IS NOT NULL
                GROUP BY import_id
                ORDER BY criado_em DESC
            """)
            imports = [dict(r) for r in c.fetchall()]
    return jsonify(imports)

# ── APAGAR IMPORT ─────────────────────────────────────────────
@app.route('/api/imports/<import_id>', methods=['DELETE'])
def apagar_import(import_id):
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("SELECT COUNT(*) as total FROM times WHERE import_id=%s", (import_id,))
            total = c.fetchone()['total']
            if total == 0:
                return jsonify({"erro": "Import não encontrado"}), 404
            c.execute("DELETE FROM times WHERE import_id=%s", (import_id,))
        conn.commit()
    return jsonify({"ok": True, "times_apagados": total})

# ── APAGAR TUDO ───────────────────────────────────────────────
@app.route('/api/apagar-tudo', methods=['DELETE'])
def apagar_tudo():
    with get_db() as conn:
        with conn.cursor() as c:
            c.execute("DELETE FROM jogadores")
            c.execute("DELETE FROM times")
        conn.commit()
    return jsonify({"ok": True})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 3000))
    app.run(host='0.0.0.0', port=port, debug=False)
