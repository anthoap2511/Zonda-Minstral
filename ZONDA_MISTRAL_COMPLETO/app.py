# -*- coding: utf-8 -*-
import os, sqlite3, uuid, json, secrets
from datetime import datetime, date
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, abort
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from engine import process_pdfs, make_excel

BASE=Path(__file__).resolve().parent
DATA_DIR=Path(os.environ.get("ZONDA_DATA_DIR", str(BASE))).resolve()
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB=DATA_DIR/"zonda.db"; UPLOAD=DATA_DIR/"uploads"; EXPORT=DATA_DIR/"exports"
UPLOAD.mkdir(parents=True, exist_ok=True); EXPORT.mkdir(parents=True, exist_ok=True)
app=Flask(__name__)
app.config.update(
 SESSION_COOKIE_HTTPONLY=True,
 SESSION_COOKIE_SAMESITE="Lax",
 SESSION_COOKIE_SECURE=os.environ.get("RENDER") == "true",
 MAX_CONTENT_LENGTH=30*1024*1024,
)
# Clave estable local: se genera una sola vez si no se configura SECRET_KEY.
secret_file=BASE/".secret_key"
if os.environ.get("SECRET_KEY"):
 app.secret_key=os.environ["SECRET_KEY"]
else:
 if not secret_file.exists(): secret_file.write_text(secrets.token_hex(32),encoding="utf-8")
 app.secret_key=secret_file.read_text(encoding="utf-8").strip()
ALLOWED={"pdf"}

DEFAULT_SETTINGS={
 "numero_rendicion":"1","descripcion_objeto":"ATENCIÓN DE EQUIPO / SERVICIO REALIZADO","detalle_gasto":"",
 "ubicacion":"",
 "fecha_transferencia":date.today().strftime("%Y-%m-%d"),"numero_cuenta":"","importe_recibido":"0",
 "tecnico_nombre":"","tecnico_dni":"","fecha_firma":"",
 "texto_rendicion":"",
 "titulo_movilidad":"PLANILLA DE MOVILIDAD {nombre}","incluir_movilidad":True
}
SETTING_FIELDS=["numero_rendicion","descripcion_objeto","ubicacion","fecha_transferencia","numero_cuenta","importe_recibido","tecnico_nombre","tecnico_dni","fecha_firma","texto_rendicion","titulo_movilidad"]

MESES_ES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
def fecha_larga_hoy():
 hoy=date.today()
 return f"Lima, {hoy.day} de {MESES_ES[hoy.month-1]} del {hoy.year}"

def siguiente_numero_rendicion(c):
 rows=c.execute("SELECT id, settings_json FROM rendiciones ORDER BY id DESC").fetchall()
 nums=[]
 # La secuencia AUTOINCREMENT conserva el máximo histórico aunque se elimine
 # la rendición con el número más alto.
 try:
  seqrow=c.execute("SELECT seq FROM sqlite_sequence WHERE name='rendiciones'").fetchone()
  if seqrow and seqrow[0] is not None: nums.append(int(seqrow[0]))
 except sqlite3.OperationalError: pass
 for row in rows:
  try:
   value=json.loads(row["settings_json"] or "{}").get("numero_rendicion","")
   if str(value).isdigit(): nums.append(int(value))
  except Exception: pass
 return str(max(nums, default=0)+1)

def normalizar_fecha_transferencia(value):
 value=(value or "").strip()
 for fmt in ("%Y-%m-%d","%d/%m/%Y","%d%m%Y"):
  try: return datetime.strptime(value,fmt).strftime("%Y-%m-%d")
  except ValueError: pass
 raise ValueError("Ingresa una fecha válida.")

def db():
 c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init_db():
 c=db()
 c.executescript("""
 CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT UNIQUE NOT NULL,password TEXT NOT NULL,role TEXT NOT NULL DEFAULT 'user',created_at TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS rendiciones(id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER NOT NULL,title TEXT NOT NULL,status TEXT NOT NULL DEFAULT 'Creada',created_at TEXT NOT NULL,updated_at TEXT NOT NULL,settings_json TEXT,records_json TEXT,result_file TEXT,error_text TEXT,FOREIGN KEY(user_id) REFERENCES users(id));
 CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY AUTOINCREMENT,rendicion_id INTEGER NOT NULL,original_name TEXT NOT NULL,stored_name TEXT NOT NULL,uploaded_at TEXT NOT NULL,FOREIGN KEY(rendicion_id) REFERENCES rendiciones(id));
 CREATE TABLE IF NOT EXISTS history(id INTEGER PRIMARY KEY AUTOINCREMENT,rendicion_id INTEGER NOT NULL,user_id INTEGER NOT NULL,action TEXT NOT NULL,detail TEXT,created_at TEXT NOT NULL);
 """)
 # Migración segura para bases creadas con la versión anterior.
 cols={x["name"] for x in c.execute("PRAGMA table_info(rendiciones)")}
 for name,typ in [("settings_json","TEXT"),("records_json","TEXT"),("result_file","TEXT"),("error_text","TEXT")]:
  if name not in cols: c.execute(f"ALTER TABLE rendiciones ADD COLUMN {name} {typ}")
 if not c.execute("SELECT 1 FROM users").fetchone():
  if os.environ.get("RENDER") == "true":
   admin_username=os.environ.get("INITIAL_ADMIN_USERNAME", "admin").strip()
   admin_password=os.environ.get("INITIAL_ADMIN_PASSWORD", "")
   if len(admin_password) < 12:
    raise RuntimeError("Configura INITIAL_ADMIN_PASSWORD (mínimo 12 caracteres) en Render antes de iniciar.")
  else:
   admin_username="admin"
   admin_password=os.environ.get("INITIAL_ADMIN_PASSWORD", "Admin123!")
  c.execute("INSERT INTO users(username,password,role,created_at) VALUES(?,?,?,?)",
            (admin_username,generate_password_hash(admin_password),"admin",datetime.now().isoformat()))
 c.commit(); c.close()

def log(rid,uid,action,detail=""):
 c=db(); c.execute("INSERT INTO history(rendicion_id,user_id,action,detail,created_at) VALUES(?,?,?,?,?)",(rid,uid,action,detail,datetime.now().isoformat())); c.commit(); c.close()

def current(): return session.get("user")
def require_login():
 if not current(): return redirect(url_for("login"))
 return None

def get_rendicion(rid):
 c=db(); row=c.execute("SELECT r.*,u.username FROM rendiciones r JOIN users u ON u.id=r.user_id WHERE r.id=?",(rid,)).fetchone(); c.close()
 if not row: abort(404)
 if current()["role"]!="admin" and row["user_id"]!=current()["id"]: abort(403)
 return row

def settings_for(row):
 try: data=json.loads(row["settings_json"] or "{}")
 except Exception: data={}
 out=DEFAULT_SETTINGS.copy(); out.update(data)
 return out

def docs_for(rid):
 c=db(); docs=c.execute("SELECT * FROM documents WHERE rendicion_id=? ORDER BY id",(rid,)).fetchall(); c.close(); return docs

@app.after_request
def security_headers(response):
 response.headers.setdefault("X-Content-Type-Options", "nosniff")
 response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
 response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
 return response

@app.route("/healthz")
def healthz(): return {"status":"ok"}, 200

@app.route("/")
def index(): return redirect(url_for("dashboard") if current() else url_for("login"))

@app.route("/login",methods=["GET","POST"])
def login():
 if request.method=="POST":
  c=db(); u=c.execute("SELECT * FROM users WHERE username=?",(request.form.get("username","").strip(),)).fetchone(); c.close()
  if u and check_password_hash(u["password"],request.form.get("password","")):
   session.clear(); session["user"]={"id":u["id"],"username":u["username"],"role":u["role"]}; return redirect(url_for("dashboard"))
  flash("Usuario o contraseña incorrectos.","error")
 return render_template("login.html")

@app.route("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

@app.route("/mi-cuenta",methods=["GET","POST"])
def mi_cuenta():
 if (r:=require_login()): return r
 if request.method=="POST":
  old=request.form.get("current_password",""); new=request.form.get("new_password","")
  if len(new)<8: flash("La nueva contraseña debe tener al menos 8 caracteres.","error")
  else:
   c=db(); u=c.execute("SELECT * FROM users WHERE id=?",(current()["id"],)).fetchone()
   if not u or not check_password_hash(u["password"],old):
    flash("La contraseña actual no es correcta.","error")
   else:
    c.execute("UPDATE users SET password=? WHERE id=?",(generate_password_hash(new),current()["id"])); c.commit()
    flash("Contraseña actualizada correctamente.","ok")
   c.close()
 return render_template("account.html")

@app.route("/dashboard")
def dashboard():
 if (r:=require_login()): return r
 c=db()
 if current()["role"]=="admin":
  rows=c.execute("SELECT r.*,u.username FROM rendiciones r JOIN users u ON u.id=r.user_id ORDER BY r.id DESC").fetchall()
 else:
  rows=c.execute("SELECT r.*,u.username FROM rendiciones r JOIN users u ON u.id=r.user_id WHERE r.user_id=? ORDER BY r.id DESC",(current()["id"],)).fetchall()
 c.close(); return render_template("dashboard.html",rows=rows)

@app.route("/contabilidad")
def contabilidad():
 if (r:=require_login()): return r
 if current()["role"]!="admin": abort(403)
 c=db()
 rows=c.execute("SELECT r.id,r.title,r.status,r.created_at,r.settings_json,r.records_json,u.username "
                "FROM rendiciones r JOIN users u ON u.id=r.user_id ORDER BY r.id DESC").fetchall()
 c.close()
 total_gastos=0.0; total_recibido=0.0; total_completadas=0
 categorias={"Alimentación":0.0,"Movilidad":0.0,"Hospedaje":0.0,"Otros":0.0}
 meses={}; movimientos=[]
 for row in rows:
  try: settings=json.loads(row["settings_json"] or "{}")
  except Exception: settings={}
  try: records=json.loads(row["records_json"] or "[]")
  except Exception: records=[]
  try: recibido=float(str(settings.get("importe_recibido","0")).replace(",","") or 0)
  except (TypeError,ValueError): recibido=0.0
  gasto=0.0
  for rec in records:
   try: importe=float(str(rec.get("total") or 0).replace(",",""))
   except (TypeError,ValueError): importe=0.0
   gasto+=importe
   detalle=str(rec.get("detalle","")).upper()
   if "ALIMENT" in detalle or "ALIMENTO" in detalle: categoria="Alimentación"
   elif "MOVILIDAD" in detalle: categoria="Movilidad"
   elif "HOSPEDAJE" in detalle: categoria="Hospedaje"
   else: categoria="Otros"
   categorias[categoria]+=importe
  total_gastos+=gasto; total_recibido+=recibido
  if row["status"]=="Completada": total_completadas+=1
  mes=str(row["created_at"] or "")[:7] or "Sin fecha"
  if mes not in meses: meses[mes]={"mes":mes,"gastos":0.0,"recibido":0.0,"cantidad":0}
  meses[mes]["gastos"]+=gasto; meses[mes]["recibido"]+=recibido; meses[mes]["cantidad"]+=1
  movimientos.append({"id":row["id"],"title":row["title"],"usuario":row["username"],
    "estado":row["status"],"gastos":gasto,"recibido":recibido,"saldo":recibido-gasto,"mes":mes})
 max_categoria=max(categorias.values(),default=0)
 categoria_rows=[{"nombre":k,"total":v,"porcentaje":(v/max_categoria*100 if max_categoria else 0)}
                 for k,v in categorias.items()]
 return render_template("contabilidad.html",total_gastos=total_gastos,total_recibido=total_recibido,
  saldo=total_recibido-total_gastos,total_completadas=total_completadas,total_rendiciones=len(rows),
  categorias=categoria_rows,meses=sorted(meses.values(),key=lambda x:x["mes"],reverse=True)[:12],
  movimientos=movimientos[:15])

@app.route("/rendicion/nueva",methods=["GET","POST"])
def new_rendicion():
 if (r:=require_login()): return r
 if request.method=="POST":
  title=request.form.get("title","").strip() or "Nueva rendición"; now=datetime.now().isoformat()
  c=db(); settings=DEFAULT_SETTINGS.copy(); settings["numero_rendicion"]=siguiente_numero_rendicion(c); settings["fecha_firma"]=fecha_larga_hoy()
  cur=c.execute("INSERT INTO rendiciones(user_id,title,created_at,updated_at,settings_json) VALUES(?,?,?,?,?)",(current()["id"],title,now,now,json.dumps(settings,ensure_ascii=False))); rid=cur.lastrowid; c.commit(); c.close()
  log(rid,current()["id"],"CREADA",title); return redirect(url_for("rendicion",rid=rid))
 return render_template("new.html")

@app.route("/rendicion/<int:rid>",methods=["GET","POST"])
def rendicion(rid):
 if (r:=require_login()): return r
 row=get_rendicion(rid)
 if request.method=="POST":
  action=request.form.get("action","upload")
  if action=="upload":
   count=0
   for f in request.files.getlist("files"):
    if not f or not f.filename: continue
    original=secure_filename(f.filename)
    if Path(original).suffix.lower()!=".pdf":
     flash(f"Se omitió {original}: solo se admiten PDF.","error"); continue
    stored=f"{uuid.uuid4().hex}.pdf"; dest=UPLOAD/stored; f.save(dest)
    if dest.stat().st_size<5 or dest.read_bytes()[:4]!=b"%PDF":
     dest.unlink(missing_ok=True); flash(f"Se omitió {original}: el archivo no parece un PDF válido.","error"); continue
    c=db(); c.execute("INSERT INTO documents(rendicion_id,original_name,stored_name,uploaded_at) VALUES(?,?,?,?)",(rid,original,stored,datetime.now().isoformat())); c.commit(); c.close(); count+=1
   c=db()
   if count:
    c.execute("UPDATE rendiciones SET status='Pendiente de procesamiento',records_json=NULL,result_file=NULL,error_text=NULL,updated_at=? WHERE id=?",(datetime.now().isoformat(),rid))
   c.commit(); c.close()
   log(rid,current()["id"],"CARGA",f"{count} PDF(s)"); flash(f"Se cargaron {count} comprobante(s). Ahora completa los datos y pulsa Procesar rendición.","ok")
  elif action=="process":
   settings=settings_for(row)
   for key in ["descripcion_objeto","detalle_gasto","ubicacion","numero_cuenta","importe_recibido","tecnico_nombre","tecnico_dni"]:
    settings[key]=request.form.get(key,"").strip()
   settings["numero_rendicion"]=str(settings.get("numero_rendicion") or siguiente_numero_rendicion(db()))
   settings["fecha_transferencia"]=normalizar_fecha_transferencia(request.form.get("fecha_transferencia",""))
   settings["fecha_firma"]=fecha_larga_hoy()
   settings["texto_rendicion"]=""
   settings["titulo_movilidad"]="PLANILLA DE MOVILIDAD {nombre}"
   settings["incluir_movilidad"]=request.form.get("incluir_movilidad")=="on"
   try:
    if not settings["tecnico_nombre"] or not settings["tecnico_dni"] or not settings["numero_cuenta"]:
     raise ValueError("Completa el nombre del técnico, DNI y número de cuenta antes de procesar.")
    if not settings["importe_recibido"]: raise ValueError("Ingresa el importe recibido.")
    importe_txt=settings["importe_recibido"].replace(",","").strip()
    if not importe_txt.isdigit():
     raise ValueError("El importe recibido debe ser un número entero, sin decimales.")
    settings["importe_recibido"]=str(int(importe_txt))
    docs=docs_for(rid)
    if not docs: raise ValueError("Primero carga al menos un comprobante PDF.")
    paths=[UPLOAD/d["stored_name"] for d in docs]
    output=EXPORT/f"rendicion_{rid}.xlsx"
    records=process_pdfs(paths,settings,output)
    c=db(); c.execute("UPDATE rendiciones SET status='Completada',settings_json=?,records_json=?,result_file=?,error_text=NULL,updated_at=? WHERE id=?",
      (json.dumps(settings,ensure_ascii=False),json.dumps(records,ensure_ascii=False),output.name,datetime.now().isoformat(),rid)); c.commit(); c.close()
    log(rid,current()["id"],"PROCESADA",f"{len(records)} comprobante(s); Excel generado.")
    flash(f"Procesamiento terminado: {len(records)} comprobante(s). Revisa los datos extraídos y descarga el Excel.","ok")
   except Exception as exc:
    c=db(); c.execute("UPDATE rendiciones SET status='Requiere revisión',settings_json=?,result_file=NULL,error_text=?,updated_at=? WHERE id=?",
      (json.dumps(settings,ensure_ascii=False),str(exc),datetime.now().isoformat(),rid)); c.commit(); c.close()
    log(rid,current()["id"],"ERROR DE PROCESAMIENTO",str(exc)); flash(str(exc),"error")
  return redirect(url_for("rendicion",rid=rid))
 docs=docs_for(rid)
 try: records=json.loads(row["records_json"] or "[]")
 except Exception: records=[]
 c=db(); hist=c.execute("SELECT h.*,u.username FROM history h JOIN users u ON u.id=h.user_id WHERE h.rendicion_id=? ORDER BY h.id DESC",(rid,)).fetchall(); c.close()
 return render_template("rendicion.html",r=row,docs=docs,hist=hist,settings=settings_for(row),records=records,fecha_firma=fecha_larga_hoy())

@app.route("/rendicion/<int:rid>/eliminar",methods=["POST"])
def eliminar_rendicion(rid):
 if (r:=require_login()): return r
 row=get_rendicion(rid)
 docs=docs_for(rid)
 c=db(); c.execute("DELETE FROM history WHERE rendicion_id=?",(rid,)); c.execute("DELETE FROM documents WHERE rendicion_id=?",(rid,)); c.execute("DELETE FROM rendiciones WHERE id=?",(rid,)); c.commit(); c.close()
 for d in docs:
  (UPLOAD/d["stored_name"]).unlink(missing_ok=True)
 if row["result_file"]: (EXPORT/row["result_file"]).unlink(missing_ok=True)
 flash("La rendición y sus archivos asociados fueron eliminados.","ok")
 return redirect(url_for("dashboard"))

@app.route("/rendicion/<int:rid>/corregir/<int:idx>",methods=["POST"])
def corregir(rid,idx):
 if (r:=require_login()): return r
 row=get_rendicion(rid)
 try:
  records=json.loads(row["records_json"] or "[]")
  if idx<0 or idx>=len(records): abort(404)
  rec=records[idx]
  for k in ["fecha","ruc","tipo","serie","numero","proveedor","detalle","total"]:
   rec[k]=request.form.get(k,"").strip()
  datetime.strptime(rec["fecha"],"%Y-%m-%d"); rec["total"]=float(rec["total"].replace(",",""))
  settings=settings_for(row); output=EXPORT/f"rendicion_{rid}.xlsx"; make_excel(records,settings,output)
  c=db(); c.execute("UPDATE rendiciones SET records_json=?,result_file=?,status='Completada',updated_at=? WHERE id=?",(json.dumps(records,ensure_ascii=False),output.name,datetime.now().isoformat(),rid)); c.commit(); c.close()
  log(rid,current()["id"],"CORRECCIÓN",f"Comprobante {idx+1} actualizado y Excel regenerado."); flash("Dato corregido y Excel actualizado.","ok")
 except Exception as exc: flash(f"No se pudo guardar la corrección: {exc}","error")
 return redirect(url_for("rendicion",rid=rid))

@app.route("/rendicion/<int:rid>/exportar")
def exportar(rid):
 if (r:=require_login()): return r
 row=get_rendicion(rid)
 if not row["result_file"]: flash("Todavía no hay un Excel generado. Procesa la rendición primero.","error"); return redirect(url_for("rendicion",rid=rid))
 path=EXPORT/row["result_file"]
 if not path.exists(): abort(404)
 return send_file(path,as_attachment=True,download_name=f"RENDICION_{rid}_COMPLETADA.xlsx")

@app.route("/admin/usuarios",methods=["GET","POST"])
def usuarios():
 if (r:=require_login()): return r
 if current()["role"]!="admin": abort(403)
 if request.method=="POST":
  username=request.form.get("username","").strip(); password=request.form.get("password",""); role=request.form.get("role","user")
  if not username or len(password)<8: flash("Indica un usuario y una contraseña de al menos 8 caracteres.","error")
  else:
   try:
    c=db(); c.execute("INSERT INTO users(username,password,role,created_at) VALUES(?,?,?,?)",(username,generate_password_hash(password),role,datetime.now().isoformat())); c.commit(); c.close(); flash("Usuario creado.","ok")
   except sqlite3.IntegrityError: flash("Ese usuario ya existe.","error")
 c=db(); users=c.execute("SELECT id,username,role,created_at FROM users ORDER BY id").fetchall(); c.close()
 return render_template("users.html",users=users)

@app.route("/documento/<int:did>")
def documento(did):
 if (r:=require_login()): return r
 c=db(); d=c.execute("SELECT d.*,r.user_id FROM documents d JOIN rendiciones r ON r.id=d.rendicion_id WHERE d.id=?",(did,)).fetchone(); c.close()
 if not d: abort(404)
 if current()["role"]!="admin" and d["user_id"]!=current()["id"]: abort(403)
 path=UPLOAD/d["stored_name"]
 if not path.exists(): abort(404)
 return send_file(path,as_attachment=True,download_name=d["original_name"])

@app.errorhandler(413)
def too_large(e): return "El archivo supera el límite de 30 MB.",413

init_db()

if __name__=="__main__":
 app.run(host="0.0.0.0" if os.environ.get("RENDER") == "true" else "127.0.0.1",
         port=int(os.environ.get("PORT",5000)),debug=False)
