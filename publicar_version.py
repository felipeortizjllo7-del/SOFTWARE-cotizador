# -*- coding: utf-8 -*-
"""
Publica una nueva version del Cotizador INNOBA:
  1. Sube el numero de version (V1.1 -> V1.2 ...).
  2. Compila el .exe (PyInstaller) y regenera el HTML.
  3. Construye el instalador (Inno Setup).
  4. Sube todo al repositorio de GitHub (commit + tag) y crea el Release
     con el instalador, actualizando version.json para la autoactualizacion.

Esquema resumido de 2 digitos: 1.0 -> 1.1 -> ... -> 1.9 -> 2.0

USO:
  python publicar_version.py            -> sube el siguiente (1.0 -> 1.1 ; 1.9 -> 2.0)
  python publicar_version.py --major    -> salta al siguiente entero (1.3 -> 2.0)
  python publicar_version.py 1.5        -> fija exactamente esa version
  (opcional)  --notas "texto de novedades"

Requiere haber corrido antes:  configurar_token.ps1
"""
import os, re, sys, json, ssl, subprocess, urllib.request, urllib.error

OWNER = "felipeortizjllo7-del"
REPO  = "SOFTWARE-cotizador"
PROJ  = os.path.dirname(os.path.abspath(__file__))
PY    = sys.executable
TOKEN_FILE = os.path.join(os.environ.get("APPDATA", PROJ), "CotizadorInnoba", "gh_token.txt")
CTX = ssl.create_default_context()

def _iscc():
    cands = [
        r"C:\Users\%s\AppData\Local\Programs\Inno Setup 6\ISCC.exe" % os.environ.get("USERNAME",""),
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    sys.exit("No se encontro ISCC.exe (Inno Setup). Instalalo con:\n"
             "  winget install JRSoftware.InnoSetup")

def leer_token():
    if not os.path.exists(TOKEN_FILE):
        sys.exit("No hay token configurado.\nEjecuta primero:  configurar_token.ps1")
    t = open(TOKEN_FILE, encoding="utf-8").read().strip()
    if not t:
        sys.exit("El token esta vacio. Vuelve a correr configurar_token.ps1")
    return t

def version_actual():
    src = open(os.path.join(PROJ, "cotizador_innoba.py"), encoding="utf-8").read()
    return re.search(r'VERSION\s*=\s*"([^"]+)"', src).group(1)

def bump(v, parte="minor"):
    """Esquema 2 digitos: sube el segundo digito; al pasar de 9 salta a (entero+1).0"""
    nums = (re.findall(r"\d+", v) + ["0", "0"])[:2]
    a, b = int(nums[0]), int(nums[1])
    if parte == "major":
        return f"{a + 1}.0"
    b += 1
    if b > 9:
        a, b = a + 1, 0
    return f"{a}.{b}"

def fijar_version(nv):
    fn = os.path.join(PROJ, "cotizador_innoba.py")
    s = open(fn, encoding="utf-8").read()
    s = re.sub(r'(VERSION\s*=\s*")[^"]+(")', r"\g<1>" + nv + r"\g<2>", s, count=1)
    open(fn, "w", encoding="utf-8").write(s)

def run(cmd, **kw):
    print("  »", cmd if isinstance(cmd, str) else " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=PROJ, **kw)

def git(*args):
    run(["git", "-C", PROJ] + list(args))

def api(url, token, data=None, method=None, ctype="application/json"):
    headers = {"Authorization": "Bearer " + token,
               "Accept": "application/vnd.github+json",
               "User-Agent": "publicar-cotizador",
               "X-GitHub-Api-Version": "2022-11-28"}
    body = None
    if data is not None:
        if ctype == "application/json":
            body = json.dumps(data).encode("utf-8")
        else:
            body = data
        headers["Content-Type"] = ctype
    req = urllib.request.Request(url, data=body, method=method or ("POST" if data else "GET"),
                                 headers=headers)
    with urllib.request.urlopen(req, context=CTX, timeout=300) as r:
        raw = r.read()
        return json.loads(raw) if raw else {}

# ---------------------------------------------------------------------------
def main():
    args = sys.argv[1:]
    notas = None
    if "--notas" in args:
        i = args.index("--notas"); notas = args[i + 1]; del args[i:i + 2]
    token = leer_token()
    actual = version_actual()
    if args and re.match(r"^\d+\.\d+$", args[0]):
        nueva = args[0]
    elif "--major" in args:
        nueva = bump(actual, "major")
    else:
        nueva = bump(actual, "minor")
    if notas is None:
        notas = f"Version {nueva} del Cotizador INNOBA."

    print(f"\n== Publicar version {actual}  ->  {nueva} ==")
    auto = ("--yes" in args) or ("-y" in args) or (not sys.stdin.isatty())
    if not auto:
        resp = input("Continuar? (s/n): ").strip().lower()
        if resp not in ("s", "si", "y", "yes"):
            sys.exit("Cancelado.")

    # 1. subir version en el codigo
    print("\n[1/6] Fijando numero de version...")
    fijar_version(nueva)

    # 2. compilar exe + regenerar html
    print("\n[2/6] Compilando .exe (PyInstaller)...")
    run([PY, "-m", "PyInstaller", "CotizadorInnoba.spec", "--noconfirm"])
    import shutil
    shutil.copyfile(os.path.join(PROJ, "dist", "CotizadorInnoba.exe"),
                    os.path.join(PROJ, "CotizadorInnoba.exe"))
    print("\n[2b] Regenerando HTML...")
    run([PY, "gen_html.py"])

    # 3. instalador
    print("\n[3/6] Construyendo instalador (Inno Setup)...")
    run([_iscc(), "/DMyAppVersion=" + nueva, "installer.iss"])
    instalador = os.path.join(PROJ, "installer_output", f"CotizadorInnoba-Setup-{nueva}.exe")
    if not os.path.exists(instalador):
        sys.exit("No se genero el instalador: " + instalador)

    # 4. version.json (para la autoactualizacion)
    print("\n[4/6] Actualizando version.json...")
    url_inst = (f"https://github.com/{OWNER}/{REPO}/releases/download/"
                f"v{nueva}/CotizadorInnoba-Setup-{nueva}.exe")
    with open(os.path.join(PROJ, "version.json"), "w", encoding="utf-8") as f:
        json.dump({"version": nueva, "installer": url_inst, "notas": notas},
                  f, ensure_ascii=False, indent=2)

    # 5. git commit + tag + push
    print("\n[5/6] Subiendo al repositorio...")
    if not os.path.exists(os.path.join(PROJ, ".git")):
        git("init")
        git("branch", "-M", "main")
    # remoto
    try:
        git("remote", "add", "origin", f"https://github.com/{OWNER}/{REPO}.git")
    except subprocess.CalledProcessError:
        pass  # ya existe
    git("add", "-A")
    try:
        git("-c", "user.name=INNOBA", "-c", "user.email=felipe@innobadmc.com",
            "commit", "-m", f"Version {nueva}")
    except subprocess.CalledProcessError:
        print("  (sin cambios nuevos para confirmar, se continua)")
    git("tag", "-f", f"v{nueva}")
    push_url = f"https://x-access-token:{token}@github.com/{OWNER}/{REPO}.git"
    def push(ref):
        print(f"  » git push origin {ref}")   # NO imprime el token
        # capturamos la salida para que el token de la URL nunca aparezca en pantalla
        r = subprocess.run(["git", "-C", PROJ, "push", push_url, ref, "--force"],
                           capture_output=True, text=True)
        if r.returncode != 0:
            err = (r.stderr or "").replace(token, "***")
            if "403" in err or "denied" in err.lower():
                sys.exit("\n[X] GitHub rechazo la subida (403): el token no tiene permiso de\n"
                         "    ESCRITURA sobre el repositorio.\n"
                         "    Crea el token con 'Repository permissions -> Contents: Read and write'\n"
                         "    (y 'Only select repositories -> SOFTWARE-cotizador'), pegalo con\n"
                         "    configurar_token.bat y vuelve a intentar.\n")
            sys.exit("[X] Error al subir al repositorio:\n" + err)
    push("HEAD:main")
    push(f"v{nueva}")

    # 6. Release + subir instalador
    print("\n[6/6] Creando Release y subiendo el instalador...")
    tag = f"v{nueva}"
    rel = None
    try:
        rel = api(f"https://api.github.com/repos/{OWNER}/{REPO}/releases", token,
                  data={"tag_name": tag, "name": tag, "body": notas,
                        "draft": False, "prerelease": False})
    except urllib.error.HTTPError as e:
        if e.code == 422:  # ya existe -> obtenerlo
            rel = api(f"https://api.github.com/repos/{OWNER}/{REPO}/releases/tags/{tag}", token)
        else:
            print(e.read().decode("utf-8", "ignore")); raise
    # borrar asset previo con el mismo nombre (si re-publicas la misma version)
    nombre = f"CotizadorInnoba-Setup-{nueva}.exe"
    for a in rel.get("assets", []):
        if a.get("name") == nombre:
            api(a["url"], token, method="DELETE")
    up = rel["upload_url"].split("{")[0]
    with open(instalador, "rb") as f:
        data = f.read()
    api(up + "?name=" + nombre, token, data=data, method="POST",
        ctype="application/octet-stream")

    print("\n==========================================================")
    print(f"  PUBLICADO v{nueva}")
    print(f"  Release:   https://github.com/{OWNER}/{REPO}/releases/tag/{tag}")
    print(f"  Instalador para instalar/actualizar: {url_inst}")
    print("  Los equipos con la app instalada veran el aviso al abrir.")
    print("==========================================================")

if __name__ == "__main__":
    main()
