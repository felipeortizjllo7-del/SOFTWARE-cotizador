# -*- coding: utf-8 -*-
"""Valida el token de GitHub guardado (sin mostrarlo) y su acceso al repositorio."""
import os, ssl, json, subprocess, urllib.request, urllib.error

OWNER = "felipeortizjllo7-del"
REPO  = "SOFTWARE-cotizador"
PROJ  = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(os.environ.get("APPDATA", ""), "CotizadorInnoba", "gh_token.txt")
CTX = ssl.create_default_context()

def probar_escritura(token):
    """Prueba REAL de escritura con 'git push --dry-run' (no sube nada).
       Devuelve (True, '') si puede escribir, (False, motivo) si no."""
    url = f"https://x-access-token:{token}@github.com/{OWNER}/{REPO}.git"
    if not os.path.exists(os.path.join(PROJ, ".git")):
        subprocess.run(["git", "-C", PROJ, "init"], capture_output=True, text=True)
    # aseguramos que haya al menos un commit para poder simular el push
    hay_commit = subprocess.run(["git", "-C", PROJ, "rev-parse", "HEAD"],
                                capture_output=True, text=True).returncode == 0
    if not hay_commit:
        return None, "sin commits locales para probar (se probara al publicar)"
    r = subprocess.run(["git", "-C", PROJ, "push", "--dry-run", url,
                        "HEAD:refs/heads/main", "--force"],
                       capture_output=True, text=True)
    err = (r.stderr or "").replace(token, "***")
    if r.returncode == 0:
        return True, ""
    if "403" in err or "denied" in err.lower():
        return False, "sin permiso de escritura (403)"
    return False, err.strip()[:200]

def api(url, token):
    req = urllib.request.Request(url, headers={
        "Authorization": "Bearer " + token,
        "Accept": "application/vnd.github+json",
        "User-Agent": "validar-cotizador",
        "X-GitHub-Api-Version": "2022-11-28"})
    r = urllib.request.urlopen(req, context=CTX, timeout=20)
    return r.status, json.loads(r.read() or "{}"), dict(r.headers)

def main():
    if not os.path.exists(TOKEN_FILE):
        print("[X]  No hay token. Ejecuta configurar_token.bat primero.")
        return
    token = open(TOKEN_FILE, encoding="utf-8").read().strip()
    if not token:
        print("[X]  El archivo del token esta vacio. Vuelve a correr configurar_token.bat.")
        return
    print(f"Token encontrado (longitud {len(token)} caracteres). Validando en GitHub...\n")

    # 1) identidad
    try:
        _, user, _ = api("https://api.github.com/user", token)
        print(f"[OK]  Token valido. Cuenta: {user.get('login')}")
    except urllib.error.HTTPError as e:
        print(f"[X]  Token invalido o vencido (HTTP {e.code}). Genera uno nuevo y vuelve a pegarlo.")
        return
    except Exception as e:
        print(f"[X]  No se pudo conectar a GitHub: {e}")
        return

    # 2) acceso al repositorio
    try:
        _, repo, _ = api(f"https://api.github.com/repos/{OWNER}/{REPO}", token)
        print(f"[OK]  Acceso al repositorio: {repo.get('full_name')}"
              f"  ({'PRIVADO' if repo.get('private') else 'PUBLICO'})")
    except urllib.error.HTTPError as e:
        print(f"[X]  El token no ve el repositorio {OWNER}/{REPO} (HTTP {e.code}).")
        print("   Crea el token con: Only select repositories -> SOFTWARE-cotizador.")
        return

    # 3) permiso REAL de escritura (prueba con push --dry-run, no sube nada)
    print("\nProbando permiso de escritura (sin subir nada)...")
    ok, motivo = probar_escritura(token)
    if ok is True:
        print("[OK]  Permiso de ESCRITURA confirmado (puede publicar versiones).")
        print("\n>>>  TODO LISTO. Ya se puede publicar.")
    elif ok is None:
        print(f"[!]   No se pudo probar del todo: {motivo}")
    else:
        print(f"[X]  El token NO puede escribir: {motivo}")
        print("   SOLUCION: crea el token en https://github.com/settings/tokens?type=beta")
        print("   con 'Repository permissions -> Contents: Read and write' sobre")
        print("   SOFTWARE-cotizador, y vuelve a pegarlo con configurar_token.bat.")

if __name__ == "__main__":
    main()
