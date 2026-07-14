# -*- coding: utf-8 -*-
"""Valida el token de GitHub guardado (sin mostrarlo) y su acceso al repositorio."""
import os, ssl, json, urllib.request, urllib.error

OWNER = "felipeortizjllo7-del"
REPO  = "SOFTWARE-cotizador"
TOKEN_FILE = os.path.join(os.environ.get("APPDATA", ""), "CotizadorInnoba", "gh_token.txt")
CTX = ssl.create_default_context()

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

    # 2) acceso al repositorio y permiso de escritura
    try:
        _, repo, _ = api(f"https://api.github.com/repos/{OWNER}/{REPO}", token)
        perms = repo.get("permissions", {})
        puede_escribir = perms.get("push") or perms.get("admin") or perms.get("maintain")
        print(f"[OK]  Acceso al repositorio: {repo.get('full_name')}"
              f"  ({'PRIVADO' if repo.get('private') else 'PUBLICO'})")
        if puede_escribir:
            print("[OK]  Permiso de ESCRITURA confirmado (puede publicar versiones).")
            print("\n>>>  TODO LISTO. Ya puedes publicar con:  publicar_version.bat 1.1.0")
        else:
            print("[!]   El token accede al repo pero SIN permiso de escritura.")
            print("   Al crear el token, en 'Repository permissions' pon:")
            print("   Contents -> Read and write. Luego vuelve a pegarlo.")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            print(f"[X]  El token no tiene acceso al repositorio {OWNER}/{REPO}.")
            print("   Al crear el token elige: Only select repositories -> SOFTWARE-cotizador,")
            print("   y permiso Contents: Read and write.")
        else:
            print(f"[X]  Error consultando el repo (HTTP {e.code}).")

if __name__ == "__main__":
    main()
