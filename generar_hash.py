"""
Genera contraseñas hasheadas para usar en .streamlit/secrets.toml.
Ejecutar una sola vez y luego borrar o no subir al repo.

Uso:
    python generar_hash.py
"""
import streamlit_authenticator as stauth

# ── Cambia estas contraseñas por las que quieras usar ──
CONTRASENAS = [
    "Producto2026.",
]
# ────────────────────────────────────────────────────────

hashes = stauth.Hasher(CONTRASENAS).generate()

print("\nContraseñas hasheadas (copia en secrets.toml):\n")
for contrasena, hash_ in zip(CONTRASENAS, hashes):
    print(f"  Original : {contrasena}")
    print(f"  Hash     : {hash_}")
    print()
