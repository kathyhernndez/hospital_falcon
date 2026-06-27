import os
import sys
from app import app, db
from models import User

def create_super_admin():
    with app.app_context():
        email = sys.argv[1] if len(sys.argv) > 1 else "falcon@admin.ve"
        
        # Check if already exists
        existing = User.query.filter_by(correo=email).first()
        if existing:
            print(f"El usuario con correo {email} ya existe.")
            return

        password = sys.argv[2] if len(sys.argv) > 2 else "06falcon2026"
        nombre = sys.argv[3] if len(sys.argv) > 3 else "Super Admin"

        admin = User(
            nombre=nombre,
            correo=email,
            centro_medico_asociado="Ministerio de Salud / Central",
            is_verified_staff=True,
            role="admin",
            account_status="approved",
        )
        admin.set_password(password)

        db.session.add(admin)
        db.session.commit()
        print(f"Superusuario '{nombre}' ({email}) creado exitosamente.")

if __name__ == "__main__":
    create_super_admin()
