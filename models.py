"""
Modelos de base de datos para la aplicación de registro de personas
en centros de salud de Venezuela.

Incluye:
- Registro: persona ubicada en un centro de salud.
- User: personal de salud con acceso al panel de carga masiva.
"""

from datetime import datetime, timezone

from flask_login import UserMixin
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


# ---------------------------------------------------------------------------
# Modelo: Usuario (Personal de Salud)
# ---------------------------------------------------------------------------
class User(UserMixin, db.Model):
    """Modelo para personal de salud que accede al panel administrativo."""

    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    nombre = db.Column(db.String(200), nullable=False)
    correo = db.Column(db.String(200), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    centro_medico_asociado = db.Column(db.String(300), nullable=False)
    is_verified_staff = db.Column(db.Boolean, default=True)
    
    # Nuevos campos para permisos y validacion
    role = db.Column(db.String(50), default='health_staff')
    account_status = db.Column(db.String(50), default='pending')
    id_card_path = db.Column(db.String(300), nullable=True)
    failed_login_attempts = db.Column(db.Integer, default=0)
    lockout_until = db.Column(db.DateTime, nullable=True)

    fecha_registro = db.Column(
        db.DateTime, default=lambda: datetime.now(timezone.utc)
    )

    # --- Seguridad de contraseñas ---
    def set_password(self, password: str) -> None:
        """Hashea y almacena la contraseña de forma segura."""
        self.password_hash = generate_password_hash(password, method="scrypt")

    def check_password(self, password: str) -> bool:
        """Verifica la contraseña contra el hash almacenado."""
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f"<User {self.id}: {self.nombre} — {self.correo}>"


# ---------------------------------------------------------------------------
# Modelo: Centro de Salud
# ---------------------------------------------------------------------------
class HealthCenter(db.Model):
    """Modelo para centros de salud, enfocado en Falcón."""

    __tablename__ = "health_centers"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    nombre = db.Column(db.String(300), nullable=False)
    tipo = db.Column(db.String(100), nullable=False) # e.g. Hospital, CDI, Ambulatorio
    latitud = db.Column(db.Float, nullable=True)
    longitud = db.Column(db.Float, nullable=True)
    is_approved = db.Column(db.Boolean, default=False)
    
    # Relación uno a muchos con Registros
    registros = db.relationship('Registro', backref='health_center', lazy=True)

    def to_dict(self):
        return {
            "id": self.id,
            "nombre": self.nombre,
            "tipo": self.tipo,
            "latitud": self.latitud,
            "longitud": self.longitud,
            "is_approved": self.is_approved
        }

    def __repr__(self):
        return f"<HealthCenter {self.id}: {self.nombre}>"


# ---------------------------------------------------------------------------
# Modelo: Registro (Persona en centro de salud)
# ---------------------------------------------------------------------------
class Registro(db.Model):
    """Modelo para registrar a una persona ubicada en un centro de salud."""

    __tablename__ = "registros"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)

    # --- Relación con el centro de salud ---
    health_center_id = db.Column(db.Integer, db.ForeignKey('health_centers.id'), nullable=False)

    # --- Campos obligatorios ---
    nombre_apellido = db.Column(db.String(200), nullable=False)
    sector_ciudad = db.Column(db.String(200), nullable=False)
    quien_registra = db.Column(db.String(50), nullable=False)  # Familiar | Particular/Amigo | Yo mismo
    telefono_contacto = db.Column(db.String(30), nullable=False)

    # --- Campos opcionales ---
    cedula = db.Column(db.String(20), nullable=True)
    edad = db.Column(db.Integer, nullable=True)
    trabaja_centro = db.Column(db.Boolean, default=False)
    foto_path = db.Column(db.String(300), nullable=True)
    observaciones = db.Column(db.Text, nullable=True)

    # --- Metadatos ---
    fecha_registro = db.Column(
        db.DateTime, default=lambda: datetime.now(timezone.utc)
    )

    def __repr__(self):
        return f"<Registro {self.id}: {self.nombre_apellido}>"

    def to_dict(self, mask_phone=False):
        """Serializa el registro a diccionario JSON-friendly."""
        phone = self.telefono_contacto or ""
        if mask_phone and len(phone) >= 7:
            phone = phone[:4] + "-***-" + phone[-4:]

        return {
            "id": self.id,
            "nombre_apellido": self.nombre_apellido,
            "sector_ciudad": self.sector_ciudad,
            "centro_salud": self.health_center.nombre if self.health_center else "Centro Desconocido",
            "health_center_id": self.health_center_id,
            "quien_registra": self.quien_registra,
            "telefono_contacto": phone,
            "cedula": self.cedula or "",
            "edad": self.edad,
            "trabaja_centro": self.trabaja_centro,
            "foto_path": self.foto_path,
            "observaciones": self.observaciones or "",
            "fecha_registro": self.fecha_registro.strftime("%d/%m/%Y %H:%M")
            if self.fecha_registro
            else "",
        }
