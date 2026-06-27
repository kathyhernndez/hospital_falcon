from app import app, db
from models import HealthCenter

CENTROS_FALCON = [
    {
        "nombre": "Hospital Universitario Dr. Alfredo Van Grieken",
        "tipo": "Hospital",
        "latitud": 11.4160,
        "longitud": -69.6730,
        "is_approved": True
    },
    {
        "nombre": "Hospital Dr. Rafael Calles Sierra",
        "tipo": "Hospital",
        "latitud": 11.7145,
        "longitud": -70.1708,
        "is_approved": True
    },
    {
        "nombre": "Hospital Dr. Lino Arévalo",
        "tipo": "Hospital",
        "latitud": 10.7933,
        "longitud": -68.3242,
        "is_approved": True
    },
    {
        "nombre": "Hospital Dr. Carlos Diez del Ciervo",
        "tipo": "Hospital",
        "latitud": 11.7589,
        "longitud": -70.1802,
        "is_approved": True
    },
    {
        "nombre": "Hospital Emigdio Ríos",
        "tipo": "Hospital",
        "latitud": 10.8208,
        "longitud": -69.5447,
        "is_approved": True
    }
]

with app.app_context():
    print("Seeding Falcón health centers...")
    for center_data in CENTROS_FALCON:
        # Check if exists
        exists = HealthCenter.query.filter_by(nombre=center_data["nombre"]).first()
        if not exists:
            center = HealthCenter(**center_data)
            db.session.add(center)
            print(f"Added {center.nombre}")
    db.session.commit()
    print("Seed complete.")
