import os
from sqlalchemy import text
from app import app, db

with app.app_context():
    # 1. Drop existing 'registros' to avoid FK constraints issues during alteration
    print("Dropping 'registros' table to recreate with FK...")
    db.session.execute(text("DROP TABLE IF EXISTS registros"))
    
    # Drop health_centers if it exists just in case
    db.session.execute(text("DROP TABLE IF EXISTS health_centers"))
    
    # 2. Re-create all missing tables (this will create health_centers and records)
    db.create_all()
    print("Database schema updated successfully for Falcón architecture!")
