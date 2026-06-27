import sqlite3

def migrate():
    conn = sqlite3.connect('hospitales.db')
    cursor = conn.cursor()
    
    # Try adding role
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN role VARCHAR(50) DEFAULT 'health_staff'")
        print("Added 'role' column.")
    except sqlite3.OperationalError as e:
        print(f"Error adding 'role': {e}")
        
    # Try adding account_status
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN account_status VARCHAR(50) DEFAULT 'pending'")
        print("Added 'account_status' column.")
    except sqlite3.OperationalError as e:
        print(f"Error adding 'account_status': {e}")
        
    # Try adding id_card_path
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN id_card_path VARCHAR(300)")
        print("Added 'id_card_path' column.")
    except sqlite3.OperationalError as e:
        print(f"Error adding 'id_card_path': {e}")
        
    # Try adding failed_login_attempts
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN failed_login_attempts INTEGER DEFAULT 0")
        print("Added 'failed_login_attempts' column.")
    except sqlite3.OperationalError as e:
        print(f"Error adding 'failed_login_attempts': {e}")

    # Try adding lockout_until
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN lockout_until TIMESTAMP")
        print("Added 'lockout_until' column.")
    except sqlite3.OperationalError as e:
        print(f"Error adding 'lockout_until': {e}")

    conn.commit()
    conn.close()
    print("Migration finished.")

if __name__ == '__main__':
    migrate()
