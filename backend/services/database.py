import os
from datetime import datetime
from bson import ObjectId
from pymongo import MongoClient
import bcrypt
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

MONGO_URI = os.getenv("MONGO_URI") or os.getenv("MONGODB_URI") or "mongodb://localhost:27017"
DB_NAME = os.getenv("DB_NAME", "slide_generator")

_client = None

def get_db():
    global _client
    if _client is None:
        try:
            print(f"[DATABASE] Connecting to MongoDB at: {MONGO_URI}")
            _client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=3000)
            # Trigger serverSelectionTimeoutMS check to verify connection is alive
            _client.server_info()
        except Exception as e:
            print(f"[DATABASE] Error: Could not connect to MongoDB: {e}")
            raise e
    return _client[DB_NAME]

def hash_password(password: str) -> str:
    # Hash password using bcrypt
    salt = bcrypt.gensalt()
    hashed = bcrypt.hashpw(password.encode('utf-8'), salt)
    return hashed.decode('utf-8')

def verify_password(password: str, hashed_val: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_val.encode('utf-8'))
    except Exception as e:
        print(f"[DATABASE] Verification error: {e}")
        return False

def create_user(username: str, password_plain: str) -> dict:
    db = get_db()
    # Check if username exists
    existing = db.users.find_one({"username": username})
    if existing:
        raise ValueError("Username already exists")
    
    password_hash = hash_password(password_plain)
    user_doc = {
        "username": username,
        "password_hash": password_hash,
        "created_at": datetime.utcnow()
    }
    result = db.users.insert_one(user_doc)
    user_doc["_id"] = str(result.inserted_id)
    return user_doc

def get_user_by_username(username: str) -> dict:
    db = get_db()
    user = db.users.find_one({"username": username})
    if user:
        user["_id"] = str(user["_id"])
    return user

def log_auth_event(username: str, event_type: str, user_id: str = None):
    try:
        db = get_db()
        event_doc = {
            "username": username,
            "user_id": user_id,
            "event_type": event_type, # "login" or "logout"
            "timestamp": datetime.utcnow()
        }
        db.auth_events.insert_one(event_doc)
        print(f"[DATABASE] Logged auth event: {username} - {event_type}")
    except Exception as e:
        print(f"[DATABASE] Warning: Failed to log auth event: {e}")

def save_deck(user_id: str, title: str, slides: list, deck_id: str = None, tone: str = None) -> str:
    db = get_db()
    now = datetime.utcnow()
    
    try:
        u_oid = ObjectId(user_id)
    except Exception as e:
        print(f"[DATABASE] Invalid user_id '{user_id}': {e}")
        return None

    # Check if user exists
    user = db.users.find_one({"_id": u_oid})
    if not user:
        print(f"[DATABASE] User {user_id} not found")
        return None

    # Retrieve existing decks list or initialize it
    decks = user.get("decks", [])
    if not isinstance(decks, list):
        decks = []
    
    # If deck_id is provided, try to update it
    updated = False
    if deck_id:
        for deck in decks:
            if deck.get("id") == deck_id:
                deck["title"] = title
                deck["slides"] = slides
                deck["updated_at"] = now
                if tone:
                    deck["tone"] = tone
                updated = True
                break
                
    if not updated:
        # Create a new deck id
        new_deck_id = deck_id if deck_id else str(ObjectId())
        new_deck = {
            "id": new_deck_id,
            "title": title,
            "slides": slides,
            "created_at": now,
            "updated_at": now
        }
        if tone:
            new_deck["tone"] = tone
        decks.append(new_deck)
        deck_id = new_deck_id

    # Save the updated decks array back into the user document
    db.users.update_one(
        {"_id": u_oid},
        {"$set": {"decks": decks}}
    )
    print(f"[DATABASE] Saved deck {deck_id} inside user {user_id}'s document")
    return deck_id

def get_user_decks(user_id: str) -> list:
    db = get_db()
    try:
        u_oid = ObjectId(user_id)
        user = db.users.find_one({"_id": u_oid})
        if not user:
            return []
        
        decks = user.get("decks", [])
        if not isinstance(decks, list):
            return []
            
        # Sort decks by updated_at (descending)
        decks.sort(key=lambda d: d.get("updated_at", d.get("created_at", datetime.min)), reverse=True)
        
        result = []
        for d in decks:
            result.append({
                "id": d.get("id"),
                "title": d.get("title", "Untitled Presentation"),
                "updated_at": d.get("updated_at", d.get("created_at")).isoformat() if d.get("updated_at") else None
            })
        return result
    except Exception as e:
        print(f"[DATABASE] Error getting user decks: {e}")
        return []

def get_deck_by_id(deck_id: str, user_id: str) -> dict:
    db = get_db()
    try:
        u_oid = ObjectId(user_id)
        user = db.users.find_one({"_id": u_oid})
        if not user:
            return None
        
        decks = user.get("decks", [])
        if not isinstance(decks, list):
            return None
            
        for d in decks:
            if d.get("id") == deck_id:
                return {
                    "id": d.get("id"),
                    "title": d.get("title"),
                    "slides": d.get("slides", []),
                    "tone": d.get("tone")
                }
    except Exception as e:
        print(f"[DATABASE] Error fetching deck {deck_id} for user {user_id}: {e}")
    return None
