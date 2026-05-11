import firebase_admin
from firebase_admin import credentials, firestore
import os
import sys

# Set encoding to utf-8 for output
sys.stdout.reconfigure(encoding='utf-8')

json_file = 'credentials.json.json'

if not os.path.exists(json_file):
    print(f"File not found: {json_file}")
else:
    try:
        cred = credentials.Certificate(json_file)
        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred)
        
        db = firestore.client()
        print(f"Connected to Firebase (Project: {cred.project_id})")
        
        print("Testing Firestore fetch from 'users'...")
        docs = db.collection("users").limit(1).stream()
        found = False
        for doc in docs:
            print(f"Success! Found document ID: {doc.id}")
            found = True
        
        if not found:
            print("Connected, but 'users' collection is empty.")
            
    except Exception as e:
        print(f"Error occurred: {e}")
