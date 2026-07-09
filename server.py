from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional, List
import pandas as pd
import os
import csv
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ============================================================
# 1. MODÈLES DE DONNÉES (Pydantic) - CORRIGÉ
# ============================================================

class Claim(BaseModel):
    Gestionnaire: Optional[str] = ""
    NumeroDossier: Optional[str] = ""  # ← CORRIGÉ : au lieu de "N° Dossier"
    Nibs: Optional[str] = ""           # ← CORRIGÉ : au lieu de "N°ibs"
    RefCIE: Optional[str] = ""         # ← CORRIGÉ : au lieu de "Ref CIE"
    Assure: Optional[str] = ""         # ← CORRIGÉ : au lieu de "Assuré"
    Police: Optional[str] = ""
    Garantie: Optional[str] = ""
    Immatriculation: Optional[str] = ""
    Expert: Optional[str] = ""
    DateSinistre: Optional[str] = ""   # ← CORRIGÉ : au lieu de "Date du sinistre"
    DateReceptionDeclaration: Optional[str] = ""
    DateOuvertureSinistre: Optional[str] = ""
    DateDesignationExpert: Optional[str] = ""
    DateReceptionDevis: Optional[str] = ""
    DateEnvoiDevis: Optional[str] = ""
    DateReceptionAccord: Optional[str] = ""
    DateEnvoiAccord: Optional[str] = ""
    DateExpertise: Optional[str] = ""
    DateReceptionFacture: Optional[str] = ""
    DateValidationRapport: Optional[str] = ""
    DateReglement: Optional[str] = ""
    ModeReglement: Optional[str] = ""
    MontantReglement: Optional[float] = 0.0
    TEL: Optional[str] = ""
    STATUT: Optional[str] = "En cours"
    OBSERVATION: Optional[str] = ""
    NumeroAssistance: Optional[str] = ""
    Beneficiaire: Optional[str] = ""

class ClaimSearch(BaseModel):
    NumeroDossier: Optional[str] = ""
    Assure: Optional[str] = ""
    Immatriculation: Optional[str] = ""
    STATUT: Optional[str] = ""
    Gestionnaire: Optional[str] = ""

class User(BaseModel):
    id: Optional[int] = None
    name: str
    email: str
    password: str
    role: str = "client"

class LoginData(BaseModel):
    email: str
    password: str

# ============================================================
# 2. GESTIONNAIRE DE DONNÉES
# ============================================================

class DataManager:
    def __init__(self, claims_file="claims.csv", users_file="users.csv"):
        self.claims_file = claims_file
        self.users_file = users_file
        self.columns = [
            'Gestionnaire', 'N° Dossier', 'N°ibs', 'Ref CIE', 'Assuré',
            'Police', 'Garantie', 'Immatriculation', 'Expert',
            'Date du sinistre', 'Date réception déclaration',
            'Date ouverture sinistre', 'Date désignation expert',
            'Date réception devis assuré', 'Date envoi devis à l\'expert',
            'Date réception Dernier Accord', 'Date Envoi Accord Assuré',
            'Date d\'expertise après réparation',
            'Date réception facture de réparation',
            'Date validation rapport d\'expertise', 'Date règlement',
            'Mode règlement', 'Montant règlement', 'TEL', 'STATUT',
            'OBSERVATION', 'N° Assistance', 'Bénéficiaire'
        ]
        self.gs_client = None
        self.gs_sheet = None
        self._init_files()
        self._init_google_sheets()

    # ===== GOOGLE SHEETS =====
    def _init_google_sheets(self):
        try:
            scope = ["https://spreadsheets.google.com/feeds",
                     "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
            self.gs_client = gspread.authorize(creds)
            self.gs_sheet = self.gs_client.open("sinistres auto").sheet1
            print("✅ Connecté à Google Sheets !")
        except Exception as e:
            print(f"⚠️ Google Sheets non disponible : {e}")
            self.gs_sheet = None

    def _save_to_google_sheets(self, claim_data):
        if not self.gs_sheet:
            return False
        try:
            row_data = []
            for col in self.columns:
                value = claim_data.get(col, '')
                if col == 'Montant règlement' and value:
                    value = float(value)
                row_data.append(value)
            self.gs_sheet.append_row(row_data)
            return True
        except Exception as e:
            print(f"⚠️ Erreur Google Sheets : {e}")
            return False

    def _sync_all_to_google_sheets(self):
        if not self.gs_sheet:
            return
        try:
            self.gs_sheet.clear()
            self.gs_sheet.append_row(self.columns)
            df = pd.read_csv(self.claims_file)
            for _, row in df.iterrows():
                row_data = [str(row.get(col, '')) for col in self.columns]
                self.gs_sheet.append_row(row_data)
            print("✅ Synchronisation Google Sheets réussie !")
        except Exception as e:
            print(f"⚠️ Erreur synchronisation : {e}")

    # ===== FICHIERS CSV =====
    def _init_files(self):
        if not os.path.exists(self.claims_file):
            df = pd.DataFrame(columns=self.columns)
            sample_data = {
                'Gestionnaire': ['DUPONT', 'MARTIN', 'BERNARD'],
                'N° Dossier': ['2024-001', '2024-002', '2024-003'],
                'N°ibs': ['IBS001', 'IBS002', 'IBS003'],
                'Ref CIE': ['CIE001', 'CIE002', 'CIE003'],
                'Assuré': ['Jean Dupont', 'Marie Martin', 'Pierre Bernard'],
                'Police': ['POL001', 'POL002', 'POL003'],
                'Garantie': ['Tous risques', 'Tiers', 'Tous risques'],
                'Immatriculation': ['AB-123-CD', 'EF-456-GH', 'IJ-789-KL'],
                'Expert': ['Expert1', 'Expert2', 'Expert1'],
                'Date du sinistre': ['2024-01-15', '2024-02-20', '2024-03-10'],
                'Date réception déclaration': ['2024-01-16', '2024-02-21', '2024-03-11'],
                'Date ouverture sinistre': ['2024-01-17', '2024-02-22', '2024-03-12'],
                'Date désignation expert': ['2024-01-18', '2024-02-23', '2024-03-13'],
                'Date réception devis assuré': ['2024-01-20', '2024-02-25', '2024-03-15'],
                'Date envoi devis à l\'expert': ['2024-01-21', '2024-02-26', '2024-03-16'],
                'Date réception Dernier Accord': ['2024-01-25', '2024-03-01', '2024-03-20'],
                'Date Envoi Accord Assuré': ['2024-01-26', '2024-03-02', '2024-03-21'],
                'Date d\'expertise après réparation': ['2024-01-28', '2024-03-04', '2024-03-23'],
                'Date réception facture de réparation': ['2024-01-30', '2024-03-06', '2024-03-25'],
                'Date validation rapport d\'expertise': ['2024-02-01', '2024-03-08', '2024-03-27'],
                'Date règlement': ['2024-02-05', '2024-03-10', '2024-03-30'],
                'Mode règlement': ['Virement', 'Chèque', 'Virement'],
                'Montant règlement': [1500.00, 850.00, 2200.00],
                'TEL': ['0123456789', '0987654321', '0567891234'],
                'STATUT': ['En cours', 'Clôturé', 'En attente'],
                'OBSERVATION': ['En attente de pièces', 'Dossier complet', 'Expertise en cours'],
                'N° Assistance': ['ASS001', 'ASS002', 'ASS003'],
                'Bénéficiaire': ['Jean Dupont', 'Marie Martin', 'Pierre Bernard']
            }
            df = pd.DataFrame(sample_data)
            df.to_csv(self.claims_file, index=False)

        if not os.path.exists(self.users_file):
            users_data = [
                {'id': 1, 'name': 'Jean Dupont', 'email': 'admin@sinistres.com', 'password': 'admin123', 'role': 'admin'},
                {'id': 2, 'name': 'Marie Martin', 'email': 'gestionnaire@sinistres.com', 'password': 'gest123', 'role': 'gestionnaire'},
                {'id': 3, 'name': 'Pierre Client', 'email': 'client@sinistres.com', 'password': 'client123', 'role': 'client'}
            ]
            df = pd.DataFrame(users_data)
            df.to_csv(self.users_file, index=False)

    # ===== CLAIMS =====
    def get_all_claims(self):
        df = pd.read_csv(self.claims_file)
        return df.to_dict('records')

    def get_claim_by_dossier(self, dossier):
        df = pd.read_csv(self.claims_file)
        result = df[df['N° Dossier'] == dossier]
        if result.empty:
            return None
        return result.to_dict('records')[0]

    def add_claim(self, claim_data):
        df = pd.read_csv(self.claims_file)
        new_row = pd.DataFrame([claim_data])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_csv(self.claims_file, index=False)
        self._save_to_google_sheets(claim_data)
        return True

    def update_claim(self, dossier, claim_data):
        df = pd.read_csv(self.claims_file)
        if dossier not in df['N° Dossier'].values:
            return False
        for key, value in claim_data.items():
            if key in df.columns:
                df.loc[df['N° Dossier'] == dossier, key] = value
        df.to_csv(self.claims_file, index=False)
        self._sync_all_to_google_sheets()
        return True

    def delete_claim(self, dossier):
        df = pd.read_csv(self.claims_file)
        df = df[df['N° Dossier'] != dossier]
        df.to_csv(self.claims_file, index=False)
        self._sync_all_to_google_sheets()
        return True

    def search_claims(self, criteria):
        df = pd.read_csv(self.claims_file)
        result = df
        for key, value in criteria.items():
            if value and key in df.columns:
                result = result[result[key].astype(str).str.contains(str(value), case=False, na=False)]
        return result.to_dict('records')

    def get_stats(self):
        df = pd.read_csv(self.claims_file)
        stats = {
            'total': len(df),
            'status_counts': df['STATUT'].value_counts().to_dict() if 'STATUT' in df.columns else {},
            'total_amount': float(df['Montant règlement'].sum()) if 'Montant règlement' in df.columns else 0
        }
        return stats

    # ===== USERS =====
    def get_all_users(self):
        df = pd.read_csv(self.users_file)
        return df.to_dict('records')

    def get_user_by_email(self, email):
        df = pd.read_csv(self.users_file)
        result = df[df['email'] == email]
        if result.empty:
            return None
        return result.to_dict('records')[0]

    def get_user_by_id(self, user_id):
        df = pd.read_csv(self.users_file)
        result = df[df['id'] == user_id]
        if result.empty:
            return None
        return result.to_dict('records')[0]

    def add_user(self, user_data):
        df = pd.read_csv(self.users_file)
        max_id = df['id'].max() if not df.empty else 0
        user_data['id'] = max_id + 1
        new_row = pd.DataFrame([user_data])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_csv(self.users_file, index=False)
        return user_data

    def update_user(self, user_id, user_data):
        df = pd.read_csv(self.users_file)
        if user_id not in df['id'].values:
            return False
        for key, value in user_data.items():
            if key in df.columns:
                df.loc[df['id'] == user_id, key] = value
        df.to_csv(self.users_file, index=False)
        return True

    def delete_user(self, user_id):
        df = pd.read_csv(self.users_file)
        df = df[df['id'] != user_id]
        df.to_csv(self.users_file, index=False)
        return True

# ============================================================
# 3. CRÉATION DE L'API
# ============================================================

app = FastAPI(title="API Gestion des Sinistres - Google Sheets")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

data_manager = DataManager()
data_manager._sync_all_to_google_sheets()

# ============================================================
# 4. ENDPOINTS - CLAIMS
# ============================================================

@app.get("/")
async def root():
    return {
        "message": "API Gestion des Sinistres",
        "status": "online",
        "google_sheets": data_manager.gs_sheet is not None
    }

@app.get("/api/claims")
async def get_claims():
    return data_manager.get_all_claims()

@app.get("/api/claims/{dossier}")
async def get_claim(dossier: str):
    claim = data_manager.get_claim_by_dossier(dossier)
    if not claim:
        raise HTTPException(status_code=404, detail="Sinistre non trouvé")
    return claim

@app.post("/api/claims")
async def add_claim(claim: Claim):
    # Convertir Claim en dictionnaire avec les bons noms de colonnes
    claim_data = {
        'Gestionnaire': claim.Gestionnaire,
        'N° Dossier': claim.NumeroDossier,
        'N°ibs': claim.Nibs,
        'Ref CIE': claim.RefCIE,
        'Assuré': claim.Assure,
        'Police': claim.Police,
        'Garantie': claim.Garantie,
        'Immatriculation': claim.Immatriculation,
        'Expert': claim.Expert,
        'Date du sinistre': claim.DateSinistre,
        'Date réception déclaration': claim.DateReceptionDeclaration,
        'Date ouverture sinistre': claim.DateOuvertureSinistre,
        'Date désignation expert': claim.DateDesignationExpert,
        'Date réception devis assuré': claim.DateReceptionDevis,
        'Date envoi devis à l\'expert': claim.DateEnvoiDevis,
        'Date réception Dernier Accord': claim.DateReceptionAccord,
        'Date Envoi Accord Assuré': claim.DateEnvoiAccord,
        'Date d\'expertise après réparation': claim.DateExpertise,
        'Date réception facture de réparation': claim.DateReceptionFacture,
        'Date validation rapport d\'expertise': claim.DateValidationRapport,
        'Date règlement': claim.DateReglement,
        'Mode règlement': claim.ModeReglement,
        'Montant règlement': claim.MontantReglement,
        'TEL': claim.TEL,
        'STATUT': claim.STATUT,
        'OBSERVATION': claim.OBSERVATION,
        'N° Assistance': claim.NumeroAssistance,
        'Bénéficiaire': claim.Beneficiaire
    }

    existing = data_manager.get_claim_by_dossier(claim.NumeroDossier)
    if existing:
        raise HTTPException(status_code=400, detail="Ce numéro de dossier existe déjà")
    data_manager.add_claim(claim_data)
    return {"message": "Sinistre ajouté avec succès"}

@app.put("/api/claims/{dossier}")
async def update_claim(dossier: str, claim: Claim):
    existing = data_manager.get_claim_by_dossier(dossier)
    if not existing:
        raise HTTPException(status_code=404, detail="Sinistre non trouvé")

    claim_data = {
        'Gestionnaire': claim.Gestionnaire,
        'N° Dossier': claim.NumeroDossier,
        'N°ibs': claim.Nibs,
        'Ref CIE': claim.RefCIE,
        'Assuré': claim.Assure,
        'Police': claim.Police,
        'Garantie': claim.Garantie,
        'Immatriculation': claim.Immatriculation,
        'Expert': claim.Expert,
        'Date du sinistre': claim.DateSinistre,
        'Date réception déclaration': claim.DateReceptionDeclaration,
        'Date ouverture sinistre': claim.DateOuvertureSinistre,
        'Date désignation expert': claim.DateDesignationExpert,
        'Date réception devis assuré': claim.DateReceptionDevis,
        'Date envoi devis à l\'expert': claim.DateEnvoiDevis,
        'Date réception Dernier Accord': claim.DateReceptionAccord,
        'Date Envoi Accord Assuré': claim.DateEnvoiAccord,
        'Date d\'expertise après réparation': claim.DateExpertise,
        'Date réception facture de réparation': claim.DateReceptionFacture,
        'Date validation rapport d\'expertise': claim.DateValidationRapport,
        'Date règlement': claim.DateReglement,
        'Mode règlement': claim.ModeReglement,
        'Montant règlement': claim.MontantReglement,
        'TEL': claim.TEL,
        'STATUT': claim.STATUT,
        'OBSERVATION': claim.OBSERVATION,
        'N° Assistance': claim.NumeroAssistance,
        'Bénéficiaire': claim.Beneficiaire
    }

    data_manager.update_claim(dossier, claim_data)
    return {"message": "Sinistre mis à jour avec succès"}

@app.delete("/api/claims/{dossier}")
async def delete_claim(dossier: str):
    existing = data_manager.get_claim_by_dossier(dossier)
    if not existing:
        raise HTTPException(status_code=404, detail="Sinistre non trouvé")
    data_manager.delete_claim(dossier)
    return {"message": "Sinistre supprimé avec succès"}

@app.post("/api/search")
async def search_claims(criteria: ClaimSearch):
    filters = {k: v for k, v in criteria.dict().items() if v and v != ""}
    if not filters:
        raise HTTPException(status_code=400, detail="Au moins un critère de recherche est requis")
    return data_manager.search_claims(filters)

@app.get("/api/statistics")
async def get_statistics():
    return data_manager.get_stats()

# ============================================================
# 5. ENDPOINTS - USERS
# ============================================================

@app.get("/api/users")
async def get_users():
    users = data_manager.get_all_users()
    for u in users:
        u.pop('password', None)
    return users

@app.post("/api/login")
async def login(login_data: LoginData):
    user = data_manager.get_user_by_email(login_data.email)
    if not user:
        raise HTTPException(status_code=401, detail="Email ou mot de passe incorrect")
    if user['password'] != login_data.password:
        raise HTTPException(status_code=401, detail="Email ou mot de passe incorrect")
    user.pop('password', None)
    return user

@app.post("/api/register")
async def register(user: User):
    existing = data_manager.get_user_by_email(user.email)
    if existing:
        raise HTTPException(status_code=400, detail="Cet email est déjà utilisé")
    user_data = user.dict()
    user_data.pop('id', None)
    new_user = data_manager.add_user(user_data)
    new_user.pop('password', None)
    return {"message": "Compte créé avec succès", "user": new_user}

@app.put("/api/users/{user_id}")
async def update_user(user_id: int, user: User):
    existing = data_manager.get_user_by_id(user_id)
    if not existing:
        raise HTTPException(status_code=404, detail="Utilisateur non trouvé")
    data_manager.update_user(user_id, user.dict())
    return {"message": "Utilisateur mis à jour"}

@app.delete("/api/users/{user_id}")
async def delete_user(user_id: int):
    existing = data_manager.get_user_by_id(user_id)
    if not existing:
        raise HTTPException(status_code=404, detail="Utilisateur non trouvé")
    data_manager.delete_user(user_id)
    return {"message": "Utilisateur supprimé"}

@app.post("/api/reset-password")
async def reset_password(email: str, new_password: str):
    user = data_manager.get_user_by_email(email)
    if not user:
        raise HTTPException(status_code=404, detail="Utilisateur non trouvé")
    data_manager.update_user(user['id'], {'password': new_password})
    return {"message": "Mot de passe réinitialisé avec succès"}

# ============================================================
# 6. LANCEMENT
# ============================================================

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8001)