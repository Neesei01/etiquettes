# Installation nécessaire
# pip install googlemaps gspread google-auth folium pandas

# Importations initiales
import folium
import googlemaps
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
from folium.plugins import MarkerCluster
from zoneinfo import ZoneInfo
import csv
import os

# Déclarer la clef d'API Maps
CLEF_API = 'AIzaSyCkP4UBHMHm86s8u-edvF7x1al6-zWPqjE'

# Initialiser le client Google Maps
gmaps = googlemaps.Client(key=API_KEY)

# Configuration Google Sheets
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Configurer pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', None)

# Emplacement du fichier de commandes
CHEMIN_COMMANDES = 'commandes.csv'

# Déterminer si le circuit doit être une boucle
boucle = True

print("🚚 DÉMARRAGE DU CALCUL D'ITINÉRAIRE OPTIMISÉ")
print("=" * 60)

# Charger uniquement les adresses
try:
    commandes = pd.read_csv(
        CHEMIN_COMMANDES,
        usecols=[
            'Name',
            'Shipping Address1',
            'Shipping City', 
            'Shipping Zip',
            'Shipping Name',
            'Shipping Phone'
        ],
        dtype={
            'Name': str,
            'Shipping Address1': str,
            'Shipping City': str,
            'Shipping Zip': str,
            'Shipping Name': str,
            'Shipping Phone': str
        }
    ).dropna(subset=['Shipping Address1']).reset_index(drop=True)
    
    print(f"✅ Fichier CSV chargé: {len(commandes)} commandes trouvées")
    
except FileNotFoundError:
    print(f"❌ Erreur: Fichier {CHEMIN_COMMANDES} non trouvé")
    exit(1)
except Exception as e:
    print(f"❌ Erreur lors du chargement du CSV: {e}")
    exit(1)

# Formattage des codes postaux
commandes['Shipping Zip'] = commandes['Shipping Zip'].apply(
    lambda x: x[1:] if x.startswith('\'') else x
)

# Créer la liste des adresses et initialiser le départ/arrivée
adresses = [
    'Rue de Lyon Bâtiment E4, 94550 Chevilly-Larue',  # Point de fin (retour)
    'Rue de Lyon Bâtiment E4, 94550 Chevilly-Larue'   # Point de départ
]

# Créer la liste des adresses hors Île-de-France
hors_idf = []

# Créer une liste pour les étiquettes sur la carte
identites = [
    ('#0000', 'Dépôt', 'XX XX XX XX XX'),  # Retour final
    ('#0000', 'Dépôt', 'XX XX XX XX XX')   # Départ
]

# Créer une liste pour stocker l'ordre initial des commandes
ordre_commandes_initial = []

print("\n📍 ANALYSE DES ADRESSES")
print("-" * 30)

# Itérer sur les lignes pour les traiter
for _, row in commandes.iterrows():
    # Extraire les informations
    rue = row['Shipping Address1']
    ville = row['Shipping City']
    code_postal = row['Shipping Zip']
    numero_commande = row['Name']

    # Vérifier que le code postal est bien en Île-de-France
    ile_de_france = (
        code_postal.startswith('75') or  # Paris
        code_postal.startswith('77') or  # Seine-et-Marne
        code_postal.startswith('78') or  # Yvelines
        code_postal.startswith('91') or  # Essonne
        code_postal.startswith('92') or  # Hauts-de-Seine
        code_postal.startswith('93') or  # Seine-Saint-Denis
        code_postal.startswith('94') or  # Val-de-Marne
        code_postal.startswith('95')     # Val-d'Oise
    )

    # Traitement selon la localisation
    if ile_de_france:
        # Ajouter l'adresse au circuit
        adresses.append(f'{rue}, {ville}, {code_postal}')
        
        # Ajouter les informations d'identité
        identites.append((
            row['Name'],
            row['Shipping Name'], 
            row['Shipping Phone']
        ))
        
        # Ajouter à l'ordre initial (sans doublons)
        if numero_commande not in ordre_commandes_initial:
            ordre_commandes_initial.append(numero_commande)
    else:
        # Adresse hors région
        hors_idf.append(f'{rue}, {ville}, {code_postal}')

# Affichage des résultats de filtrage
if len(hors_idf) == 0:
    print('✅ Toutes les adresses sont en Île-de-France')
else:
    print(f'⚠️  Adresses hors IDF ignorées: {len(hors_idf)}')
    for adresse in hors_idf:
        print(f'   - {adresse}')

print(f'\n📦 Adresses de livraison conservées: {len(adresses) - 2}')
for i, adresse in enumerate(adresses[2:], 1):
    print(f'   {i}. {adresse}')

print(f'\n🔄 Configuration du circuit:')
if boucle:
    print(f'   Type: Boucle (retour au dépôt)')
    print(f'   Départ: {adresses[1]}')
    print(f'   Retour: {adresses[0]}')
else:
    print(f'   Type: Trajet simple')
    print(f'   Départ: {adresses[1]}')
    print(f'   Fin: {adresses[-1]}')

print("\n🌍 GÉOCODAGE DES ADRESSES")
print("-" * 30)

# Transformer les adresses en coordonnées Google Maps
try:
    geo_codes = []
    for i, adresse in enumerate(adresses):
        print(f"   Géocodage {i+1}/{len(adresses)}: {adresse[:50]}...")
        result = gmaps.geocode(adresse)
        if result:
            geo_codes.append(result)
        else:
            print(f"❌ Échec du géocodage pour: {adresse}")
            exit(1)
    
    print("✅ Géocodage terminé avec succès")
    
except Exception as e:
    print(f"❌ Erreur lors du géocodage: {e}")
    exit(1)

# Créer un dictionnaire des identités pour recherche rapide
identites_dict = {}
for geocode, identite in zip(geo_codes, identites):
    identites_dict[geocode[0]['formatted_address']] = identite

# Extraire les coordonnées
latitudes = [result[0]['geometry']['location']['lat'] for result in geo_codes]
longitudes = [result[0]['geometry']['location']['lng'] for result in geo_codes]

print("\n⏰ PLANIFICATION TEMPORELLE")
print("-" * 30)

# Calcul de l'heure d'exécution
timestamp = datetime.now(ZoneInfo('Europe/Paris'))
print(f"   Heure actuelle: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}")

# Trouver le prochain samedi 11h
jours_jusqu_samedi = (5 - timestamp.weekday()) % 7
prochain_samedi_11h = (timestamp + timedelta(days=jours_jusqu_samedi)).replace(
    hour=11, minute=0, second=0, microsecond=0
)

# Choisir le moment d'exécution
if timestamp >= prochain_samedi_11h:
    execution = 'now'
    print(f"   Calcul d'itinéraire: Maintenant")
else:
    execution = prochain_samedi_11h
    print(f"   Calcul d'itinéraire: {prochain_samedi_11h.strftime('%Y-%m-%d %H:%M:%S')}")

print("\n🗺️  CALCUL DE L'ITINÉRAIRE OPTIMISÉ")
print("-" * 40)

# Définir les points selon le type de circuit
if boucle:
    destination = adresses[0]  # Retour au dépôt
    waypoints = adresses[2:]   # Toutes les livraisons
else:
    destination = adresses[-1]     # Dernière adresse
    waypoints = adresses[2:-1]     # Adresses intermédiaires

print(f"   Origine: {adresses[1][:50]}...")
print(f"   Destination: {destination[:50]}...")
print(f"   Points intermédiaires: {len(waypoints)} adresses")
print("   Optimisation en cours...")

# Appel à l'API Google Directions avec optimisation
try:
    itineraire = gmaps.directions(
        origin=adresses[1],
        destination=destination,
        waypoints=waypoints,
        optimize_waypoints=True,
        mode='driving',
        departure_time=execution
    )
    
    if not itineraire:
        print("❌ Aucun itinéraire trouvé")
        exit(1)
        
    print("✅ Itinéraire optimisé calculé avec succès")
    
except Exception as e:
    print(f"❌ Erreur lors du calcul d'itinéraire: {e}")
    exit(1)

print("\n📊 EXTRACTION DE L'ORDRE OPTIMISÉ")
print("-" * 40)

# Extraire l'ordre optimisé des commandes
ordre_optimise = []

# Parcourir les segments de l'itinéraire dans l'ordre optimisé
for index, leg in enumerate(itineraire[0]['legs']):
    # Récupérer l'adresse de destination du segment
    adresse_destination = leg['end_address']
    
    # Retrouver l'identité correspondante
    if adresse_destination in identites_dict:
        identite = identites_dict[adresse_destination]
        numero_commande = identite[0]
        
        # Ajouter à l'ordre optimisé (exclure le retour final #0000)
        if numero_commande != '#0000':
            ordre_optimise.append(numero_commande)
            print(f"   {len(ordre_optimise)}. {numero_commande} - {identite[1]}")

print(f"\n✅ Ordre optimisé extrait: {len(ordre_optimise)} commandes")

print("\n💾 EXPORT DE L'ORDRE D'ITINÉRAIRE")
print("-" * 40)

# Chemin du fichier d'export
chemin_ordre = 'ordre_itineraire.csv'

# Exporter l'ordre optimisé
try:
    with open(chemin_ordre, 'w', newline='', encoding='utf-8') as fichier:
        writer = csv.writer(fichier)
        
        # Écrire l'en-tête
        writer.writerow(['Numero_Commande'])
        
        # Écrire chaque commande dans l'ordre optimisé
        for commande in ordre_optimise:
            writer.writerow([commande])
    
    print(f"✅ Fichier d'ordre exporté: {chemin_ordre}")
    print(f"📊 Nombre de commandes: {len(ordre_optimise)}")
    
    # Affichage de l'ordre (début et fin)
    if len(ordre_optimise) <= 10:
        print(f"📋 Ordre complet: {', '.join(ordre_optimise)}")
    else:
        premieres = ', '.join(ordre_optimise[:5])
        dernieres = ', '.join(ordre_optimise[-5:])
        print(f"📋 Premières commandes: {premieres}")
        print(f"📋 Dernières commandes: {dernieres}")
    
    print(f"\n💡 Utilisez ce fichier dans le générateur d'étiquettes")
    print(f"   pour trier automatiquement par ordre de livraison.")
    
except Exception as e:
    print(f"❌ Erreur lors de l'export: {e}")
    exit(1)

print("\n🗺️  GÉNÉRATION DE LA CARTE")
print("-" * 30)

# Créer une carte interactive
try:
    map_center = [latitudes[0], longitudes[0]]
    carte = folium.Map(location=map_center, zoom_start=12)

    # Ajouter un cluster de marqueurs
    marker_cluster = MarkerCluster().add_to(carte)

    # Ajouter les marqueurs pour chaque adresse
    for adresse, lat, lng, identite in zip(adresses, latitudes, longitudes, identites):
        folium.Marker(
            [lat, lng],
            popup=(
                f'Commande: {identite[0]}<br>'
                f'Adresse: {adresse}<br>'
                f'Client: {identite[1]}<br>'
                f'Téléphone: {identite[2]}'
            ),
            opacity=1,
        ).add_to(marker_cluster)

    print("✅ Marqueurs ajoutés à la carte")

    # Initialisation des métriques de trajet
    distances_segments = []
    durees_segments = []

    # Ajouter l'itinéraire avec un gradient de couleurs
    for index, leg in enumerate(itineraire[0]['legs']):
        # Ajouter les métriques
        durees_segments.append(leg['duration']['value'])
        distances_segments.append(leg['distance']['value'])

        # Calcul de la couleur (gradient)
        ratio = index / max(1, len(itineraire[0]['legs']) - 1)
        red = int(64 + 128 * ratio * (1 - ratio))
        green = int(64 + 128 * ratio)
        blue = int(64 + 128 * (1 - ratio))
        couleur = f'rgb({red}, {green}, {blue})'

        # Ajouter les segments de route
        for etape in leg['steps']:
            depart = etape['start_location']
            arrivee = etape['end_location']
            
            folium.PolyLine(
                locations=[
                    [depart['lat'], depart['lng']],
                    [arrivee['lat'], arrivee['lng']]
                ],
                color=couleur,
                weight=4,
                opacity=0.8
            ).add_to(carte)

    print("✅ Itinéraire tracé sur la carte")
    
    # Sauvegarder la carte
    nom_carte = 'itineraire_optimise.html'
    carte.save(nom_carte)
    print(f"✅ Carte sauvegardée: {nom_carte}")
    
except Exception as e:
    print(f"❌ Erreur lors de la génération de carte: {e}")

print("\n📈 MÉTRIQUES DU TRAJET")
print("-" * 30)

# Temps de pause par livraison (en secondes)
temps_pause_livraison = 600  # 10 minutes

# Calculs des métriques
distance_totale = round(sum(distances_segments) / 1000, 1)  # en km
duree_conduite = sum(durees_segments)  # en secondes
duree_pauses = temps_pause_livraison * (len(adresses) - 2)  # pauses livraisons
duree_totale = duree_conduite + duree_pauses

# Conversion en heures/minutes/secondes
heures = duree_totale // 3600
minutes = (duree_totale % 3600) // 60
secondes = duree_totale % 60

print(f"   📏 Distance totale: {distance_totale} km")
print(f"   🚗 Temps de conduite: {duree_conduite//3600}h {(duree_conduite%3600)//60}min")
print(f"   ⏸️  Temps de pauses: {duree_pauses//60} min ({len(adresses)-2} livraisons)")
print(f"   ⏱️  Durée totale estimée: {heures}h {minutes}min {secondes}sec")

print("\n📋 PLANNING DES LIVRAISONS")
print("-" * 40)

# Simulation des horaires de livraison
heure_depart = 10 * 3600  # 10h00 en secondes
heure_courante_min = heure_depart
heure_courante_max = heure_depart

# Préparation des données pour export
donnees_livraisons = []

print(f"   🕘 Heure de départ estimée: 10h00")
print(f"   📍 Estimations d'arrivée (ordre optimisé):")

for index, leg in enumerate(itineraire[0]['legs']):
    # Récupération des informations
    adresse_destination = leg['end_address']
    identite = identites_dict[adresse_destination]
    
    # Mise à jour des horaires
    heure_courante_min += leg['duration']['value'] + (temps_pause_livraison if index > 0 else 0)
    heure_courante_max += leg['duration']['value'] + (temps_pause_livraison if index > 0 else 0)
    
    # Conversion en format lisible
    h_min = (heure_courante_min // 3600) % 24
    m_min = (heure_courante_min % 3600) // 60
    h_max = (heure_courante_max // 3600) % 24
    m_max = (heure_courante_max % 3600) // 60
    
    horaire_min = f"{h_min}h{m_min:02d}"
    horaire_max = f"{h_max}h{m_max:02d}"
    
    # Affichage
    if identite[0] != '#0000':
        print(f"      {index+1:2d}. {identite[0]} - {identite[1][:25]:<25} | {horaire_min}-{horaire_max}")
    else:
        print(f"      {index+1:2d}. RETOUR AU DÉPÔT{'':<32} | {horaire_min}")
    
    # Ajout aux données d'export
    donnees_livraisons.append({
        'ordre': index + 1,
        'commande': identite[0],
        'client': identite[1],
        'telephone': identite[2],
        'adresse': adresse_destination,
        'horaire': f"{horaire_min}-{horaire_max}",
        'distance_km': round(leg['distance']['value'] / 1000, 1),
        'duree_min': round(leg['duration']['value'] / 60)
    })

print("\n📄 GÉNÉRATION DU RAPPORT")
print("-" * 30)

# Création d'un DataFrame pour l'export
df_rapport = pd.DataFrame(donnees_livraisons)

# Export CSV du rapport
nom_rapport = 'rapport_livraisons.csv'
try:
    df_rapport.to_csv(nom_rapport, index=False, encoding='utf-8')
    print(f"✅ Rapport détaillé exporté: {nom_rapport}")
except Exception as e:
    print(f"❌ Erreur export rapport: {e}")

print("\n🎯 RÉSUMÉ FINAL")
print("=" * 60)
print(f"📦 Commandes à livrer: {len(ordre_optimise)}")
print(f"📏 Distance totale: {distance_totale} km")
print(f"⏱️  Durée estimée: {heures}h{minutes}min")
print(f"🗺️  Itinéraire optimisé: Google Maps API")
print(f"📂 Fichiers générés:")
print(f"   • {chemin_ordre} (ordre pour étiquettes)")
print(f"   • {nom_rapport} (rapport détaillé)")
print(f"   • {nom_carte} (carte interactive)")

print(f"\n🔄 PROCHAINES ÉTAPES:")
print(f"1️⃣  Chargez {chemin_ordre} dans le générateur d'étiquettes")
print(f"2️⃣  Les étiquettes seront triées par ordre de livraison")
print(f"3️⃣  Première commande à livrer = première ligne")
print(f"4️⃣  Consultez {nom_carte} pour visualiser l'itinéraire")

print("\n✅ CALCUL D'ITINÉRAIRE TERMINÉ AVEC SUCCÈS")
print("=" * 60)

# ==========================================
# OPTIONNEL : INTÉGRATION GOOGLE SHEETS
# ==========================================

def integrer_google_sheets():
    """
    Fonction optionnelle pour intégrer les données dans Google Sheets
    Nécessite une configuration d'authentification préalable
    """
    try:
        # Configuration Google Sheets (décommentez selon votre méthode d'auth)
        
        # Option 1: Fichier de clés de service
        # creds = Credentials.from_service_account_file('chemin/vers/cles_service.json', scopes=SCOPES)
        
        # Option 2: Authentification OAuth (Google Colab)
        # from google.colab import auth
        # auth.authenticate_user()
        # from google.auth import default
        # creds, _ = default()
        
        # Initialiser le client gspread
        # gc = gspread.authorize(creds)
        
        # Nom du spreadsheet
        SHEET_NAME = "Tournées_Livraisons_Master"
        
        # Créer ou ouvrir le spreadsheet
        # spreadsheet = gc.open(SHEET_NAME)
        # worksheet = spreadsheet.sheet1
        
        # Préparer les données pour l'insertion
        timestamp_execution = datetime.now(ZoneInfo('Europe/Paris')).strftime("%Y-%m-%d %H:%M:%S")
        date_livraison = (datetime.now(ZoneInfo('Europe/Paris')) + timedelta(days=2)).strftime('%d/%m/%Y')
        
        # Insérer les données
        # for i, livraison in enumerate(donnees_livraisons):
        #     ligne = [
        #         timestamp_execution if i == 0 else '',
        #         'commandes.csv' if i == 0 else '',
        #         date_livraison,
        #         livraison['ordre'],
        #         livraison['commande'],
        #         livraison['client'],
        #         livraison['telephone'],
        #         livraison['adresse'],
        #         livraison['horaire'],
        #         livraison['distance_km'],
        #         livraison['duree_min'],
        #         f'{distance_totale} km' if i == 0 else '',
        #         f'{heures}h{minutes}min' if i == 0 else ''
        #     ]
        #     worksheet.append_row(ligne)
        
        print("\n📊 INTÉGRATION GOOGLE SHEETS")
        print("-" * 30)
        print("⚠️  Fonction désactivée - Configurez l'authentification Google")
        print("   Décommentez les lignes dans la fonction integrer_google_sheets()")
        print("   pour activer l'intégration automatique")
        
    except Exception as e:
        print(f"❌ Erreur Google Sheets: {e}")

# Appeler la fonction (actuellement désactivée)
integrer_google_sheets()

# ==========================================
# FONCTION DE VALIDATION
# ==========================================

def valider_resultats():
    """
    Validation des résultats et vérifications de cohérence
    """
    print(f"\n🔍 VALIDATION DES RÉSULTATS")
    print("-" * 30)
    
    # Vérifications de base
    verifications = []
    
    # 1. Cohérence du nombre de commandes
    nb_commandes_csv = len(commandes['Name'].unique())
    nb_commandes_optimise = len(ordre_optimise)
    nb_commandes_hors_idf = len(hors_idf)
    
    if nb_commandes_optimise + nb_commandes_hors_idf == nb_commandes_csv:
        verifications.append("✅ Nombre de commandes cohérent")
    else:
        verifications.append("❌ Incohérence dans le nombre de commandes")
    
    # 2. Fichiers générés
    if os.path.exists(chemin_ordre):
        verifications.append("✅ Fichier d'ordre créé")
    else:
        verifications.append("❌ Fichier d'ordre manquant")
    
    if os.path.exists(nom_rapport):
        verifications.append("✅ Rapport détaillé créé")
    else:
        verifications.append("❌ Rapport détaillé manquant")
    
    if os.path.exists(nom_carte):
        verifications.append("✅ Carte interactive créée")
    else:
        verifications.append("❌ Carte interactive manquante")
    
    # 3. Métriques réalistes
    if 10 <= distance_totale <= 500:  # Entre 10km et 500km semble raisonnable
        verifications.append("✅ Distance totale réaliste")
    else:
        verifications.append(f"⚠️  Distance totale à vérifier: {distance_totale}km")
    
    if 1 <= heures <= 12:  # Entre 1h et 12h semble raisonnable
        verifications.append("✅ Durée totale réaliste")
    else:
        verifications.append(f"⚠️  Durée totale à vérifier: {heures}h{minutes}min")
    
    # Affichage des résultats
    for verification in verifications:
        print(f"   {verification}")
    
    # Score de validation
    score_succes = len([v for v in verifications if v.startswith("✅")])
    score_total = len(verifications)
    
    print(f"\n📊 Score de validation: {score_succes}/{score_total}")
    
    if score_succes == score_total:
        print("🎉 Tous les contrôles sont validés !")
    elif score_succes >= score_total * 0.8:
        print("😊 La plupart des contrôles sont validés")
    else:
        print("⚠️  Plusieurs points nécessitent une vérification")

# Exécuter la validation
valider_resultats()

print(f"\n📚 INFORMATIONS TECHNIQUES")
print("-" * 30)
print(f"   🐍 Version Python requise: 3.8+")
print(f"   📦 Dépendances principales:")
print(f"      • googlemaps (pip install googlemaps)")
print(f"      • folium (pip install folium)")
print(f"      • pandas (pip install pandas)")
print(f"      • gspread (pip install gspread) [optionnel]")
print(f"   🔑 Configuration requise:")
print(f"      • Clé API Google Maps valide")
print(f"      • Fichier CSV avec colonnes: Name, Shipping Address1, Shipping City, etc.")
print(f"   💾 Espace disque: ~5MB pour les fichiers générés")

print(f"\n🆘 RÉSOLUTION DES PROBLÈMES COURANTS")
print("-" * 40)
print(f"   🔹 'Aucun itinéraire trouvé':")
print(f"      → Vérifiez la validité des adresses")
print(f"      → Contrôlez votre quota API Google Maps")
print(f"   🔹 'Erreur de géocodage':")
print(f"      → Vérifiez votre connexion internet")
print(f"      → Assurez-vous que l'API Geocoding est activée")
print(f"   🔹 'Fichier CSV non trouvé':")
print(f"      → Vérifiez le chemin: {CHEMIN_COMMANDES}")
print(f"      → Assurez-vous que le fichier existe")
print(f"   🔹 'Erreur d'export':")
print(f"      → Vérifiez les permissions d'écriture")
print(f"      → Fermez le fichier s'il est ouvert dans Excel")

print(f"\n📞 SUPPORT ET CONTACT")
print("-" * 20)
print(f"   📧 Pour signaler un bug ou demander de l'aide:")
print(f"      → Créez une issue sur le repository GitHub")
print(f"      → Incluez le message d'erreur complet")
print(f"      → Précisez votre système d'exploitation")
print(f"   📖 Documentation complète:")
print(f"      → README.md du projet")
print(f"      → Commentaires dans le code source")

print(f"\n🎯 SCRIPT TERMINÉ - PRÊT POUR LA PRODUCTION")
print("=" * 60)("💡 Utilisez ce fichier dans le générateur d'étiquettes pour trier automatiquement les étiquettes par ordre de livraison.")
    
except Exception as e:
    print(f"❌ Erreur lors de l'export: {e}")

print()

# ==========================================
# SUITE DU CODE ORIGINAL...
# ==========================================

# Créer une carte
map_center = [latitudes[0], longitudes[0]]
map_ = folium.Map(location=map_center, zoom_start=12)

# Marquer les adresses
marker_cluster = MarkerCluster().add_to(map_)
for adresse, lat, lng, id in zip(adresses, latitudes, longitudes, identites):
    folium.Marker(
        [lat, lng],
        popup=(
            f'Commande n°{id[0]}' + ', ' +
            f'{adresse}' + ', ' +
            f'Client {id[1]}' + ', ' +
            f'Téléphone {id[2]}'
        ),
        opacity=1,
    ).add_to(marker_cluster)

# Initialiser la distance et le temps des tronçons de trajet
distances_tronçons = []
temps_tronçons = []

# Ajouter l'itinéraire (tous les legs et steps)
for index, leg in enumerate(itineraire[0]['legs']):
    # Ajouter le temps total
    temps_tronçons.append(leg['duration']['value'])

    # Ajouter la distance totale
    distances_tronçons.append(leg['distance']['value'])

    # Petit passe-passe pour un gradient de couleurs
    ratio = index / (len(itineraire[0]['legs']) - 1)
    red = int(64 + 128 * ratio * (1 - ratio))
    green = int(64 + 128 * ratio)
    blue = int(64 + 128 * (1 - ratio))

    # Créer le code couleur
    color = f'rgb({red}, {green}, {blue})'

    # Ajouter les petits segments
    for etape in leg['steps']:
        depart_etape = etape['start_location']
        arrivee_etape = etape['end_location']
        folium.PolyLine(
            locations=[
                [depart_etape['lat'], depart_etape['lng']],
                [arrivee_etape['lat'], arrivee_etape['lng']]
            ],
            color=color,
            weight=4,
            opacity=1
        ).add_to(map_)

# Initialiser le temps de pause moyen par livraison
temps_livraison = 600

# Calculer la distance cumulée et le temps total (incluant les pauses)
distance_trajet = round(sum(distances_tronçons) / 1000, 1)
temps_trajet = sum(temps_tronçons) + temps_livraison * (len(adresses) - 2)

# Convertir le temps en heures, minutes et secondes
heures = temps_trajet // 3600
minutes = (temps_trajet % 3600) // 60
secondes = temps_trajet % 60

# Initier les temps minimaux (fourchette de 6h)
temps_min = 10 * 3600
temps_max = 16 * 3600

# Afficher la distance cumulée et le temps total
print(f'Distance totale estimée du trajet : {distance_trajet} km')
print(f'Durée totale estimée du trajet : {heures}h {minutes} min {secondes} sec')
print()

# Préparer les données pour Google Sheets
donnees_livraisons = []

# Afficher les estimations pour chaque commande et préparer les données
print('Estimation des temps d\'arrivée (dans l\'ordre optimisé) :')

# Itérer dans l'ordre des adresses OPTIMISÉ
for index, leg in enumerate(itineraire[0]['legs']):
    # Retouver la clef adresse du dictionnaire
    adresse = leg['end_address']

    # Retrouver l'identité associée
    id = identites_dict[adresse]

    # Afficher les informations sur la commande
    print(f'\tTrajet n°{index + 1}')
    print(f'\t\t- {"Commande " + id[0] if id[0][1] != "0" else "RETOUR FINAL"}')
    print(f'\t\t- {adresse}')
    print(f'\t\t- {"Client" if id[0][1] != "0" else "Fournisseur"} : {id[1]}')
    print(f'\t\t- Téléphone : {id[2]}')

    # Mettre à jour les temps mins et max
    temps_min += leg['duration']['value'] + temps_livraison * min(index, 1)
    temps_max += leg['duration']['value'] + temps_livraison * min(index, 1)

    # Calculer en heures, minutes et secondes
    heure_min = temps_min // 3600
    minute_min = round((temps_min % 3600) / 60)
    string_min = f'{heure_min % 24}h{0 if minute_min < 10 else ""}{minute_min}'

    heure_max = temps_max // 3600
    minute_max = round((temps_max % 3600) // 60)
    string_max = f'{heure_max % 24}h{0 if minute_max < 10 else ""}{minute_max}'

    print(f'\t\t- Arrivée prévue entre {string_min} et {string_max}.')
    print()

    # Ajouter les données à la liste pour Google Sheets
    donnees_livraisons.append([
        index + 1,  # Numéro de trajet
        id[0] if id[0][1] != "0" else "RETOUR FINAL",  # Numéro de commande
        id[1],  # Nom du client
        id[2],  # Téléphone
        adresse,  # Adresse complète
        f'{string_min} et {string_max}',  # Créneau d'arrivée
        round(leg['distance']['value'] / 1000, 1),  # Distance en km
        round(leg['duration']['value'] / 60)  # Durée en minutes
    ])

# Configuration du nom du Google Sheet (constant)
SHEET_NAME = "Tournées_Livraisons_Master"

# Fonction pour créer ou mettre à jour le Google Sheet
def gerer_google_sheet():
    try:
        # ATTENTION: Vous devez configurer l'authentification Google
        # Option 1: Utiliser un fichier de clés de service
        # creds = Credentials.from_service_account_file('path/to/your/service_account.json', scopes=SCOPES)

        # Option 2: Utiliser l'authentification OAuth (pour Google Colab)
        from google.colab import auth
        auth.authenticate_user()

        import gspread
        from google.auth import default
        creds, _ = default()

        # Initialiser le client gspread
        gc = gspread.authorize(creds)

        # Essayer d'ouvrir le spreadsheet existant
        try:
            spreadsheet = gc.open(SHEET_NAME)
            worksheet = spreadsheet.sheet1
            print(f"📋 Utilisation du Google Sheet existant: {SHEET_NAME}")

        except gspread.SpreadsheetNotFound:
            # Si le spreadsheet n'existe pas, le créer
            print(f"📋 Création d'un nouveau Google Sheet: {SHEET_NAME}")
            spreadsheet = gc.create(SHEET_NAME)
            worksheet = spreadsheet.sheet1

            # Définir les en-têtes pour le nouveau sheet
            headers = [
                'Date/Heure Calcul',
                'Fichier Source',
                'Date Livraison Prévue',
                'N° Trajet',
                'N° Commande',
                'Nom Client',
                'Téléphone',
                'Adresse',
                'Heure d\'arrivée prévue',
                'Distance (km)',
                'Durée trajet (min)',
                'Distance Totale Tournée',
                'Durée Totale Tournée'
            ]

            # Ajouter les en-têtes
            worksheet.append_row(headers)

            # Formater les en-têtes
            worksheet.format('A1:M1', {
                'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.8},
                'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
            })

            # Partager le document (lecture pour tous)
            spreadsheet.share('', perm_type='anyone', role='reader')

        # Ajouter une ligne de séparation pour cette nouvelle exécution
        timestamp_execution = datetime.now(ZoneInfo('Europe/Paris')).strftime("%Y-%m-%d %H:%M:%S")

        # Calculer la date de livraison (date actuelle + 2 jours) en format français
        date_livraison = datetime.now(ZoneInfo('Europe/Paris')) + timedelta(days=2)

        # Convertir en format français
        mois_francais = {
            1: 'Janvier', 2: 'Février', 3: 'Mars', 4: 'Avril',
            5: 'Mai', 6: 'Juin', 7: 'Juillet', 8: 'Août',
            9: 'Septembre', 10: 'Octobre', 11: 'Novembre', 12: 'Décembre'
        }

        date_livraison_fr = f"{date_livraison.day} {mois_francais[date_livraison.month]} {date_livraison.year}"

        # Obtenir le nom du fichier CSV utilisé
        nom_fichier = CHEMIN_COMMANDES.split('/')[-1]

        # Ajouter les données de cette exécution
        for i, row in enumerate(donnees_livraisons):
            # Préparer la ligne complète avec les métadonnées
            ligne_complete = [
                timestamp_execution if i == 0 else '',  # Date/heure seulement sur la première ligne
                nom_fichier if i == 0 else '',  # Nom du fichier seulement sur la première ligne
                date_livraison_fr,  # Date de livraison pour chaque client
                row[0],  # N° Trajet
                row[1],  # N° Commande
                row[2],  # Nom Client
                row[3],  # Téléphone
                row[4],  # Adresse
                row[5],  # Heure d'arrivée prévue
                row[6],  # Distance (km)
                row[7],  # Durée trajet (min)
                f'{distance_trajet} km' if i == 0 else '',  # Distance totale seulement sur la première ligne
                f'{heures}h{minutes}min' if i == 0 else ''  # Durée totale seulement sur la première ligne
            ]

            worksheet.append_row(ligne_complete)

        # Ajouter une ligne vide pour séparer les exécutions
        worksheet.append_row(['', '', '', '', '', '', '', '', '', '', '', '', ''])

        # Mettre en évidence la nouvelle section ajoutée
        derniere_ligne = len(worksheet.get_all_values())
        premiere_ligne_execution = derniere_ligne - len(donnees_livraisons) - 1

        # Colorer légèrement la nouvelle section
        worksheet.format(f'A{premiere_ligne_execution}:M{derniere_ligne-1}', {
            'backgroundColor': {'red': 0.95, 'green': 0.98, 'blue': 1}
        })

        # Retourner l'URL
        return spreadsheet.url

    except Exception as e:
        print(f"Erreur lors de la gestion du Google Sheet: {e}")
        return None

# Mettre à jour le Google Sheet et afficher le lien
print("=" * 50)
print("MISE À JOUR DU GOOGLE SHEET")
print("=" * 50)

url_sheet = gerer_google_sheet()

if url_sheet:
    print(f"✅ Google Sheet mis à jour avec succès!")
    print(f"🔗 Lien: {url_sheet}")
    print()
    print("Nouvelles données ajoutées:")
    print(f"- Fichier source: {CHEMIN_COMMANDES.split('/')[-1]}")
    print(f"- {len(donnees_livraisons)} arrêts pour cette tournée")
    print(f"- Distance totale: {distance_trajet} km")
    print(f"- Durée totale: {heures}h{minutes}min")
    print(f"- Date de livraison prévue: {(datetime.now(ZoneInfo('Europe/Paris')) + timedelta(days=2)).strftime('%d %B %Y').replace('January', 'Janvier').replace('February', 'Février').replace('March', 'Mars').replace('April', 'Avril').replace('May', 'Mai').replace('June', 'Juin').replace('July', 'Juillet').replace('August', 'Août').replace('September', 'Septembre').replace('October', 'Octobre').replace('November', 'Novembre').replace('December', 'Décembre')}")
    print()
    print("💡 Chaque nouvelle exécution ajoutera des lignes au même document")
else:
    print("❌ Erreur lors de la mise à jour du Google Sheet")
    print("Veuillez vérifier votre configuration d'authentification Google")

# Alternative: Afficher les données sous forme de DataFrame pandas
print("\n" + "=" * 50)
print("APERÇU DES DONNÉES (ORDRE OPTIMISÉ)")
print("=" * 50)

df_livraisons = pd.DataFrame(donnees_livraisons, columns=[
    'N° Trajet', 'N° Commande', 'Nom Client', 'Téléphone',
    'Adresse', 'Heure d\'arrivée prévue', 'Distance (km)', 'Durée trajet (min)'
])

print(df_livraisons.to_string(index=False))

print("\n" + "=" * 50)
print("FICHIERS GÉNÉRÉS")
print("=" * 50)
print(f"📍 Itinéraire et carte: Affichés ci-dessus")
print(f"📊 Google Sheet: {url_sheet if url_sheet else 'Erreur lors de la création'}")
print(f"🗂️ Ordre d'itinéraire: {chemin_ordre}")
print()
print("🔄 PROCHAINES ÉTAPES:")
print("1. Utilisez le fichier 'ordre_itineraire.csv' dans le générateur d'étiquettes")
print("2. Les étiquettes seront automatiquement triées par ordre de livraison")
print("3. Première commande à livrer = première ligne d'étiquettes")
