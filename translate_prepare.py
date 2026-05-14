import pandas as pd
from deep_translator import GoogleTranslator
from tqdm import tqdm
import concurrent.futures
import threading
from functools import partial
import json
import os
from pathlib import Path


# Cache global pour les traductions
translation_cache = {}
CACHE_FILE = "translation_cache.json"


def load_translation_cache():
    """Charge le cache de traductions depuis le fichier"""
    global translation_cache
    try:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                translation_cache = json.load(f)
                return len(translation_cache)
        else:
            translation_cache = {}
            return 0
    except:
        translation_cache = {}
        return 0


def save_translation_cache():
    """Sauvegarde le cache de traductions"""
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(translation_cache, f, ensure_ascii=False, indent=2)
    except:
        pass


def get_cached_translation(text, translator):
    """Obtient une traduction avec cache"""
    if text in translation_cache:
        return translation_cache[text]
    
    try:
        translated = translator.translate(text)
        translation_cache[text] = translated
        return translated
    except:
        return text


def translate_batch_optimized(batch, translator_instance):
    """Traduit un batch avec gestion d'erreur optimisée et cache"""
    results = []
    uncached_texts = []
    uncached_indices = []
    
    # Vérifier le cache pour chaque texte
    for i, text in enumerate(batch):
        if text in translation_cache:
            results.append(translation_cache[text])
        else:
            results.append(None)  # Placeholder
            uncached_texts.append(text)
            uncached_indices.append(i)
    
    # Traduire seulement les textes non cachés
    if uncached_texts:
        try:
            translated_uncached = translator_instance.translate_batch(uncached_texts)
            # Mettre à jour le cache et les résultats
            for i, (original, translated) in enumerate(zip(uncached_texts, translated_uncached)):
                translation_cache[original] = translated
                results[uncached_indices[i]] = translated
        except Exception as e:
            # En cas d'erreur, essayer individuellement
            for i, text in enumerate(uncached_texts):
                try:
                    translated = translator_instance.translate(text)
                    translation_cache[text] = translated
                    results[uncached_indices[i]] = translated
                except:
                    results[uncached_indices[i]] = text  # Garder original
    
    # Remplir les None restants (au cas où)
    for i, result in enumerate(results):
        if result is None:
            results[i] = batch[i]
    
    return results


def process_column_parallel(col_name, valeurs_uniques, progress_callback=None, batch_size=50, max_workers=4):
    """Traite une colonne avec parallélisation"""
    if not valeurs_uniques:
        return {}
    
    # Créer plusieurs instances de traducteur pour la parallélisation
    translators = [GoogleTranslator(source='auto', target='en') for _ in range(max_workers)]
    
    traductions = {}
    batches = [valeurs_uniques[i:i + batch_size] for i in range(0, len(valeurs_uniques), batch_size)]
    total_batches = len(batches)
    
    def update_progress(completed_batches):
        if progress_callback:
            progress = int((completed_batches / total_batches) * 100)
            progress_callback(f"Traduction {col_name}: {completed_batches}/{total_batches} lots ({progress}%)")
    
    completed = 0
    update_progress(0)
    
    # Utiliser ThreadPoolExecutor pour la parallélisation
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Soumettre tous les batches
        future_to_batch = {
            executor.submit(translate_batch_optimized, batch, translators[i % max_workers]): batch 
            for i, batch in enumerate(batches)
        }
        
        # Traiter les résultats au fur et à mesure
        for future in concurrent.futures.as_completed(future_to_batch):
            batch = future_to_batch[future]
            try:
                translated_batch = future.result()
                # Associer original et traduit
                for original, traduit in zip(batch, translated_batch):
                    traductions[original] = traduit
            except Exception as e:
                if progress_callback:
                    progress_callback(f"Erreur sur un batch: {e}")
                # En cas d'erreur, garder les originaux
                for original in batch:
                    traductions[original] = original
            
            completed += 1
            update_progress(completed)
    
    return traductions

def preparer_et_traduire_excel(fichier_entree, fichier_sortie, progress_callback=None):
    # Charger le cache au début
    load_translation_cache()
    
    if progress_callback:
        progress_callback("Préparation du fichier Excel...")
    else:
        print("Préparation du fichier Excel...")
    
    df = pd.read_excel(fichier_entree)
    
    df.columns = [str(c).strip() for c in df.columns]

    
    cols_a_combiner = ['Designation', 'Marque', 'Modèle']
    
    # Vérification que les colonnes existent avant de combiner
    for c in cols_a_combiner:
        if c not in df.columns:
            error_msg = f"Erreur : La colonne '{c}' est absente du fichier !"
            if progress_callback:
                progress_callback(error_msg)
            else:
                print(error_msg)
            return

    df['Name'] = (df['ID'].astype(str).replace('nan', '') + " " + 
                df['Designation'].astype(str).replace('nan', '') + " " + 
                df['Marque'].astype(str).replace('nan', '') + " " + 
                df['Modèle'].astype(str).replace('nan', ''))
    
    
    colonnes_finales = ['ID','Provenance', 'SN', 'Name', 'Commentaires','YOM']
    df = df[colonnes_finales]
    
    # Traduction avec optimisation
    if progress_callback:
        progress_callback("Traduction du fichier Excel (optimisée)...")
    else:
        print("Traduction du fichier Excel (optimisée)...")
    
    cols_to_translate = ['Name', 'Commentaires']
    
    # Préparer les données pour chaque colonne
    column_data = {}
    for col in cols_to_translate:
        valeurs_uniques = df[col].dropna().unique().tolist()
        # Filtre optimisé : supprimer vides, très courts, et numériques purs
        valeurs_uniques = [
            str(v).strip() for v in valeurs_uniques 
            if str(v).strip() and len(str(v).strip()) > 2 and not str(v).strip().isdigit()
        ]
        column_data[col] = valeurs_uniques
    
    # Traiter les colonnes en parallèle
    def column_callback(col_name, message):
        if progress_callback:
            progress_callback(f"[{col_name}] {message}")
    
    traductions_results = {}
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        # Soumettre les deux colonnes en parallèle
        future_to_col = {
            executor.submit(
                process_column_parallel, 
                col, 
                column_data[col], 
                partial(column_callback, col),
                50,  # batch_size augmenté
                3    # max_workers par colonne
            ): col for col in cols_to_translate
        }
        
        # Récupérer les résultats
        for future in concurrent.futures.as_completed(future_to_col):
            col = future_to_col[future]
            try:
                traductions_results[col] = future.result()
            except Exception as e:
                if progress_callback:
                    progress_callback(f"Erreur colonne {col}: {e}")
                traductions_results[col] = {}
    
    # Appliquer les traductions
    for col in cols_to_translate:
        if col in traductions_results:
            df[col] = df[col].astype(str).map(traductions_results[col]).fillna(df[col])

    df.to_excel(fichier_sortie, index=False)
    success_msg = f"Terminé ! Sauvegardé sous : {fichier_sortie}"
    if progress_callback:
        progress_callback(success_msg)
    else:
        print(success_msg)
    
    # Sauvegarder le cache mis à jour
    save_translation_cache()
    
    return df


def import_format_csv(fichier_excel,nom_fichier_csv="resultat.csv", progress_callback=None):
    if progress_callback:
        progress_callback("Importation du fichier Excel et conversion en CSV...")
    else:
        print("Importation du fichier Excel et conversion en CSV...")

    df = pd.read_excel(fichier_excel)
    df.to_csv(nom_fichier_csv, index=False, encoding='utf-8-sig')

    success_msg = f"Le fichier CSV '{nom_fichier_csv}' a été créé avec succès."
    if progress_callback:
        progress_callback(success_msg)
    else:
        print(success_msg)



def main():

    ####################################################################################
    numero_vente_hilditch = "9525" # Remplacez par le numéro de vente réel
    fichier_excel = "Vente_en_cours.xlsx" # Remplacez par le chemin réel de votre fichier Excel
    ####################################################################################

    preparer_et_traduire_excel(fichier_excel, numero_vente_hilditch + "_excel.xlsx")
    import_format_csv(numero_vente_hilditch + "_excel.xlsx", numero_vente_hilditch + "_CSV.csv")



if __name__ == "__main__":
    main()