import pandas as pd
from deep_translator import GoogleTranslator

def traduire_texte(texte):
    translator = GoogleTranslator(source='auto', target='en')
    if pd.isna(texte) or str(texte).strip() == "":
        return texte
    try:
         return translator.translate(str(texte))
    except Exception as e:
         print(f"Erreur sur '{texte}': {e}")
         return texte



def traduire_excel(fichier_entree, fichier_sortie, colonnes_a_traduire):
    print(f"Traduction en cours pour : {fichier_entree}")
    df = pd.read_excel(fichier_entree)
    
    for col in colonnes_a_traduire:
        df[col] = df[col].apply(traduire_texte)

    df.to_excel(fichier_sortie, index=False)
    print(f"Terminé ! Sauvegardé sous : {fichier_sortie}")



def main():
    fichier = "test.xlsx"
    colonnes = ["Designation", "Commentaires"] 
    resultat = "excel_traduit.xlsx"

    traduire_excel(fichier, resultat, colonnes)


if __name__ == "__main__":
    main()