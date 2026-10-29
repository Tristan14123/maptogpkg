# Conversion MapInfo TAB → GeoPackage

Script Python pour convertir récursivement des fichiers MapInfo (`.tab`) en GeoPackage (`.gpkg`), avec génération d'un rapport Excel des résultats.

Conçu pour s'exécuter avec le Python embarqué dans QGIS (testé avec QGIS 3.40.12).

## Fonctionnalités

- Parcours récursif d'un dossier source à la recherche de fichiers `.TAB`.
- Conversion de chaque TAB vers un GeoPackage via OGR/GDAL.
- Le GeoPackage final n'est publié que si :
  - la conversion réussit,
  - et qu'il contient au moins une entité.
- Suppression automatique du fichier temporaire en cas d'échec ou de couche vide.
- Génération d'un rapport Excel (`rapport_conversion_mapinfo_vers_gpkg.xlsx`) avec, pour chaque fichier TAB :
  - chemin du fichier source,
  - chemin du GeoPackage produit,
  - statut (`SUCCÈS`, `VIDE`, `ÉCHEC`),
  - nombre d'entités,
  - message détaillé.

## Prérequis

- QGIS installé (pour disposer de Python avec les modules `osgeo` et `openpyxl`).
- Le pilote OGR `GPKG` doit être disponible.

## Installation

1. Clonez ce dépôt :
   ```bash
   git clone https://github.com/VOTRE-UTILISATEUR/convertir-tab-vers-gpkg.git
   cd convertir-tab-vers-gpkg
   ```
2. Configurez les chemins dans le script (`convertir_tab_vers_gpkg.py`), section **PARAMÈTRES** :
   - `DOSSIER_SOURCE` : dossier contenant les fichiers `.TAB` à convertir.
   - `NOM_DOSSIER_SORTIE` : dossier où seront écrits les GeoPackages et le rapport.

   Vous pouvez aussi utiliser les variables d'environnement :
   - `TAB_SOURCE_DIR`
   - `GPKG_OUTPUT_DIR`

## Utilisation

Lancez le script avec le Python fourni avec QGIS :

```bash
"C:\Program Files\QGIS 3.40.12\apps\Python312\python.exe" convertir_tab_vers_gpkg.py
```

Adaptez le chemin vers `python.exe` selon votre installation de QGIS.

À la fin de l'exécution, un rapport Excel s'affiche avec le nombre de succès, de fichiers vides et d'échecs.

## Comportement par statut

| Statut | Signification | Fichier GPKG créé ? |
|--------|---------------|---------------------|
| `SUCCÈS` | Conversion vérifiée, au moins une entité | Oui |
| `VIDE` | Aucune couche ou aucune entité | Non |
| `ÉCHEC` | Erreur durant la conversion | Non (fichier temporaire supprimé) |

## Anonymisation

Ce script a été anonymisé pour publication sur GitHub : aucun chemin personnel, nom d'utilisateur ou nom d'organisation n'est présent dans le code. Pensez à adapter la section **PARAMÈTRES** avant votre première exécution.

## Licence

Ce projet est sous licence MIT. Voir le fichier `LICENSE` pour plus de détails.
