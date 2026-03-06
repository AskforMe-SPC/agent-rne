# Agent RNE - version VS Code

## 1) Préparer l'environnement

```powershell
cd "c:\Users\askme\Downloads\Agent RNE"
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements_vscode.txt
```

Si `py` n'existe pas, utilise Python installé dans VS Code (`Python: Select Interpreter`) puis:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements_vscode.txt
```

## 2) Variables d'environnement

```powershell
$env:INPI_USERNAME="ton_login"
$env:INPI_PASSWORD="ton_password"
$env:INPI_NOMENCLATURE_XLSX="c:\Users\askme\Downloads\Dictionnaire_de_donnees_INPI_2025_05_09.xlsx"
```

## 3) Lancer

```powershell
python app_vscode.py
```

Puis ouvre: http://127.0.0.1:5000

## Ce qui change

- Analyse entreprise: **sans Mistral** (résumé déterministe depuis JSON INPI).
- Logo: externalisé en `static/logo.png` (et source conservée dans `assets/logo.png`).
- Nomenclature `typeDocument`: chargée depuis le fichier Excel INPI (onglet `typeDocument`).
