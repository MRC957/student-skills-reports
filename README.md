(english version below)
# [FR] Rapports de compétences des étudiants
Automatisation des rapports de compétences pour les athénées de Wallonie

## Contexte
Dans l'enseignement secondaire en Wallonie, il est demandé de créer un rapport de compétences pour chaque élève, résumant les compétences acquises ou non au cours de l'année. L'acquisition de ces compétences se reflète dans les notes obtenues lors des évaluations réalisées tout au long de l'année scolaire. Ce rapport de compétences est donc une simple mise en forme des données déjà présentes dans le bulletin.

Afin d'éviter le travail fastidieux et à faible valeur ajoutée de recopiage de ces informations, ce programme génère un fichier PDF complet basé sur les notes de chaque élève dans un fichier Excel, ainsi qu'un modèle Word pour le rapport de compétences.

## Pré-requis
En bref:
- Utiliser l'Excel tout au long de l'année pour écrire les notes, puis appuyer sur le bouton `export CSV`.
- Créer des copies du template Word pour chaque classe et remplir les descriptions dans les tableaux

Les macros doivent être activées dans Excel et Word. 
### Excel:
- 1 page par classe. Copier les pages existantes si nécessaire (/!\ La feuille `export CSV` doit rester la dernière feuille).
- Pour chaque feuille :
    - Le nom de la classe doit être indiqué en haut à gauche.
    - Remplir Nom/Prénom de chaque élève + supprimer les lignes superflues.
    - Le titre des évaluations doit contenir `UAA` s'il s'agit d'une UAA. La détection est ensuite automatique. Le pourcentage pour lequel les UAA comptent dans la moyenne pondérée doit être indiqué en haut à droite.
    - La note maximale de l'évaluation doit être mise dans la ligne `Total`.
    - La date est facultative.
    - Si l'élève est absent, indiquer `ABS`.
    - La ligne bleue doit être complétée en accord avec le fichier Word pour pouvoir placer chaque évaluation par rapport aux tableaux dans Word. La syntaxe est "\<X>\<Y>" avec X comme numéro de tableau (1 ou 2) et Y comme ligne représentant l'évaluation dans ce tableau (ex: "23" pour a ligne 3 du tableau 2).

### Word:
- Créer 1 copie de `[template].docm` pour chaque classe et la renommer avec le nom de la classe (identique au nom dans l'Excel).
- Pour chaque classe :
    - Mettre à jour le titre et le pied de page (/!\ Ne pas toucher à l'en-tête qui sera automatiquement complétée par le nom de l'élève).
    - Supprimer les lignes inutiles des tableaux.
    - Compléter la première colonne de chaque tableau par rapport à l'évaluation évaluée. L'ordre des lignes correspond à ce qui a été écrit dans les cases bleues de l'Excel.

## Mode d'emploi pour créer le Rapport de Compétences
Les fichiers Excel et Word se trouvent dans `src/inputs`. Les fichiers de sortie se rouveront dans `src/outputs`.

1. Une fois l'Excel complété, cliquez sur le bouton `ExportCSV` de la feuille 'export CSV'. Un fichier CSV sera créé dans le dossier 'outputs', qui sera utilisé pour les étapes suivantes. Des messages contextuels apparaîtront pour fournir des retours sur ce qui se passe.

2. Exécutez le programme Python avec la ligne de commande `python create_RDC.py` depuis le dossier `src` ou cliquez sur `launch_script.bat`.
Si Python n'est pas installé, exécutez le fichier `install_python.bat`, redémarrez votre ordinateur, puis exécutez-le à nouveau. Cela peut prendre plusieurs minutes et déclencher l'antivirus. Vous pouvez aussi l'instaler manuellement depuis [ce lien](https://www.python.org/downloads/) 

3. Pour chaque classe, un dossier sera créé contenant les rapports de compétences en Word et en PDF pour chaque élève. Les PDF seront ensuite tous fusionnés dans le fichier `All.pdf` dans le dossier `outputs`.


## Détails de fonctionnement
Ce projet demande l'utilisation d'un template Excel, d'un template Word, et d'un script Python.

### Excel:
- Chaque feuille Excel permet de noter les points obtenus par les élèves à chaque interro pendant l'année. Des statistiques sont calculées en temps réel ainsi qu'un graphique pour faciliter la comparaison des notes moyennes obtenues. Cette moyenne globale est pondérée en suivant le principe des interros "UAA" qui comptent plus que les interros "non UAA". Ce programme interprète une interro comme une UAA si `UAA` est présent dans le titre.
- Lorsque la macro `ExportToCsv` est activée, elle nettoie la dernière feuille, y recopie les notes inscrites dans les autres feuilles, les exportent dans un fichier CSV `outputs\Data.csv`, puis nettoie à nouveau le feuille. Le format pécis de ce CSV est celui attendu par le script Python.
- Pour les périodes suivantes, il est également possible de garder ce template en effaçant uniquement les notes (i.e. on garde le nom des élèves et la classe dans chaque feuille) grâce à la macro `ResetResults` de la dernière feuille.

### Word:
- Le template Word est conçu pour compléter les cases de chaque tableau en fonction des propriétés du document. Ces propriétés sont instanciées à l'ouverture du document grâce à son nom `Nom_Prénom_classe__11_12_...__21_22_...` (généré par le script Python pour chaque élève).
- Lors de la fermeture du document, un dossier et un fichier PDF sont créés.
- Le template Word doit être instancié manuellement pour chaque classe:
    - La taille des tableaux doit être adaptée par rapport aux interros UAA/non-UAA renseignées dans l'Excel.
    - La description des interros doit être ajoutée
    - Le titre et le pied de page doivent être mis à jour manuellement (/!\ pas l'en-tête qui est mise à jour automatiquement avec les informations de l'élève)
- Le template Word par classe est instancié automatiquement par le script Python pour chaque élève.

### Python:
- Le script Python permet de générer un PDF qui regroupe les Rapports de Compétences de tous les élèves présents dans l'Excel. Sur base de `outputs\Data.csv`, il crée un Word par élève puis l'ouvre et le ferme pour générer le PDF associé. Il fusionne ensuite tous ces PF en 1 seul fichier.
- Le script Python peut être utilisé en entier ou uniquement pour sa partie concaténation de PDF.
- Le script Python utilise un DataFrame pandas pour gérer les data. Pour chaque classe, les résultats de chaque élève sont transformés en 3 catégories:
    - 0 = a obtenu moins de la moitié
    - 1 = a obtenu plus de la moitié
    - 2 = absent
- Pour chaque élève, le template Word contenant la classe dans son titre (ex: resultats_6TQ.docm) est copié dans le dossier `outputs\<classe>\` pour être instancié avec le nom de l'élève et ses résultats. Le script Python se charge ensuite d'ouvrir et fermer ce Word afin que les macros du Words crée le PDF associé dans le dossier `outputs\<classe>\PDF`
- Finalement, le script Python lit tous les PDF et les fusionne en 1 PF par classe. Il fusionne ensuite ces "PDF par classe" en 1 seul PDF `outputs\All.pdf`



# [EN] Student Skills Reports
Automation of skill reports for students in public schools of Wallonia

## Context (EN)
In secondary education in Wallonia, it is required to create a skills report for each student, summarizing the skills acquired or not during the year. The acquisition of these skills is reflected in the grades obtained during evaluations throughout the school year. This skills report is therefore a simple formatting of the data already present in the report card.
To avoid the tedious and low-value-added work of copying this information, this program generates a complete PDF file based on the grades of each student in an Excel file, as well as a Word template for the skills report.

### Prerequisites (EN)
Macros must be enabled in Excel and Word.
...
