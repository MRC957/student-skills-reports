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
Les fichiers Excel et Word se trouvent dans `src/inputs`. Les fichiers de sortie se retrouveront dans `src/outputs`.

1. Une fois l'Excel complété, cliquez sur le bouton `ExportCSV` de la feuille 'export CSV'. Un fichier CSV sera créé dans le dossier 'outputs', qui sera utilisé pour les étapes suivantes. Des messages contextuels apparaîtront pour fournir des retours sur ce qui se passe.

2. Exécutez le programme Python avec la ligne de commande `python create_RDC.py` depuis le dossier `src` ou cliquez sur `launch_script.bat`.
Si Python n'est pas installé, exécutez le fichier `install_python.bat`, redémarrez votre ordinateur, puis exécutez-le à nouveau. Cela peut prendre plusieurs minutes et déclencher l'antivirus. Vous pouvez aussi l'instaler manuellement depuis [ce lien](https://www.python.org/downloads/) 

3. Pour chaque classe, un dossier sera créé contenant les rapports de compétences en Word et en PDF pour chaque élève. Les PDF seront ensuite tous fusionnés dans le fichier `All.pdf` dans le dossier `outputs`.

## Détails de fonctionnement
Ce projet requiert l'utilisation d'un template Excel, d'un template Word et d'un script Python.

### Excel:
- Chaque feuille Excel permet d'enregistrer les points obtenus par les élèves à chaque interrogation pendant l'année. Des statistiques sont calculées en temps réel ainsi qu'un graphique pour faciliter la comparaison des notes moyennes obtenues. Cette moyenne globale est pondérée en suivant le principe des interrogations "UAA" qui ont plus de poids que les interrogations "non UAA". Ce programme interprète une interrogation comme une UAA si `UAA` est présent dans le titre.
- Lorsque la macro `ExportToCsv` est activée, elle nettoie la dernière feuille, y recopie les notes inscrites dans les autres feuilles, les exporte dans un fichier CSV `outputs\Data.csv`, puis nettoie à nouveau la feuille. Le format précis de ce CSV est celui attendu par le script Python.
- Pour les périodes suivantes, il est également possible de garder ce template en effaçant uniquement les notes (c'est-à-dire conserver le nom des élèves et la classe dans chaque feuille) grâce à la macro `ResetResults` de la dernière feuille.

### Word:
- Le template Word est conçu pour compléter les cases de chaque tableau en fonction des propriétés du document. Ces propriétés sont instanciées à l'ouverture du document grâce à son nom `Nom_Prénom_classe__11_12_...__21_22_...` (généré par le script Python pour chaque élève).
- Lors de la fermeture du document, un dossier et un fichier PDF sont créés.
- Le template Word doit être instancié manuellement pour chaque classe :
    - La taille des tableaux doit être adaptée par rapport aux interrogations UAA/non-UAA renseignées dans l'Excel.
    - La description des interrogations doit être ajoutée.
    - Le titre et le pied de page doivent être mis à jour manuellement (/!\ pas l'en-tête qui est mis à jour automatiquement avec les informations de l'élève).
- Le template Word par classe est instancié automatiquement par le script Python pour chaque élève.

### Python:
- Le script Python permet de générer un PDF qui regroupe les Rapports de Compétences de tous les élèves présents dans l'Excel. Sur la base de `outputs\Data.csv`, il crée un Word par élève puis l'ouvre et le ferme pour générer le PDF associé. Il fusionne ensuite tous ces PDF en un seul fichier.
- Le script Python peut être utilisé dans son intégralité ou uniquement pour sa partie de concaténation de PDF.
- Le script Python utilise un DataFrame pandas pour gérer les données. Pour chaque classe, les résultats de chaque élève sont transformés en 3 catégories :
    - 0 = a obtenu moins de la moitié
    - 1 = a obtenu plus de la moitié
    - 2 = absent
- Pour chaque élève, le template Word contenant la classe dans son titre (ex : resultats_6TQ.docm) est copié dans le dossier `outputs\<classe>\` pour être instancié avec le nom de l'élève et ses résultats. Le script Python se charge ensuite d'ouvrir et fermer ce Word afin que les macros du Word créent le PDF associé dans le dossier `outputs\<classe>\PDF`.
- Finalement, le script Python lit tous les PDF et les fusionne en un PDF par classe. Il fusionne ensuite ces "PDF par classe" en un seul PDF `outputs\All.pdf`.


# [EN] Student Skills Reports
Automation of skill reports for secondary schools in Wallonia.

## Context
In secondary education in Wallonia, it is required to create a skills report for each student, summarizing the skills acquired or not during the year. The acquisition of these skills is reflected in the grades obtained during evaluations conducted throughout the school year. This skills report is therefore a simple formatting of the data already present in the report card.

To avoid the tedious and low-value-added work of copying this information, this program generates a complete PDF file based on the grades of each student in an Excel file, as well as a Word template for the skills report.

## Prerequisites
In brief:
- Use Excel throughout the year to record grades, then press the `export CSV` button.
- Create copies of the Word template for each class and fill in the descriptions in the tables.

Macros must be enabled in Excel and Word.
### Excel:
- 1 page per class. Copy existing pages if necessary (/!\ The `export CSV` sheet must remain the last sheet).
- For each sheet:
    - The name of the class must be indicated at the top left.
    - Fill in the Last name/First name of each student + delete unnecessary rows.
    - The title of evaluations must contain `UAA` if it is a UAA. Detection is then automatic. The percentage for which UAAs count in the weighted average must be indicated at the top right.
    - The maximum score of the evaluation must be entered on the `Total` line.
    - The date is optional.
    - If the student is absent, indicate `ABS`.
    - The blue line must be completed in accordance with the Word file to be able to place each evaluation relative to the tables in Word. The syntax is "\<X>\<Y>" with X as the table number (1 or 2) and Y as the line representing the evaluation in this table (e.g., "23" for line 3 of table 2).

### Word:
- Create 1 copy of `[template].docm` for each class and rename it with the name of the class (identical to the name in Excel).
- For each class:
    - Update the title and footer (/!\ Do not touch the header, which will be automatically completed with the student's name).
    - Delete unnecessary rows from the tables.
    - Complete the first column of each table with respect to the evaluated assessment. The order of the lines corresponds to what was written in the blue cells of Excel.

## How to generate the Skills Report
The Excel and Word files are located in `src/inputs`. The output files will be found in `src/outputs`.

1. Once the Excel file is completed, click on the `ExportCSV` button on the 'export CSV' sheet. A CSV file will be created in the 'outputs' folder, which will be used for the following steps. Contextual messages will appear to provide feedback on what is happening.

2. Run the Python program with the command line `python create_RDC.py` from the `src` folder or click on `launch_script.bat`.
If Python is not installed, run the `install_python.bat` file, restart your computer, and then run it again. This may take several minutes and trigger the antivirus. You can also install it manually from [this link](https://www.python.org/downloads/).

3. For each class, a folder will be created containing the skills reports in Word and PDF for each student. The PDFs will then all be merged into the `All.pdf` file in the `outputs` folder.

## Operational details
This project requires the use of an Excel template, a Word template, and a Python script.

### Excel:
- Each Excel sheet allows you to record the points obtained by students for each test during the year. Real-time statistics are calculated as well as a graph to facilitate the comparison of average scores obtained. This overall average is weighted according to the principle of "UAA" tests, which carry more weight than "non-UAA" tests. This program interprets a test as a UAA if `UAA` is present in the title.
- When the `ExportToCsv` macro is activated, it cleans the last sheet, copies the scores recorded in the other sheets, exports them to a CSV file `outputs\Data.csv`, and then cleans the sheet again. The precise format of this CSV is the one expected by the Python script.
- For subsequent periods, it is also possible to keep this template by only erasing the scores (i.e., keep the names of the students and the class in each sheet) using the `ResetResults` macro of the last sheet.

### Word:
- The Word template is designed to complete the cells of each table according to the properties of the document. These properties are instantiated when the document is opened thanks to its name `Name_Firstname_class__11_12_...__21_22_...` (generated by the Python script for each student).
- When the document is closed, a folder and a PDF file are created.
- The Word template must be instantiated manually for each class:
    - The size of the tables must be adjusted according to the UAA/non-UAA tests entered in Excel.
    - The descriptions of the tests must be added.
    - The title and footer must be updated manually (/!\ not the header, which is automatically updated with the student's information).
- The Word template per class is automatically instantiated by the Python script for each student.

### Python:
- The Python script generates a PDF that gathers the Skill Reports of all students present in Excel. Based on `outputs\Data.csv`, it creates a Word document for each student, then opens and closes it to generate the associated PDF. It then merges all these PDFs into a single file.
- The Python script can be used in its entirety or only for its PDF concatenation part.
- The Python script uses a pandas DataFrame to manage the data. For each class, the results of each student are transformed into 3 categories:
    - 0 = obtained less than half
    - 1 = obtained more than half
    - 2 = absent
- For each student, the Word template containing the class in its title (e.g., results_6TQ.docm) is copied to the `outputs\<class>\` folder to be instantiated with the student's name and results. The Python script then opens and closes this Word document so that the Word macros create the associated PDF in the `outputs\<class>\PDF` folder.
- Finally, the Python script reads all the PDFs and merges them into one PDF per class. It then merges these "PDFs per class" into a single PDF `outputs\All.pdf`.
