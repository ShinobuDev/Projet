# RA-PRO ℹ️

⚠️ Dans le cadre d'une alternance
( Aucune information privé sur une entreprise ne sera dévoilé )

## 1 - Quel est son but ?

Le but du logiciel de rapprochement bancaire est d'automatiser le processus de comparaison des transactions entre deux fichiers Excel (généralement des relevés bancaires) afin d'identifier les correspondances et les différences. Cela permet aux utilisateurs de vérifier rapidement et efficacement l'exactitude de leurs enregistrements financiers.
## 2 - Quel sont ses avantages et inconvénient ?

Voici les avantages et inconvénients du logiciel de rapprochement bancaire que nous avons développé :

### Avantages ✔️

1. **Interface Utilisateur Intuitive** :
   - Utilisation de `tkinter` et `customtkinter` pour une interface moderne et conviviale, facilitant l'utilisation même pour les utilisateurs non techniques.

2. **Automatisation du Rapprochement** :
   - Le logiciel automatise le processus de rapprochement bancaire, réduisant le temps et les efforts nécessaires pour comparer manuellement les transactions.

3. **Gestion des Fichiers Excel** :
   - Capacité à charger et traiter des fichiers Excel, qui sont couramment utilisés pour les relevés bancaires, ce qui facilite l'intégration dans les flux de travail existants.

4. **Rapport de Sortie** :
   - Génération d'un fichier de sortie contenant les résultats du rapprochement, ce qui permet une documentation facile et un suivi des transactions.

5. **Gestion des Erreurs** :
   - Intégration de la gestion des exceptions pour informer l'utilisateur des erreurs potentielles, ce qui améliore la robustesse du logiciel.

6. **Personnalisation** :
   - Possibilité d'adapter les couleurs et le style de l'interface pour correspondre à l'identité visuelle de l'organisation (Soliha Normandie).

### Inconvénients ❌

1. **Dépendance aux Fichiers Excel** :
   - Le logiciel nécessite que les données soient au format Excel, ce qui peut être une limitation si les utilisateurs préfèrent d'autres formats de fichiers.

2. **Complexité des Données** :
   - Si les fichiers Excel contiennent des données mal formatées ou des erreurs, cela peut entraîner des échecs dans le rapprochement ou des résultats incorrects.

3. **Fonctionnalités Limitées** :
   - Le logiciel se concentre principalement sur le rapprochement des montants et des descriptions. D'autres fonctionnalités avancées, comme l'analyse des tendances ou la gestion des exceptions complexes, ne sont pas incluses.

4. **Nécessité d'Installation** :
   - Les utilisateurs doivent installer Python et les bibliothèques nécessaires, ce qui peut être un obstacle pour ceux qui ne sont pas familiers avec la programmation.

5. **Performance** :
   - Pour des fichiers très volumineux, le traitement peut être lent, et des optimisations peuvent être nécessaires pour améliorer la performance.

6. **Manque de Support Multilingue** :
   - Actuellement, le logiciel est conçu pour un public francophone, ce qui peut limiter son utilisation dans des contextes multilingues.

## 3 - Comment fonctionne t-il ?

Voici une explication du fonctionnement du logiciel de rapprochement bancaire que nous avons développé :

### 1. **Interface Utilisateur (GUI)**

- **Création de la fenêtre principale** : Le logiciel utilise `tkinter` et `customtkinter` pour créer une interface graphique moderne. La fenêtre principale affiche le titre "Rapprochement Bancaire - SOLIHA Normandie" et utilise des couleurs correspondant à la charte graphique de SOLIHA.

- **Chargement des éléments graphiques** : 
  - Un logo est affiché en haut de la fenêtre.
  - Des champs de saisie permettent à l'utilisateur de sélectionner deux fichiers Excel (fichiers bancaires) et de spécifier un fichier de sortie.
  - Un bouton "Lancer le rapprochement" déclenche le processus de rapprochement.

### 2. **Sélection des fichiers**

- **Fonctionnalité de sélection de fichiers** : L'utilisateur peut sélectionner les fichiers à l'aide de boîtes de dialogue. Les fichiers doivent être au format Excel (.xlsx).

### 3. **Traitement des fichiers**

- **Chargement des fichiers** : Lorsque l'utilisateur clique sur le bouton pour lancer le rapprochement, le logiciel récupère les chemins des fichiers sélectionnés.

- **Classe `BankReconciliation`** : 
  - Cette classe est responsable de la logique de rapprochement. Elle charge les fichiers Excel, extrait les données pertinentes (dates, descriptions, montants) et effectue le rapprochement.
  - Les montants sont comparés pour trouver des correspondances entre les deux fichiers. Si des montants correspondent, les lignes sont enregistrées dans un fichier de sortie.

### 4. **Rapprochement**

- **Logique de rapprochement** : 
  - Le logiciel compare les montants des deux fichiers. Il vérifie d'abord les montants de débit et de crédit, puis les descriptions pour trouver des correspondances.
  - Si des correspondances sont trouvées, elles sont ajoutées à une liste de résultats.

### 5. **Exportation des résultats**

- **Fichier de sortie** : Une fois le rapprochement terminé, les résultats sont exportés dans un nouveau fichier Excel spécifié par l'utilisateur. Ce fichier contient les lignes correspondantes ainsi que les informations sur les montants.

### 6. **Barre de progression et messages d'état**

- **Barre de progression** : Pendant le traitement, une barre de progression indique à l'utilisateur que le logiciel est en cours d'exécution.
- **Messages d'état** : Des messages d'information ou d'erreur sont affichés pour informer l'utilisateur du succès ou de l'échec du traitement.

### 7. **Gestion des erreurs**

- **Gestion des exceptions** : Le logiciel inclut des blocs `try-except` pour gérer les erreurs potentielles, comme des fichiers non trouvés ou des problèmes de format, et affiche des messages d'erreur appropriés.










