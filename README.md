
L'objectif principal de cette application est de simplifier diverses tâches afin de les accélérer et de réduire le risque d'erreurs résultant de la copie manuelle de grandes quantités d'informations. Le développement devait se concentrer sur la création d'une interface très simple et ergonomique pour que la simplification de la tâche soit réelle et importante.

Après avoir toutes les spécifications fonctionnelles possible, j'ai pu mettre en évidence les meilleures fonctionnalités de l'application.
Les fonctions principales de l’application devaient être :

-	L’affichage de contenu du fichier (vérifiant l’état des stocks à l’écran)
- Ajouter les nouveaux produits.
- Modifier les produits existants.
- Suppression les produits.
- Cas d’une livraison, saisir les quantités sorties de stock.
- Cas d’une retourner, saisir les quantités de produits retourné en stock.
- Création un fichier d’état de stock avec la date

De plus lorsque le client demande l'état du stock, nous devons afficher le contenu du fichier avec la date actuelle de l'export dans le nom du fichier (par exemple : Astra Zeneca- état du stock le 07 04 2022.xlsx si la demande d'export est le 04/07/2022).

![image](https://github.com/user-attachments/assets/2c59e53d-4777-4210-b6de-f6caacfd21f5)
![image](https://github.com/user-attachments/assets/e2382c84-34b3-468c-b6b4-a4daec4c0120)


#Les produits
Pour chaque produit, on souhaite connaître :

-	Son Référence
-	Son nom (Etude)
-	Note
-	Quantité
-	Commentaires
-	S’il est disponible
De nombreux produits peuvent avoir la même référence et le même nom d'étude, il est donc nécessaire de les distinguer dans les Notes, chaque produit a sa propre note.


Après avoir fait l’étude des besoins et réalisé un mini cahier des charges j’ai pu passer à la phase de conception de l’application. J’ai commencé de faire l’interface de saisie qui doit contenant les boutons el les fonctionnalités suivantes :
- Bouton Ajouter : l'ajout de nouvelles lignes avec une gestion d’erreur si on clique sur le bouton Valider sans remplir les 3 premiers champs (Référence, Etude et Note).
- Bouton Modifier : Modification les données avec une gestion d’erreur si on clique sur le bouton Valider sans modifier les données.
- Bouton Supprimer : la suppression de la ligne sélectionnée, avec un message à la fin de sa mission.
- Bouton Sortie de Stock : la saisie de la quantité que nous voulons sortir de stock avec soustraction du stock en cours, avec une gestion d’erreur si la quantité de sortie est supérieure à celle disponible en stock.
- Bouton Entrée au stock : la saisie de la quantité que nous voulons entrer dans le stock avec addition du stock en cours.
- Bouton Etat de Stock : L’exportation d'état de stock dans le même dossier de programme avec le même nom du fichier + la date d'exportation.
- Bouton Fermer : nous pouvons fermer le programme.
- Écran de vérification (une zone de liste contient toutes les valeurs de fichier) Avec la zone de liste on peut voir toutes les données et les modifications ce qu'on a fait.
 

#Procédure :
J’ai commencé avec la création de première macro qui est va permettre de lancer l’interface de l’application et j’ai le lié avec un bouton « Gestion de stock » dans la page d’accueil pour lancer le programme en un clic.

![image](https://github.com/user-attachments/assets/49fa93fd-98d9-42a1-b56e-5b372f0e87e9)


Ensuite j’ai constaté que le seul moyen d’avoir accès au contenu de fichier afin de faire l’écran de vérification était de créé une macro qui permet de copier les données de la colonne A à E, et jusqu’à la dernier ligne non vide grâce une méthode en VBA derniereLigne = lastRow(1) + 1, et de les afficher dans une zon de texte sur l’interface d’application.

![image](https://github.com/user-attachments/assets/63df4936-359a-4c76-bc36-9f2485c0a974)

![image](https://github.com/user-attachments/assets/b68ec3d2-58a8-4d37-8c0c-f4108495a398)

traitements j’ai commencé la partie la plus ardue du développement a été la rédaction de modules en Visual Basic.

J’ai divisé le code en plusieurs Marcos et en plusieurs boutons afin de sépares les fonctionnalités demandées pour qu’elles soient facile à utiliser et de les corriger en cas de problèmes sur le code,
Chaque bouton affiche une fenêtre lorsqu'on appuie dessus, j’ai créé 5 fenêtre liée avec les boutons Ajouter, Modifier, Sortie de stock et Entrée en Stock.
-	fenêtre d’Ajouter
-	fenêtre de Modifier
-	fenêtre d’Article
-	fenêtre Sortie de Stock
-	fenêtre Entrée en Stock
![image](https://github.com/user-attachments/assets/3870661d-6fcb-4795-bc56-8dab25fbb43c)

Toutes les fenêtres ont les champs pour saisir les informations de produit ainsi trois boutons :
-	Choix de l’article : pour afficher la fenêtre d’Article afin de choisir un article et remplir les champs de saisir automatiquement.
-	Fermer : pour fermer la fenêtre
-	Valider : pour valider les informations saisies
Chaque bouton avait une gestion d’erreur qui affiche un message en cas de manque d’information sur les champs de saisir ou aucune modification dans la fenêtre de modification.

![image](https://github.com/user-attachments/assets/a49621e3-80c7-41c1-9181-428d80b7c1c9)
![image](https://github.com/user-attachments/assets/20712ec9-5b90-4e11-8979-f1d13492fa38)

Voici la surface d’application gestion de stock et des exemples des fenêtres :

![image](https://github.com/user-attachments/assets/82b71f5c-cb27-402b-8403-468a45b78213)
![image](https://github.com/user-attachments/assets/9301165c-4a50-467f-924e-747b0e5ff7a5)

Pour tous les traitements j’ai mis un message de validation afin d’avoir la possibilité d’arrêter l’action si le bouton est enfoncé par erreur.

![image](https://github.com/user-attachments/assets/fb28a2ee-f76f-4671-9b0e-17682b6136b2)


Pour savoir si la macro a fini de tourner, à la fin de son exécution j’ai mis un message indiquant son arrêt.


![image](https://github.com/user-attachments/assets/da6ca075-c850-4247-9eda-6f908418dc77)




