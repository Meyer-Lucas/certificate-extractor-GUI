# Certificate Extractor GUI
Visualiser et extraire facilement des certificats ainsi que les autorités de certification associées avec un outil graphique.

L'objectif de cet outil étant de simplifier et d'accéler la récupération des autorités de certification contenu dans un certificat. Des solutions comme passer par le magasin de certificat, utiliser les commandes OpenSSL voire l'utilitaire certutil.exe de Windows sont fonctionnelles mais pas pratique ou rapide.

Cet outil pour fonctionner se base sur l'outil certutil ainsi que le magasin de certificat local pour lire le contenu des certificats, les exporter ou récupérer les autorités de certification.

Les certificats exportés auront l'extension crt et seront encodés au format DER. Le nom sera de la forme suivante :
```
[<Niveau> - ]<Common Name>[ (<Organisation>)].crt
```

Où :
* _Niveau_ correspond au niveau du certificat par rapport à l'autorité de certification racine (niveau 0) si disponible
* _Common Name_ correspond au CN du certificat
* _Organisation_ correspond à l'organisation s'il est renseigné dans l'objet du certificat (champ O)

## Fonctionnalités proposées 
Les fonctionnalités principales sont les suivantes :
* Interface graphique simple mais efficace
* Visualiser le contenu des fichiers de certificats communs : pfx/p12/p7b/pem/crt/cer/txt
* Reconstruction de la chaîne de certification si les autorités intermédaires ou racine ne sont pas inclus dans le fichier et si ces certificats sont disponibles dans le magasin de certification
* Export de tous les certificats ou ceux sélectionnés par l'utilisateur 

Les fonctionnalités secondaires sont les suivantes :
* Affichage des informations de niveau (si la remontée vers l'autorité de certification racine a été réussie)
* Affichage des informations d'expiration du certificat
* Affichage de l'objet ainsi que l'émetteur du certificat
* Drag and Drop possible pour visualiser le contenu du fichier de certificat 
* Code couleur pour le drop d'un certificat : zone bleu = possible de déposer le certificat ; zone rouge = le ou les fichiers pris en drag ne seront pas acceptés
* Export dans le répertoire du certificat sélectionné ou dans un répertoire définit

## OS Supportés 
Les OS suivants ont été testés et validés pour un fonctionnement complet de l'outil :
* Windows 10
* Windows Serveur 2019
* Windows Serveur 2016
* Windows 8.1
* Windows Serveur 2012 R2
* Windows 8 (les p7b ne sont pas supportés)
* Windows 2012 (les p7b ne sont pas supportés)

## Limitations connues
L'outil ne peut gérer qu'un fichier à la fois, il ne sera pas possible de visualiser de multiples certificats avec une sélection, c'est aussi valable pour le Drag and Drop. Pour autant il est possible d'enchainer la visualisation de plusieurs certificats en sélectionnant un nouveau certificat à charger.

L'outil ne va rechercher le certificat et sa chaîne de certification que pour le premier certificat trouvé. Un pfx avec toute sa chaîne de certification ne pose pas de problème. Ce serait des p7b qui contienne différents certificats qui pourrait ne pas afficher toutes les informations qui y sont contenues.
