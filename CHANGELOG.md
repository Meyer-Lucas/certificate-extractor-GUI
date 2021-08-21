# CHANGELOG des différentes versions de Certificate Extractor GUI

## 2.2.4
Correction de l'affichage de la fenêtre avec l'exécution du ps1 où aucun espacement n'apparaissait en bas et à droite contrairement avec l'exe.

## 2.2.3
Vérification du support et du bon fonctionnement de l'application sur Windows 11.
Ajout de la version, de l'auteur et du lien Github vers l'outil.

## 2.2.2
Correction sur l'accès au répertoire temporaire qui peut ne pas fonctionner correctement avec l'utilisation de profil itinérant.
La variable système TEMP n'est plus utilisé au profit de LOCALAPPDATA\Temp qui ne provoque aucune erreur.

## 2.2.1
Correction de l'encodage des caractères pour l'exe.
Les caractères accentués n'étaient pas correctement affichés et provoquait des erreurs de fonctionnement.

## 2.2
Validation du fonctionnement pour les OS suivants : 
* Windows 10
* Windows Serveur 2019
* Windows Serveur 2016
* Windows 8.1
* Windows Serveur 2012 R2
* Windows 8 (les p7b ne sont pas supportés)
* Windows 2012 (les p7b ne sont pas supportés)

Ajout d'une vérification de l'OS et message d'avertissement si l'OS n'est pas supporté par le script.

## 2.1
Ajout de fonctionnalité :
* Ajout de la possibilité de sélectionner un certificat via drag and drop, des indications visuelles montrent l'emplacement de la zone de drop et la couleur indique si le drop est possible

## 2.0
Ajout de fonctionnalités :
* Support des formats p7b, pem, crt, cer et txt
* Recherche des certificats de la chaîne de certification manquant dans le magasin de certificats si disponible
* Lorsque le certificat sélectionné n'est pas un p12/pfx, le mot de passe n'est pas demandé à être défini

Correction de bug :
* La sélection d'export de certificat pouvait ne pas fonctionner correctement pour les CN court suite à la nouvelle nomenclature introduite en v1.1

## 1.1
Ajout de fonctionnalités :
* Ajout d'une checkbox permettant de sélectionner s'il faut enregistrer l'export de certificat dans le même répertoire que le répertoire sélectionner
* Modification du nom des fichiers exportés pour inclure le CN puis entre parentèses l'organisation si présent dans le certificat ("O=" dans le DN)

Correction de bug :
* Lorsque plusieurs certificats contiennent un champ "O=" alors une différence est réalisée sur le nom d'export du certificat

## 1.0
Version initiale permettant de :
* Sélectionner un p12/pfx pour voir les différents certificats qu'il contient
* Exporter tous les certificats que contient le p12/pfx dans un répertoire donné
* Exporter certains certificats sélectionnés par l'utilisateur dans un répertoire donné