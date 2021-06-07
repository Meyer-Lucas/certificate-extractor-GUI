# CHANGELOG des différentes versions de Certificate Extractor GUI

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
