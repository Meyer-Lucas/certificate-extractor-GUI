# Génération d'un exe à partir du powershell en se basant sur le module de PS2EXE de MScholtes
# Page Powershell Gallery : https://www.powershellgallery.com/packages/ps2exe/
# Page Github du projet : https://github.com/MScholtes/PS2EXE

# Vérification de la présence de PS2EXE sur le poste, si non présent alors il sera télécharger
if ((Get-Command -All).Name -notcontains "PS2EXE") {
    Install-Module ps2exe
}

# Création de l'exe
ps2exe -inputFile .\certificate-extractor-GUI.ps1 `
    -outputFile .\certificate-extractor-GUI.exe `
    -noConsole `
    -title "Certificate Generator GUI" `
    -description "Visualiser et extraire facilement des certificats ainsi que les autorités de certification associées avec un outil graphique." `
    -copyright "Meyer Lucas - MIT License" `
    -version "2.2" `
    -supportOS