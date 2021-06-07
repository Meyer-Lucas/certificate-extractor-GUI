<#

.SYNOPSIS
Script Powershell GUI permettant de facilement pouvoir extraire tous les éléments ou seulement certains des pfx/p12.

.DESCRIPTION

Script PowerShell fonctionnant avec une interface graphique permettant de manipuler les certificats.

Cet outil a comme objectif de :
    - simplifier l'extraction des autorités de certification intermédiaires et racine
    - simplifier la transformation d'un p12/pfx vers des crt/pem/cer

Cet outil est un substitut plus pratique et simple que les commandes OpenSSL ou bien d'avoir à importer les p12/pfx dans le magasin de certificat pour pouvoir extraire ses composants.

/!\ Le fonctionnement de l'outil s'appuie sur la commande certutil.exe qui limite le support à Windows.

.NOTE
    Version : 1.1
    Auteur  : Lucas MEYER
    Github  : https://github.com/Meyer-Lucas/certificate-extractor-GUI
    Licence : MIT License

#>

# Préparation du script pour une utilisation graphique
Add-Type -AssemblyName system.windows.forms
[System.Windows.Forms.Application]::EnableVisualStyles()


# -------------------------------------------------- Variables globales ----------------------------------------------


[String]$Font = "Segoe UI, 11"
[String]$FontMin = "Segoe UI, 8"
[Int]$LargeurFenêtre = 600
[Int]$HauteurFenêtre = 495

$CertificatsSHA1 = New-Object System.Collections.ArrayList
$CertificatsExpiration = New-Object System.Collections.ArrayList
$CertificatsEmetteur = New-Object System.Collections.ArrayList
$CertificatsObjet = New-Object System.Collections.ArrayList
$OrdreTriCertificat = New-Object System.Collections.ArrayList
$CertificatNiveau = New-Object System.Collections.ArrayList


# ----------------------------------------------------- Fonctions ----------------------------------------------------


# Désactive la group box n°3, vide la liste de certificat et réinitialise la valeur de la check box d'export du p12
function ResetGroupBoxCertificatListe {
    $GroupBoxCertificatListe.Enabled = $false
    $DataGridViewCertificatListe.Rows.ForEach({$DataGridViewCertificatListe.Rows.Remove($_)})
    $DataGridViewCertificatListe.Rows.ForEach({$DataGridViewCertificatListe.Rows.Remove($_)}) # Duplication de la ligne pour supprimer la ligne restante après première passe
    $CheckBoxEnregistrerMemeRepertoire.Checked = $true
    $CertificatsSHA1.Clear()
    $CertificatsExpiration.Clear()
    $CertificatsEmetteur.Clear()
    $CertificatsObjet.Clear()
    $OrdreTriCertificat.Clear()
    $CertificatNiveau.Clear()
}

# Vérification que le mot de passe fourni est bien le bon pour le certificat choisi
function TestMotDePasse {
    certutil.exe -p $TextBoxMotDePasse.Text -dump $labelEmplacementCertificat.Text | Out-Null
    return $?
}

# Complète la DataGridView avec les différents certificats trouvés
function CompleteDataGridViewCertificats {
    $SortieCertutil = certutil.exe -p $TextBoxMotDePasse.Text -dump $labelEmplacementCertificat.Text
    
    $SortieCertutil | Select-String -CaseSensitive "sha1" | ForEach-Object { $CertificatsSHA1.Add($_.ToString().Split(" ")[-1]) }
    $SortieCertutil | Select-String -CaseSensitive "After" | ForEach-Object { $CertificatsExpiration.Add($_.ToString().Substring(12)) }

    $switch = $true
    $SortieCertutil | Select-String -CaseSensitive "CN=" | ForEach-Object { 
        $temp = $_.ToString().Split(":")[1].Trim()
        if ($switch) { $CertificatsEmetteur.Add($temp) }
                else { $CertificatsObjet.Add($temp) }
        $switch = !$switch
    }

    # Recherche d'un certificat racine et classement des certificats tels que donnés par certutil
    for ($i = 0 ; $i -lt $CertificatsEmetteur.Count ; $i++) {
        if ($CertificatsEmetteur[$i] -eq $CertificatsObjet[$i]) { $OrdreTriCertificat.Add($i) }
    }
    if ($OrdreTriCertificat.Count -ne 0) {
        $CertificatNiveau.Add("0")
        $antiInfiniteLoop = 0
        $i = 0
        $niveau = 1
        while ($OrdreTriCertificat.Count -ne $CertificatsEmetteur.Count -and $antiInfiniteLoop -lt 1000) {
            if ($CertificatsEmetteur[$i] -eq $CertificatsObjet[$OrdreTriCertificat[-1]] -and $OrdreTriCertificat -notcontains $i) { 
                $OrdreTriCertificat.Add($i) 
                $CertificatNiveau.Add($niveau)
                $niveau++
            }
            $i++
            if ($i -eq $CertificatsEmetteur.Count) { $i = 0 }
            $antiInfiniteLoop++
        }
    } else {
        $i = 0
        $CertificatsEmetteur.ForEach({$CertificatNiveau.Add("?") ; $OrdreTriCertificat.Add($i) ; $i++})
    }
    
    # Ajout des certificats dans le DataGridView
    $i = 0
    $OrdreTriCertificat.ForEach({
        $DataGridViewCertificatListe.Rows.Add($CertificatNiveau[$i], $CertificatsExpiration[$_], $CertificatsObjet[$_], $CertificatsEmetteur[$_])
        $i++
    })
}

# Fonction chargée de renommée les certificats qui ont étés exportés
function RenommeCertificat {
    param([string]$Repertoire)
    $Certificats = Get-ChildItem -File $RepertoireTemp.FullName
    $Certificats.ForEach({ 
        for ($i = 0 ; $i -lt $CertificatsSHA1.Count ; $i++) {
            if ($_.ToString() -match $CertificatsSHA1[$i]) {
                $Nom = $CertificatsObjet[$i].Substring($CertificatsObjet[$i].IndexOf("CN=")).Split(",")[0].Substring(3).Replace("*","_")
                if ($CertificatsObjet[$i] -match "O=") { $Nom += " (" + $CertificatsObjet[$i].Substring($CertificatsObjet[$i].IndexOf("O=")+2).Split(",")[0].Replace("*","_") + ")" }
                if ($CertificatNiveau[$OrdreTriCertificat[$i]] -ne "?") { $Nom = $CertificatNiveau[$OrdreTriCertificat[$i]].ToString() + " - "  + $Nom }
                $Nom += ".crt"
                Move-Item -Force $_.FullName -Destination $Nom
            }
        }
        if ($_.Name -match "\.p12$") { Remove-Item $_.FullName }
    })
}


# ---------------------------------------------------- Gestion GUI ---------------------------------------------------


$fenêtreCertificateExtractor = New-Object System.Windows.Forms.Form -Property @{
    Height = $HauteurFenêtre
    Width = $LargeurFenêtre
    StartPosition = 'CenterScreen'
    Text = 'Certificate Extractor'
    Font = $Font
    Autosize = $false
    SizeGripStyle = "Hide"
    MaximizeBox = $false
    FormBorderStyle = "FixedSingle"
    Padding = 0
}

# -------------------- GroupBoxes

$GroupBoxCertificat = New-Object System.Windows.Forms.GroupBox -Property @{
    Text = "1 - Choix du certificat : "
    Location = "10, 10"
    Height = 100
    Width = $LargeurFenêtre - 30
    Padding = 0
    Font = "Segoe UI, 10"
}
$fenêtreCertificateExtractor.Controls.Add($GroupBoxCertificat)

$GroupBoxMotDePasse = New-Object System.Windows.Forms.GroupBox -Property @{
    Text = "2 - Mot de passe du certificat : "
    Location = "10, 120"
    Height = 100
    Width = $LargeurFenêtre - 30
    Padding = 0
    Font = "Segoe UI, 10"
    Enabled = $False
}
$fenêtreCertificateExtractor.Controls.Add($GroupBoxMotDePasse)

$GroupBoxCertificatListe = New-Object System.Windows.Forms.GroupBox -Property @{
    Text = "3 - Liste des certificats contenus : "
    Location = "10, 230"
    Height = 220
    Width = $LargeurFenêtre - 30
    Padding = 0
    Font = "Segoe UI, 10"
    Enabled = $false
}
$fenêtreCertificateExtractor.Controls.Add($GroupBoxCertificatListe)

# -------------------- Contenu de la GroupBox n°1 : recherche du certificat

$buttonSélectionCertificat = New-Object System.Windows.Forms.Button -Property @{
    Text = "Ouvrir"
    Location = "20, 38"
    Font = $Font
    Autosize = $true
}
$GroupBoxCertificat.Controls.Add($buttonSélectionCertificat)

$labelEmplacementCertificat = New-Object System.Windows.Forms.Label -Property @{
    Text = "Pas de certificat sélectionné"
    TextAlign = "MiddleLeft"
    Location = "120, 20"
    Width = ($LargeurFenêtre-170)
    Height = 70
    Font = $Font
}
$GroupBoxCertificat.Controls.Add($labelEmplacementCertificat)

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
	Filter = 'Certificat (*.pfx ou *.p12)|*.pfx;*.p12'
    FileName = ""
    Title = "Choix du certificat pour Certificate Extractor"
}

# -------------------- Contenu de la GroupBox n°2 : Spécification du mot de passe du certificat

$TextBoxMotDePasse = New-Object System.Windows.Forms.TextBox -Property @{
    Location = "10, 20"
    ScrollBars = 0
    Multiline = $False
    Enabled = $True
    ReadOnly = $False
    Font = $Font
    Width = $LargeurFenêtre - 50
    Height = 30
    PasswordChar = "*"
}
$GroupBoxMotDePasse.Controls.Add($TextBoxMotDePasse)

$CheckBoxMotDePasseClair = New-Object System.Windows.Forms.CheckBox -Property @{
    Font = $Font
    Text = 'Afficher le mot de passe'
    Location = "10, 60"
    Checked = $false
    Width = 200
}
$GroupBoxMotDePasse.Controls.Add($CheckBoxMotDePasseClair)

$ButtonValidationMotDePasse = New-Object System.Windows.Forms.Button -Property @{
    Text = "Valider"
    Location = (""+($LargeurFenêtre - 114)+", 55")
    Font = $Font
    Autosize = $true
}
$GroupBoxMotDePasse.Controls.Add($ButtonValidationMotDePasse)

# -------------------- Contenu de la GroupBox n°3 : Visualisation des certificats et export

$DataGridViewCertificatListe = New-Object System.Windows.Forms.DataGridView -Property @{
    Width = $LargeurFenêtre - 50
    Height = 150
    Location = "10, 20"
    Font = $FontMin
    ColumnCount = 4
    ColumnHeadersVisible = $true
    RowHeadersVisible = $false
    AllowUserToResizeRows = $false
    AllowUserToAddRows = $false
    SelectionMode = "FullRowSelect"
    MultiSelect = $true
    ReadOnly = $true
    ScrollBars = "Both"
    AutoSizeColumnsMode = "AllCells"
}
$GroupBoxCertificatListe.Controls.Add($DataGridViewCertificatListe)

# Définition des colonnes du DataGridView
$DataGridViewCertificatListe.Columns[0].Name = "Niveau"
$DataGridViewCertificatListe.Columns[1].Name = "Expiration"
$DataGridViewCertificatListe.Columns[2].Name = "Objet"
$DataGridViewCertificatListe.Columns[3].Name = "Emetteur"

# Empêchement de trier les différentes colonnes
$DataGridViewCertificatListe.Columns.ForEach({ $_.SortMode = 0 })

$CheckBoxEnregistrerMemeRepertoire = New-Object System.Windows.Forms.CheckBox -Property @{
    Font = $FontMin
    Text = 'Exporter les certificats dans le même répertoire que le répertoire du certificat sélectionné'
    Location = "10, 175"
    Checked = $true
    Width = 275
    Height = 40
}
$GroupBoxCertificatListe.Controls.Add($CheckBoxEnregistrerMemeRepertoire)

$ButtonExportSelection = New-Object System.Windows.Forms.Button -Property @{
    Text = "Exporter la sélection"
    Location = (""+($LargeurFenêtre - 280)+", 178")
    Font = $Font
    Autosize = $true
}
$GroupBoxCertificatListe.Controls.Add($ButtonExportSelection)

$ButtonExport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Exporter"
    Location = (""+($LargeurFenêtre - 114)+", 178")
    Font = $Font
    Autosize = $true
}
$GroupBoxCertificatListe.Controls.Add($ButtonExport)

$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    RootFolder = 'MyComputer'
    Description = "Où exporter la sélection ?"
}


# ----------------------------------------------- Evénements graphiques ----------------------------------------------


$buttonSélectionCertificat.add_Click({
    if ($OpenFileDialog.ShowDialog() -eq "OK") {
        $labelEmplacementCertificat.Text = $OpenFileDialog.FileName
        $GroupBoxMotDePasse.Enabled = $true
        $TextBoxMotDePasse.Text = ""
        ResetGroupBoxCertificatListe
        $TextBoxMotDePasse.Focus()
    }
})

$TextBoxMotDePasse.add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $_.SuppressKeyPress = $true
        $ButtonValidationMotDePasse.PerformClick()
    }
})

$CheckBoxMotDePasseClair.add_click({
    if ($CheckBoxMotDePasseClair.Checked) { $TextBoxMotDePasse.PasswordChar = 0 }
    else { $TextBoxMotDePasse.PasswordChar = "*" }
})

$ButtonValidationMotDePasse.add_Click({
    ResetGroupBoxCertificatListe
    if (TestMotDePasse) { 
        $GroupBoxCertificatListe.Enabled = $true
        CompleteDataGridViewCertificats
        $ButtonExport.Focus()
    } else {
        [System.Windows.Forms.MessageBox]::Show("Le mot de passe indiqué n'est pas bon !",'Erreur','OK','Error')
    }
})

$ButtonExportSelection.add_Click({
    if (!$CheckBoxEnregistrerMemeRepertoire.Checked) {
        if ($FolderBrowser.ShowDialog() -eq "OK") { $Repertoire = $FolderBrowser.SelectedPath } 
                                             else { $Repertoire = $null }
    } else { $Repertoire = (Split-Path -path $OpenFileDialog.FileName) }
    
    if ($Repertoire -ne $null) {
        certutil.exe -p $TextBoxMotDePasse.Text -dump -split -silent $labelEmplacementCertificat.Text
        RenommeCertificat

        Get-ChildItem -File | ForEach-Object {
            $NomFichier = $NomFichierTemp = $_.Name
            if ($NomFichier -match "^[?|0-9]* - ") { $NomFichierTemp = $NomFichier.Substring(6) }
            $NomFichierTemp = $NomFichierTemp.Substring(0, ($NomFichierTemp.IndexOf(" (")))
            $DataGridViewCertificatListe.SelectedRows.ForEach({
                if ($_.Cells.Item("Objet").Value -match $NomFichierTemp) { Move-Item $NomFichier -Destination $Repertoire -Force }
            })
        }

        if ($DataGridViewCertificatListe.SelectedRows.Count -eq 1) { [System.Windows.Forms.MessageBox]::Show("L'export du certificat sélectionné est terminée.", "Tâche terminée", "OK","Info") }
        else { [System.Windows.Forms.MessageBox]::Show("L'export des certificats sélectionnés est terminée.", "Tâche terminée", "OK","Info") }
    }
})

$ButtonExport.add_Click({
    if (!$CheckBoxEnregistrerMemeRepertoire.Checked) {
        if ($FolderBrowser.ShowDialog() -eq "OK") { $Repertoire = $FolderBrowser.SelectedPath } 
                                             else { $Repertoire = $null }
    } else { $Repertoire = (Split-Path -path $OpenFileDialog.FileName) }
    
    if ($Repertoire -ne $null) {
        certutil.exe -p $TextBoxMotDePasse.Text -dump -split -silent $labelEmplacementCertificat.Text
        RenommeCertificat
        Move-Item *.crt -Destination $Repertoire -Force
        [System.Windows.Forms.MessageBox]::Show("L'export de tous les certificats est terminée.", "Tâche terminée", "OK","Info")
    }
})


# ---------------------------------------------- Affichage de la fenêtre ---------------------------------------------


$RepertoireTemp = New-Item -ItemType Directory ($env:TEMP+"\Certificate-Extractor-" + (Get-Random -Maximum 99999))
Set-Location $RepertoireTemp.FullName

$fenêtreCertificateExtractor.ShowDialog() | Out-Null

Set-Location $env:TEMP
Remove-Item -Recurse -Path $RepertoireTemp.FullName