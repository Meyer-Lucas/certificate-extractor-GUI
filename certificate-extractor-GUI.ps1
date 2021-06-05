<#

.SYNOPSIS
Script Powershell GUI permettant de facilement pouvoir extraire tous les �l�ments ou seulement certains des pfx/p12.

.DESCRIPTION

Script PowerShell fonctionnant avec une interface graphique permettant de manipuler les certificats.

Cet outil a comme objectif de :
    - simplifier l'extraction des autorit�s de certification interm�diaires et racine
    - simplifier la transformation d'un p12/pfx vers des crt/pem/cer
    - r�cup�rer la clef priv� d'un p12/pfx

Cet outil est un substitut plus pratique et simple que les commandes OpenSSL ou bien d'avoir � importer les p12/pfx dans le magasin de certificat pour pouvoir extraire ses composants.

/!\ Le fonctionnement de l'outil s'appuie sur la commande certutil.exe qui limite le support � Windows.

.NOTE
    Version : 0.1
    Auteur  : Lucas MEYER
    Github  : https://github.com/Meyer-Lucas
    Licence : MIT License

#>

# Pr�paration du script pour une utilisation graphique
Add-Type -AssemblyName system.windows.forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# -------------------------------------------------- Variables globales ----------------------------------------------


[String]$Font = "Segoe UI, 11"
[String]$FontMin = "Segoe UI, 8"
[Int]$LargeurFen�tre = 600
[Int]$HauteurFen�tre = 495

$CertificatsSHA1 = New-Object System.Collections.ArrayList
$CertificatsExpiration = New-Object System.Collections.ArrayList
$CertificatsClef = New-Object System.Collections.ArrayList
$CertificatsEmetteur = New-Object System.Collections.ArrayList
$CertificatsObjet = New-Object System.Collections.ArrayList
$OrdreTriCertificat = New-Object System.Collections.ArrayList
$CertificatNiveau = New-Object System.Collections.ArrayList


# ----------------------------------------------------- Fonctions ----------------------------------------------------


# D�sactive la group box n�3, vide la liste de certificat et r�initialise la valeur de la check box d'export du p12
function ResetGroupBoxCertificatListe {
    $GroupBoxCertificatListe.Enabled = $false
    $DataGridViewCertificatListe.Rows.ForEach({$DataGridViewCertificatListe.Rows.Remove($_)})
    $DataGridViewCertificatListe.Rows.ForEach({$DataGridViewCertificatListe.Rows.Remove($_)}) # Duplication de la ligne pour supprimer la ligne restante apr�s premi�re passe
    $CheckBoxExportClef.Enabled = $false
    $CertificatsSHA1.Clear()
    $CertificatsExpiration.Clear()
    $CertificatsClef.Clear()
    $CertificatsEmetteur.Clear()
    $CertificatsObjet.Clear()
    $OrdreTriCertificat.Clear()
    $CertificatNiveau.Clear()
}

# V�rification que le mot de passe fourni est bien le bon pour le certificat choisi
function TestMotDePasse {
    certutil.exe -p $TextBoxMotDePasse.Text -dump $labelEmplacementCertificat.Text | Out-Null
    return $?
}

# Compl�te la DataGridView avec les diff�rents certificats trouv�s
function CompleteDataGridViewCertificats {
    $SortieCertutil = certutil.exe -p $TextBoxMotDePasse.Text -dump $labelEmplacementCertificat.Text
    
    $SortieCertutil | Select-String -CaseSensitive "sha1" | ForEach-Object { $CertificatsSHA1.Add($_.ToString().Split(" ")[-1]) }
    $SortieCertutil | Select-String -CaseSensitive "After" | ForEach-Object { $CertificatsExpiration.Add($_.ToString().Substring(12)) }
    
    $switch = $false
    $SortieCertutil.ForEach({
        if ($swicth) { 
            if ($_ -match "^ ") { $CertificatsClef.Add("OUI") ; $CheckBoxExportClef.Enabled = $true }
                           else { $CertificatsClef.Add("NON") }
        }
        $swicth = ($_ -match "^----------------")
    })

    $switch = $true
    $SortieCertutil | Select-String -CaseSensitive "CN=" | ForEach-Object { 
        $temp = $_.ToString().Split(":")[1].Trim()
        if ($switch) { $CertificatsEmetteur.Add($temp) }
                else { $CertificatsObjet.Add($temp) }
        $switch = !$switch
    }

    # Recherche d'un certificat racine et classement des certificats tels que donn�s par certutil
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
        $DataGridViewCertificatListe.Rows.Add($CertificatNiveau[$i], $CertificatsClef[$_], $CertificatsExpiration[$_], $CertificatsObjet[$_], $CertificatsEmetteur[$_])
        $i++
    })
}


# ---------------------------------------------------- Gestion GUI ---------------------------------------------------


$fen�treCertificateExtractor = New-Object System.Windows.Forms.Form -Property @{
    Height = $HauteurFen�tre
    Width = $LargeurFen�tre
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
    Width = $LargeurFen�tre - 30
    Padding = 0
    Font = "Segoe UI, 10"
}
$fen�treCertificateExtractor.Controls.Add($GroupBoxCertificat)

$GroupBoxMotDePasse = New-Object System.Windows.Forms.GroupBox -Property @{
    Text = "2 - Mot de passe du certificat : "
    Location = "10, 120"
    Height = 100
    Width = $LargeurFen�tre - 30
    Padding = 0
    Font = "Segoe UI, 10"
    Enabled = $False
}
$fen�treCertificateExtractor.Controls.Add($GroupBoxMotDePasse)

$GroupBoxCertificatListe = New-Object System.Windows.Forms.GroupBox -Property @{
    Text = "3 - Liste des certificats contenus : "
    Location = "10, 230"
    Height = 220
    Width = $LargeurFen�tre - 30
    Padding = 0
    Font = "Segoe UI, 10"
    Enabled = $false
}
$fen�treCertificateExtractor.Controls.Add($GroupBoxCertificatListe)

# -------------------- Contenu de la GroupBox n�1 : recherche du certificat

$buttonS�lectionCertificat = New-Object System.Windows.Forms.Button -Property @{
    Text = "Ouvrir"
    Location = "20, 38"
    Font = $Font
    Autosize = $true
}
$GroupBoxCertificat.Controls.Add($buttonS�lectionCertificat)

$labelEmplacementCertificat = New-Object System.Windows.Forms.Label -Property @{
    Text = "Pas de certificat s�lectionn�"
    TextAlign = "MiddleLeft"
    Location = "120, 20"
    Width = ($LargeurFen�tre-170)
    Height = 70
    Font = $Font
}
$GroupBoxCertificat.Controls.Add($labelEmplacementCertificat)

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
	Filter = 'Certificat (*.pfx ou *.p12)|*.pfx;*.p12'
    FileName = ""
    Title = "Choix du certificat pour Certificate Extractor"
}

# -------------------- Contenu de la GroupBox n�2 : Sp�cification du mot de passe du certificat

$TextBoxMotDePasse = New-Object System.Windows.Forms.TextBox -Property @{
    Location = "10, 20"
    ScrollBars = 0
    Multiline = $False
    Enabled = $True
    ReadOnly = $False
    Font = $Font
    Width = $LargeurFen�tre - 50
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
    Location = (""+($LargeurFen�tre - 114)+", 55")
    Font = $Font
    Autosize = $true
}
$GroupBoxMotDePasse.Controls.Add($ButtonValidationMotDePasse)

# -------------------- Contenu de la GroupBox n�3 : Visualisation des certificats et export

$DataGridViewCertificatListe = New-Object System.Windows.Forms.DataGridView -Property @{
    Width = $LargeurFen�tre - 50
    Height = 150
    Location = "10, 20"
    Font = $FontMin
    ColumnCount = 5
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

# D�finition des colonnes du DataGridView
$DataGridViewCertificatListe.Columns[0].Name = "Niveau"
$DataGridViewCertificatListe.Columns[1].Name = "Cl� associ�e"
$DataGridViewCertificatListe.Columns[2].Name = "Expiration"
$DataGridViewCertificatListe.Columns[3].Name = "Objet"
$DataGridViewCertificatListe.Columns[4].Name = "Emetteur"

# Emp�chement de trier les diff�rentes colonnes
$DataGridViewCertificatListe.Columns.ForEach({ $_.SortMode = 0 })

$CheckBoxExportClef = New-Object System.Windows.Forms.CheckBox -Property @{
    Font = $FontMin
    Text = 'Exporter le certificat avec cl� associ�e au format p12 sans mot de passe'
    Location = "10, 175"
    Checked = $false
    #Autosize = $true
    Enabled = $false
    Width = 250
    Height = 40
}
$GroupBoxCertificatListe.Controls.Add($CheckBoxExportClef)

$ButtonExportSelection = New-Object System.Windows.Forms.Button -Property @{
    Text = "Exporter la s�lection"
    Location = (""+($LargeurFen�tre - 280)+", 178")
    Font = $Font
    Autosize = $true
}
$GroupBoxCertificatListe.Controls.Add($ButtonExportSelection)

$ButtonExport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Exporter"
    Location = (""+($LargeurFen�tre - 114)+", 178")
    Font = $Font
    Autosize = $true
}
$GroupBoxCertificatListe.Controls.Add($ButtonExport)


# ----------------------------------------------- Ev�nements graphiques ----------------------------------------------


$buttonS�lectionCertificat.add_Click({
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
    } else {
        [System.Windows.Forms.MessageBox]::Show("Le mot de passe indiqu� n'est pas bon !",'Erreur','OK','Error')
    }
})


# ---------------------------------------------- Affichage de la fen�tre ---------------------------------------------


$fen�treCertificateExtractor.ShowDialog()