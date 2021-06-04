<#

.SYNOPSIS
Script Powershell GUI permettant de facilement pouvoir extraire tous les �l�ments ou seulement certains des pfx/p12 ou bien des certificats (cer/pem/crt).

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
[Int]$LargeurFen�tre = 900
[Int]$HauteurFen�tre = 495
[String]$EmplacementCertificat = ""


# ----------------------------------------------------- Fonctions ----------------------------------------------------


function ResetGroupBoxCertificatListe {
    $GroupBoxCertificatListe.Enabled = $false
    $DataGridViewCertificatListe.Rows.ForEach({$DataGridViewCertificatListe.Rows.Remove($_)})
    $CheckBoxExportClef.Enabled = $false
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

$GroupBoxCertificat = New-Object System.Windows.Forms.GroupBox -Property @{
    Text = "1 - Choix du certificat : "
    Location = "10, 10"
    Height = 100
    Width = $LargeurFen�tre - 30
    Padding = 0
    Font = "Segoe UI, 10"
}
$fen�treCertificateExtractor.Controls.Add($GroupBoxCertificat)

$buttonS�lectionCertificat = New-Object System.Windows.Forms.Button -Property @{
    Text = "Ouvrir"
    Location = "20, 38"
    Font = $Font
    Autosize = $true
}
$GroupBoxCertificat.Controls.Add($buttonS�lectionCertificat)

$labelNomCertificat = New-Object System.Windows.Forms.Label -Property @{
    Text = "Pas de certificat s�lectionn�"
    TextAlign = "MiddleLeft"
    Location = "120, 20"
    Width = ($LargeurFen�tre-170)
    Height = 70
    Font = $Font
}
$GroupBoxCertificat.Controls.Add($labelNomCertificat)

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
	Filter = 'Certificat (*.pfx;*.p12;*.cer;*.crt;*.pem;*.txt)|*.pfx;*.p12;*.cer;*.crt;*.pem;*.txt'
    FileName = ""
    Title = "Choix du certificat pour Certificate Extractor"
}

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

$DataGridViewCertificatListe.Columns[0].Name = "Niveau"
$DataGridViewCertificatListe.Columns[1].Name = "Cl� associ�e"
$DataGridViewCertificatListe.Columns[2].Name = "Expiration"
$DataGridViewCertificatListe.Columns[3].Name = "Objet"
$DataGridViewCertificatListe.Columns[4].Name = "Emetteur"

$DataGridViewCertificatListe.Columns.ForEach({ $_.SortMode = 0 })

$CheckBoxExportClef = New-Object System.Windows.Forms.CheckBox -Property @{
    Font = $Font
    Text = 'Exporter le certificat avec cl� associ�e au format p12 sans mot de passe'
    Location = "10, 182"
    Checked = $false
    Autosize = $true
    Enabled = $false
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
    $OpenFileDialog.ShowDialog()
    if ($OpenFileDialog.FileName.Length -gt 0) {
        $labelNomCertificat.Text = $OpenFileDialog.FileName
        $GroupBoxMotDePasse.Enabled = $true
        $TextBoxMotDePasse.Text = ""
        ResetGroupBoxCertificatListe
    }
})

$TextBoxMotDePasse.add_TextChanged({
    ResetGroupBoxCertificatListe
})

$CheckBoxMotDePasseClair.add_click({
    if ($CheckBoxMotDePasseClair.Checked) { $TextBoxMotDePasse.PasswordChar = 0 }
    else { $TextBoxMotDePasse.PasswordChar = "*" }
})


# ---------------------------------------------- Affichage de la fen�tre ---------------------------------------------


$fen�treCertificateExtractor.ShowDialog()