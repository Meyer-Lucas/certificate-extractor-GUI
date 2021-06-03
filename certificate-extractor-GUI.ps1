<#.SYNOPSIS
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
    Auteur : Lucas MEYER
    Github : https://github.com/Meyer-Lucas
    Licence : MIT License

#>

# Pr�paration du script pour une utilisation graphique
Add-Type -AssemblyName system.windows.forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# -------------------------------------------------- Variables globales ----------------------------------------------


[String]$Font = "Segoe UI, 11"
[Int]$LargeurFen�tre = 500
[Int]$HauteurFen�tre = 500


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

$buttonS�lectionCertificat = New-Object System.Windows.Forms.Button -Property @{
    Text = "Choisir un certificat"
    Location = "10, 10"
    Font = $Font
    Autosize = $true
}
$fen�treCertificateExtractor.Controls.Add($buttonS�lectionCertificat)

$labelNomCertificat = New-Object System.Windows.Forms.Label -Property @{
    Text = "Pas de certificat s�lectionn�"
    TextAlign = "MiddleLeft"
    Location = "160, 15"
    Width = ($LargeurFen�tre-170)
    Font = $Font
}
$fen�treCertificateExtractor.Controls.Add($labelNomCertificat)

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
	Filter = 'Certificats (*.pfx;*.p12;*.cer;*.crt;*.pem;*.txt)|*.pfx;*.p12;*.cer;*.crt;*.pem;*.txt'
    FileName = ""
    Title = "Choix du certificat pour Certificate Extractor"
}

$GroupBoxMotDePasse = New-Object System.Windows.Forms.GroupBox -Property @{
    Text = "Mot de passe du certificat :"
    Location = "10, 50"
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
    Width = 450
    Height = 30
    PasswordChar = "*"
}
$GroupBoxMotDePasse.Controls.Add($TextBoxMotDePasse)

$CheckBoxMotDePasseClair= New-Object System.Windows.Forms.CheckBox -Property @{
    Font = $Font
    Text = 'Afficher le mot de passe'
    Location = "10, 60"
    Checked = $false
    Width = 200
}
$GroupBoxMotDePasse.Controls.Add($CheckBoxMotDePasseClair)

$ButtonValidationMotDePasse = New-Object System.Windows.Forms.Button -Property @{
    Text = "Valider"
    Location = "386, 55"
    Font = $Font
    Autosize = $true
}
$GroupBoxMotDePasse.Controls.Add($ButtonValidationMotDePasse)


# ----------------------------------------------- Ev�nements graphiques ----------------------------------------------


$buttonS�lectionCertificat.add_Click({
    $OpenFileDialog.ShowDialog()
    if ($OpenFileDialog.FileName.Length -gt 0) { 
        $labelNomCertificat.Text = $OpenFileDialog.FileName.Split("\")[-1] 
        $GroupBoxMotDePasse.Enabled = $true
    }
})

$CheckBoxMotDePasseClair.add_click({
    if ($CheckBoxMotDePasseClair.Checked) { $TextBoxMotDePasse.PasswordChar = 0 }
    else { $TextBoxMotDePasse.PasswordChar = "*" }
})


# ---------------------------------------------- Affichage de la fen�tre ---------------------------------------------


$fen�treCertificateExtractor.ShowDialog()