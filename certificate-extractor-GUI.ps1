<#

.SYNOPSIS
Script Powershell GUI permettant de facilement pouvoir extraire tous les éléments ou seulement certains des pfx/p12/p7b/pem/crt/cer/txt.

.DESCRIPTION

Script PowerShell fonctionnant avec une interface graphique permettant de manipuler les certificats.

Cet outil a comme objectif de :
    - simplifier l'extraction des autorités de certification intermédiaires et racine pour les formats de certificat standards
    - simplifier la transformation d'un p12/pfx vers des crt

Cet outil est un substitut plus pratique et simple que les commandes OpenSSL ou bien d'avoir à importer les p12/pfx dans le magasin de certificat pour pouvoir extraire ses certificats.

/!\ Le fonctionnement de l'outil s'appuie sur la commande certutil.exe qui limite le support à Windows ainsi que le magasin de certificat de Windows.

.NOTE
    Version : 2.0
    Auteur  : Lucas MEYER
    Github  : https://github.com/Meyer-Lucas/certificate-extractor-GUI
    Licence : MIT License

#>

# Préparation du script pour une utilisation graphique
Add-Type -AssemblyName system.windows.forms
[System.Windows.Forms.Application]::EnableVisualStyles()


# -------------------------------------------------- Variables globales ----------------------------------------------


[String]$Font           = "Segoe UI, 11"
[String]$FontMin        = "Segoe UI, 8"
[Int]$LargeurFenêtre    = 600
[Int]$HauteurFenêtre    = 495

$CertificatsSHA1        = New-Object System.Collections.ArrayList
$CertificatsExpiration  = New-Object System.Collections.ArrayList
$CertificatsEmetteur    = New-Object System.Collections.ArrayList
$CertificatsObjet       = New-Object System.Collections.ArrayList
$OrdreTriCertificat     = New-Object System.Collections.ArrayList
$CertificatNiveau       = New-Object System.Collections.ArrayList


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
}

# Vérification que le mot de passe fourni est bien le bon pour le certificat choisi
function TestMotDePasse {
    certutil.exe -p $TextBoxMotDePasse.Text -dump $labelEmplacementCertificat.Text | Out-Null
    return $?
}

# Parsing des pfx/p12
function ParsingPfxP12 {
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
    
    certutil.exe -p $TextBoxMotDePasse.Text -dump -split -silent $labelEmplacementCertificat.Text
}

# Parsing de la sortie de certutil pour tous les autres types de certificats
function ParsingCertsCertutil {
    param($SortieCertutil)

    $objet = ""
    $emetteur = ""

    # Cette variable va retenir le numéro du bloc où il y a 4 espaces en début, le 2ème bloc correspond à l'émetteur tandis que le 3ème correspond à l'objet
    $BlocsEspaces = 0
    # Variable pour retenir si l'on parse un bloc avec 4 espaces
    $InterieurBloc = $false

    $SortieCertutil.ForEach({
        if ($_ -match "^    ") {
            if (!$InterieurBloc) { $BlocsEspaces++ }
            $InterieurBloc = $true
        } else { $InterieurBloc = $false }
        
        if ($InterieurBloc -and $BlocsEspaces -eq 2) { $emetteur += $_.Trim() + ", " }
        if ($InterieurBloc -and $BlocsEspaces -eq 3) { $objet += $_.Trim() + ", " }
    })

    $CertificatsObjet.Add($objet.Substring(0,$objet.Length-2))
    $CertificatsEmetteur.Add($emetteur.Substring(0,$emetteur.Length-2))
    $CertificatsExpiration.Add(($SortieCertutil -match "NotAfter")[0].Substring(($SortieCertutil -match "NotAfter")[0].IndexOf(":")+1).Trim())
    $CertificatsSHA1.Add(($SortieCertutil -match "(sha1)")[-1].Split(" ")[-1])
}

# Parsing des p7b
function ParsingP7b {
    $SortieCertutil = certutil.exe -dump -split $labelEmplacementCertificat.Text

    $DebutSectionCertificat = New-Object System.Collections.ArrayList
    $FinSectionCertificat = New-Object System.Collections.ArrayList

    for ($i = 0; $i -lt $SortieCertutil.Length; $i++) {
        if ($SortieCertutil[$i] -match "================") { $DebutSectionCertificat.Add($i) }
        if ($SortieCertutil[$i] -match "----------------") { $FinSectionCertificat.Add($i) }
    }

    for ($i = 0; $i -lt $DebutSectionCertificat.Count; $i++) { ParsingCertsCertutil -SortieCertutil ($SortieCertutil[($DebutSectionCertificat[$i]+1)..($FinSectionCertificat[$i]-1)]) }
}

# Parsing des autres certificats, vérifie si c'est encodé en DER et si ce n'est pas le cas une identification de la présence de plusieurs certificats a lieu
function ParsingCerts {
    $Certificat = Get-Content $OpenFileDialog.FileName

    # Si le certificat sélectionné est encodé en base64
    if ($Certificat -match "-----BEGIN CERTIFICATE-----") {
        # Identification des possibles différents certificats contenus dans le fichier
        # Vérification si le fichier ne contient qu'une seule ligne (l'extraction sera différente suivant le cas)
        if ($Certificat.Count -eq 1) {
            # Recherche et isolement des différents certificats pour ensuite récupérer les informations de chacuns d'entre eux
            while ($Certificat.IndexOf("------") -gt 0) {
                $Jonction = $Certificat.IndexOf("------") + 5
                $Certificat.Substring(0, $Jonction) | Out-File -FilePath certTemp.pem
                ParsingCertsCertutil -SortieCertutil (certutil.exe -dump -split certTemp.pem)
                Remove-Item certTemp.pem
                $Certificat = $Certificat.Substring($Jonction)
            }
            $Certificat | Out-File -FilePath certTemp.pem
            ParsingCertsCertutil -SortieCertutil (certutil.exe -dump -split certTemp.pem)
            Remove-Item certTemp.pem
        } else {
            $DebutCert = New-Object System.Collections.ArrayList
            $FinCert = New-Object System.Collections.ArrayList
            # Recherche de l'emplacement des différents certificats
            for ($i = 0; $i -lt $Certificat.Count; $i++) {
                if ($Certificat[$i] -match "-----BEGIN CERTIFICATE-----") { $DebutCert.Add($i) }
                if ($Certificat[$i] -match "-----END CERTIFICATE-----") { $FinCert.Add($i) }
            }
            # Obtention des différents certificats
            for ($i = 0; $i -lt $DebutCert.Count; $i++) {
                $Certificat[$DebutCert[$i]..$FinCert[$i]] | Out-File -FilePath certTemp.pem
                ParsingCertsCertutil -SortieCertutil (certutil.exe -dump -split certTemp.pem)
                Remove-Item certTemp.pem
            }
        }
    } else { ParsingCertsCertutil -SortieCertutil (certutil.exe -dump -split $OpenFileDialog.FileName) }
}

# Recherche de la présence d'un certificat racine
function PresenceCertificatRacine {
    $sortie = $false

    for ($i = 0 ; $i -lt $CertificatsEmetteur.Count ; $i++) {
        if ($CertificatsEmetteur[$i] -eq $CertificatsObjet[$i]) { $sortie = $true }
    }

    return $sortie
}

# Classement des certificats par rapport aux émetteurs et objets
function TrieDesCertificats {
    $OrdreTriCertificat.Clear()
    $OrdreTriCertificat.Add(0)
    $AntiInfiniteLoop = 30
    while ($CertificatsEmetteur.Count -gt 1 -and $OrdreTriCertificat.Count -lt $CertificatsEmetteur.Count -and $AntiInfiniteLoop) {
        for ($i = 1; $i -lt $CertificatsEmetteur.Count; $i++) {
            # Si le certificat actuel a un émetteur qui correspond à l'objet du certificat le plus bas dans le tri temporaire et qu'il n'a pas déjà été ajouté à la liste
            if ($CertificatsEmetteur[$i] -eq $CertificatsObjet[$OrdreTriCertificat[-1]] -and $OrdreTriCertificat -notcontains $i) { $OrdreTriCertificat.Add($i) }
            # Sinon si l'objet du certificat actuel correspond à l'émetteur du premier certificat dans le tri temporaire et qu'il n'a pas déjà été ajouté à la liste
            elseif ($CertificatsObjet[$i] -eq $CertificatsEmetteur[$OrdreTriCertificat[0]] -and $Script:OrdreTriCertificat -notcontains $i) { $OrdreTriCertificat.Insert(0,$i) }
        }
        $AntiInfiniteLoop--
    }

    # Définition des niveaux si un certificat racine est présent
    $CertificatNiveau.Clear()
    if (PresenceCertificatRacine) { 
        for ($i = 0; $i -lt $CertificatsEmetteur.Count; $i++) { $CertificatNiveau.Add($i) }
    } else {
        for ($i = 0; $i -lt $CertificatsEmetteur.Count; $i++) { $CertificatNiveau.Add("?") }
    }
}

# Complète la chaîne de certificatation en se basant sur les certificats contenus dans le magasin de certification
function AjoutChaineDeCertification {
    $reussi = $false

    $CACert = ((Get-ChildItem -Recurse Cert:\) -match $CertificatsEmetteur[$OrdreTriCertificat[0]])[0]

    if ($CACert -ne $null) {
        $reussi = $true
        $CertificatsObjet.add($CACert.Subject)
        $CertificatsEmetteur.Add($CACert.Issuer)
        $CertificatsSHA1.Add($CACert.Thumbprint)
        $CertificatsExpiration.Add($CACert.NotAfter)
        Export-Certificate -Cert $CACert -Type CERT -FilePath (""+$CACert.Thumbprint+".crt")
    }

    return $reussi
}

# Complète la DataGridView avec les différents certificats trouvés
function CompleteDataGridViewCertificats {
    switch -Regex ($OpenFileDialog.FileName) {
        ".*\.p(12|fx)$" { ParsingPfxP12 }
        ".*\.p7b$" { ParsingP7b }
        default { ParsingCerts }
        #default { ParsingCertsCertutil -SortieCertutil (certutil.exe -dump $OpenFileDialog.FileName) }
    }
    
    if ($CertificatsSHA1.Count -ne 0) {
        do {
            TrieDesCertificats
            $boucle = $false
            if (!(PresenceCertificatRacine)) { $boucle = AjoutChaineDeCertification }
        } while ($boucle)
    } else { [System.Windows.Forms.MessageBox]::Show("Le certificat sélectionné ne semble pas être correct !",'Erreur','OK','Error') }
    
    # Ajout des certificats dans le DataGridView
    $i = 0
    $OrdreTriCertificat.ForEach({
        $DataGridViewCertificatListe.Rows.Add($CertificatNiveau[$i], $CertificatsExpiration[$_], $CertificatsObjet[$_], $CertificatsEmetteur[$_])
        $i++
    })
}

# Fonction chargée de renommer les certificats qui ont étés sélectionnés et de les exporter
function RenommeCertificatExport {
    param([string]$RepertoireCible, $CertificatsSHA1Export = $CertificatsSHA1)
    $Certificats = Get-ChildItem -File $RepertoireTemp.FullName
    $Certificats.ForEach({ 
        for ($i = 0 ; $i -lt $CertificatsSHA1.Count ; $i++) {
            if ($_.ToString() -match $CertificatsSHA1[$i] -and $CertificatsSHA1Export -contains $CertificatsSHA1[$i]) {
                $Nom = $CertificatsObjet[$i].Substring($CertificatsObjet[$i].IndexOf("CN=")).Split(",")[0].Substring(3).Replace("*","_")
                if ($CertificatsObjet[$i] -match "O=") { $Nom += " (" + $CertificatsObjet[$i].Substring($CertificatsObjet[$i].IndexOf("O=")+2).Split(",")[0].Replace("*","_") + ")" }
                if ($CertificatNiveau[$OrdreTriCertificat[$i]] -ne "?") { $Nom = $CertificatNiveau[$OrdreTriCertificat[$i]].ToString() + " - "  + $Nom }
                $Nom += ".crt"
                Copy-Item -Force $_.FullName -Destination ($RepertoireCible + "\" + $Nom)
            }
        }
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
	Filter = 'Certificat (*.pfx; *.p12; *.p7b; *.pem; *.crt; *.cer; *.txt)|*.pfx;*.p12;*.p7b;*.pem;*.crt;*.cer;*.txt'
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
        
        if ($OpenFileDialog.FileName -match ".*\.p(12|fx)$") {
            $GroupBoxMotDePasse.Enabled = $true
            $TextBoxMotDePasse.Text = ""
            ResetGroupBoxCertificatListe
            $TextBoxMotDePasse.Focus()
        } else {
            $GroupBoxMotDePasse.Enabled = $false
            $TextBoxMotDePasse.Text = ""
            ResetGroupBoxCertificatListe
            CompleteDataGridViewCertificats
            $GroupBoxCertificatListe.Enabled = $true
            $ButtonExport.Focus()
        }
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
        $CertificatsSHA1Selectionne = New-Object System.Collections.ArrayList
        $DataGridViewCertificatListe.SelectedRows.ForEach({ $CertificatsSHA1Selectionne.Add($CertificatsSHA1[$CertificatsObjet.IndexOf($_.Cells.Item("Objet").Value)]) })

        RenommeCertificatExport -RepertoireCible $Repertoire -CertificatsSHA1Export $CertificatsSHA1Selectionne

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
        RenommeCertificatExport -RepertoireCible $Repertoire

        [System.Windows.Forms.MessageBox]::Show("L'export de tous les certificats est terminée.", "Tâche terminée", "OK","Info")
    }
})


# ---------------------------------------------- Affichage de la fenêtre ---------------------------------------------


$RepertoireTemp = New-Item -ItemType Directory ($env:TEMP+"\Certificate-Extractor-" + (Get-Random -Maximum 99999))
Set-Location $RepertoireTemp.FullName

$fenêtreCertificateExtractor.ShowDialog() | Out-Null

Set-Location $env:TEMP
Remove-Item -Recurse -Path $RepertoireTemp.FullName