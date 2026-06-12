<#
================================================================================
                 ENVOI DE MASSE - QR CODES TOKENS PAR EMAIL
                  (avec prévisualisation avant chaque envoi)
================================================================================

  OBJECTIF
  --------
  À partir du CSV enrichi produit par le générateur de QR codes
  (colonnes : Nom ; Email ; Token ; CheminQR), ce script :
      - identifie l'adresse email de chaque ligne
      - récupère l'image QR correspondante (colonne CheminQR)
      - l'injecte dans le corps du mail (image INLINE, pas une simple PJ)
      - affiche une PRÉVISUALISATION fidèle de chaque mail
      - n'envoie qu'après validation (mail par mail, ou mode auto)

  SCHÉMA DE FONCTIONNEMENT
  ------------------------

      ┌──────────────────────────┐
      │  CSV enrichi  (;)        │   Nom ; Email ; ... ; CheminQR
      └────────────┬─────────────┘
                   │  [GUI] chargement + mapping colonnes Email / Nom / CheminQR
                   ▼
      ┌──────────────────────────┐
      │  Paramètres SMTP         │   serveur / port / SSL / expéditeur
      │  (laissés VIERGES,       │   + identifiants optionnels
      │   à renseigner)          │
      └────────────┬─────────────┘
                   ▼
      ┌──────────────────────────┐
      │  File d'attente          │   1 destinataire = 1 mail
      └────────────┬─────────────┘
                   ▼
      ┌──────────────────────────┐     [Envoyer]      ┌──────────────────┐
      │  PRÉVISUALISATION        │───────────────────▶│  SMTP (inline    │
      │  destinataire / objet /  │     [Passer]       │  cid: QR code)   │
      │  corps HTML + QR affiché │──────┐             └──────────────────┘
      └────────────┬─────────────┘      │
                   │ [Tout envoyer]     ▼
                   ▼              destinataire suivant
      ┌──────────────────────────┐
      │  Journal + bilan final   │   OK / échec / passés
      └──────────────────────────┘

  MODES D'ENVOI
  -------------
  - "Envoyer"            : envoie le mail affiché, passe au suivant
  - "Passer"             : ignore le destinataire affiché, passe au suivant
  - "Tout envoyer"       : envoie tous les mails RESTANTS sans nouvelle
                           confirmation (à utiliser après avoir validé
                           quelques prévisualisations)

  POINTS DE SÉCURITÉ
  ------------------
  - Le QR est injecté en image inline (Content-ID) : le secret ne transite
    que vers le destinataire concerné, une seule image par mail.
  - Le mot de passe SMTP éventuel est saisi en champ masqué et converti en
    PSCredential ; il n'est jamais écrit sur disque ni dans le journal.
  - Anti-erreur d'aiguillage : contrôle de cohérence avant chaque envoi
    (email non vide + fichier QR existant), sinon mise en échec et passage
    au suivant.
  - Pause configurable entre deux envois (mode "Tout envoyer") pour ne pas
    déclencher les seuils anti-spam / throttling du serveur.

  PRÉREQUIS
  ---------
  - Windows PowerShell 5.1 ou PowerShell 7+ (Windows)
  - Un serveur SMTP joignable et un compte autorisé à émettre
    (relais interne, connecteur dédié, etc.)
  - Le CSV enrichi ET les images QR encore présents sur le disque
    (ne pas purger avant la fin de la campagne !)
================================================================================
#>

# ==============================================================================
# CONFIGURATION - MODÈLE DU MAIL (personnalisable)
# ==============================================================================

# Chemin de l'IMAGE DE SIGNATURE (logo, bannière de service, signature scannée...)
# - Renseignez un chemin complet, ex : "C:\Outils\signature.png"
# - Laissez vide ("") pour ne pas inclure de signature.
# L'image est injectée en INLINE (cid:) comme le QR : elle s'affiche dans le
# corps du mail sans apparaître comme pièce jointe à ouvrir.
$CheminSignature = ""

# Largeur d'affichage de la signature dans le mail (en pixels)
$LargeurSignaturePx = 300

# Objet du mail. {NOM} sera remplacé par le nom du destinataire.
$ModeleObjet = "Votre code d'activation personnel"

# Corps HTML. Variables disponibles :
#   {NOM}        -> nom du destinataire
#   {QR}         -> emplacement où l'image QR est injectée (NE PAS SUPPRIMER)
#   {SIGNATURE}  -> emplacement de l'image de signature (retirée automatiquement
#                   si $CheminSignature est vide ou si le fichier est introuvable)
$ModeleCorps = @"
<html>
<body style="font-family: Segoe UI, Arial, sans-serif; color:#222; max-width:600px;">
  <p>Bonjour {NOM},</p>
  <p>Veuillez trouver ci-dessous votre code QR personnel d'activation.</p>
  <p>Scannez-le avec l'application prévue à cet effet :</p>
  <p style="text-align:center; margin:25px 0;">{QR}</p>
  <p>Ce code est strictement personnel : ne le transférez à personne.</p>
  <p>En cas de difficulté, contactez le support informatique.</p>
  <p>Cordialement,<br/>Le support informatique</p>
  <p style="margin-top:20px;">{SIGNATURE}</p>
</body>
</html>
"@

# Pause entre deux envois en mode "Tout envoyer" (en millisecondes)
$PauseEntreEnvoisMs = 1500

# ==============================================================================
# CHARGEMENT DES ASSEMBLIES GUI
# ==============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==============================================================================
# 1. FONCTIONS
# ==============================================================================

function Send-QRMail {
    <#
        Envoie UN mail avec le QR code en image inline (Content-ID).
        Utilise System.Net.Mail.SmtpClient (et non Send-MailMessage) car
        c'est le seul moyen propre d'embarquer une image inline via cid:.

        Retourne $true si l'envoi a réussi, sinon lève une exception
        (interceptée par l'appelant pour journalisation).
    #>
    param(
        [string]$Serveur,
        [int]$Port,
        [bool]$UtiliserSSL,
        [pscredential]$Credential,   # $null = authentification Windows intégrée (relais AD)
        [string]$Expediteur,
        [string]$Destinataire,
        [string]$Objet,
        [string]$CorpsHtml,          # doit contenir le marqueur {QR}
        [string]$CheminQR
    )

    # --- Construction du corps : {QR} devient <img cid:...>, {SIGNATURE} idem ---
    # IMPORTANT : Content-ID UNIQUE par mail (GUID). Avec un CID fixe, certains
    # clients (Outlook en vue conversation) mettent l'image en cache et
    # réaffichent le PREMIER QR reçu dans tous les mails suivants.
    $cidQR  = "qr-"  + [guid]::NewGuid().ToString("N")
    $cidSig = "sig-" + [guid]::NewGuid().ToString("N")
    $html   = $CorpsHtml -replace '\{QR\}', "<img src=""cid:$cidQR"" alt=""QR code"" style=""width:280px;height:280px;"" />"

    # Signature : injectée seulement si le fichier existe, sinon le marqueur est retiré
    $sigOk  = (-not [string]::IsNullOrWhiteSpace($script:CheminSignature)) -and (Test-Path $script:CheminSignature)
    if ($sigOk) {
        $html = $html -replace '\{SIGNATURE\}', "<img src=""cid:$cidSig"" alt=""Signature"" style=""width:$($script:LargeurSignaturePx)px;"" />"
    } else {
        $html = $html -replace '\{SIGNATURE\}', ""
    }

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From       = New-Object System.Net.Mail.MailAddress($Expediteur)
    $mail.Subject    = $Objet
    $mail.IsBodyHtml = $true
    $mail.To.Add($Destinataire) | Out-Null

    # --- Vue HTML + ressources liées (QR + signature, référencées par Content-ID) ---
    $vue = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString(
        $html, $null, [System.Net.Mime.MediaTypeNames+Text]::Html)

    $ressource = New-Object System.Net.Mail.LinkedResource($CheminQR, "image/png")
    $ressource.ContentId = $cidQR
    $ressource.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
    $vue.LinkedResources.Add($ressource)

    $ressourceSig = $null
    if ($sigOk) {
        # Type MIME déduit de l'extension (png/jpg/gif pris en charge)
        $mime = switch ([IO.Path]::GetExtension($script:CheminSignature).ToLower()) {
            '.jpg'  { 'image/jpeg' }
            '.jpeg' { 'image/jpeg' }
            '.gif'  { 'image/gif'  }
            default { 'image/png'  }
        }
        $ressourceSig = New-Object System.Net.Mail.LinkedResource($script:CheminSignature, $mime)
        $ressourceSig.ContentId = $cidSig
        $ressourceSig.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
        $vue.LinkedResources.Add($ressourceSig)
    }

    $mail.AlternateViews.Add($vue)

    # --- Client SMTP ---
    $smtp = New-Object System.Net.Mail.SmtpClient($Serveur, $Port)
    $smtp.EnableSsl = $UtiliserSSL
    if ($Credential) {
        # Compte explicite (ex: compte de service)
        $smtp.Credentials = $Credential.GetNetworkCredential()
    } else {
        # Authentification Windows intégrée : utilise le compte AD de la session
        # (cas typique d'un relais SMTP interne autorisant le compte machine/utilisateur)
        $smtp.UseDefaultCredentials = $true
    }

    try {
        $smtp.Send($mail)
        return $true
    }
    finally {
        # Libération systématique des ressources (fichiers verrouillés sinon)
        $ressource.Dispose()
        if ($ressourceSig) { $ressourceSig.Dispose() }
        $vue.Dispose()
        $mail.Dispose()
        $smtp.Dispose()
    }
}

function Find-Column {
    <# Auto-détection d'une colonne par liste de motifs regex (ordre = priorité). #>
    param([string[]]$Colonnes, [string[]]$MotsCles)
    foreach ($mc in $MotsCles) {
        $match = $Colonnes | Where-Object { $_ -match $mc } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

# ==============================================================================
# 2. CONSTRUCTION DE L'INTERFACE
# ==============================================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Envoi de masse - QR codes par email"
$form.Size            = New-Object System.Drawing.Size(900, 720)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# ------------------------------------------------------------------ CSV enrichi
$lblCsv = New-Object System.Windows.Forms.Label
$lblCsv.Text = "1. CSV enrichi (avec colonne CheminQR) :"
$lblCsv.Location = '15,12'; $lblCsv.AutoSize = $true
$form.Controls.Add($lblCsv)

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = '15,33'; $txtCsv.Size = '650,25'; $txtCsv.ReadOnly = $true
$form.Controls.Add($txtCsv)

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = "Parcourir..."
$btnCsv.Location = '675,31'; $btnCsv.Size = '95,27'
$form.Controls.Add($btnCsv)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Aucun fichier chargé."
$lblCount.Location = '780,36'; $lblCount.AutoSize = $true
$lblCount.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblCount)

# ------------------------------------------------------------- Mapping colonnes
$grpMap = New-Object System.Windows.Forms.GroupBox
$grpMap.Text = "2. Mapping des colonnes"
$grpMap.Location = '15,65'; $grpMap.Size = '420,85'
$form.Controls.Add($grpMap)

$combos = @{}
$mapDefs = @(
    @{ Cle = 'Email';    Libelle = 'Email :';      X = 10  },
    @{ Cle = 'Nom';      Libelle = 'Nom :';        X = 145 },
    @{ Cle = 'CheminQR'; Libelle = 'Chemin QR :';  X = 280 }
)
foreach ($d in $mapDefs) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $d.Libelle; $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point($d.X, 22)
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.DropDownStyle = 'DropDownList'; $cmb.Size = '125,25'
    $cmb.Location = New-Object System.Drawing.Point($d.X, 44)
    $grpMap.Controls.Add($lbl); $grpMap.Controls.Add($cmb)
    $combos[$d.Cle] = $cmb
}

# -------------------------------------------------------------- Paramètres SMTP
$grpSmtp = New-Object System.Windows.Forms.GroupBox
$grpSmtp.Text = "3. Paramètres SMTP (à renseigner)"
$grpSmtp.Location = '445,65'; $grpSmtp.Size = '435,150'
$form.Controls.Add($grpSmtp)

# Serveur SMTP -- LAISSÉ VIERGE volontairement
$lblSrv = New-Object System.Windows.Forms.Label
$lblSrv.Text = "Serveur :"; $lblSrv.Location = '10,25'; $lblSrv.AutoSize = $true
$grpSmtp.Controls.Add($lblSrv)
$txtSrv = New-Object System.Windows.Forms.TextBox
$txtSrv.Location = '70,22'; $txtSrv.Size = '200,25'
$txtSrv.Text = ""                                   # <-- serveur SMTP : VIERGE
$grpSmtp.Controls.Add($txtSrv)

$lblPort = New-Object System.Windows.Forms.Label
$lblPort.Text = "Port :"; $lblPort.Location = '280,25'; $lblPort.AutoSize = $true
$grpSmtp.Controls.Add($lblPort)
$txtPort = New-Object System.Windows.Forms.TextBox
$txtPort.Location = '320,22'; $txtPort.Size = '50,25'
$txtPort.Text = "25"                                # 25 = relais interne classique (587 si soumission authentifiée)
$grpSmtp.Controls.Add($txtPort)

$chkSsl = New-Object System.Windows.Forms.CheckBox
$chkSsl.Text = "SSL/TLS"; $chkSsl.Location = '375,23'; $chkSsl.AutoSize = $true
$grpSmtp.Controls.Add($chkSsl)

# Expéditeur -- LAISSÉ VIERGE volontairement
$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = "Expéditeur :"; $lblFrom.Location = '10,55'; $lblFrom.AutoSize = $true
$grpSmtp.Controls.Add($lblFrom)
$txtFrom = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = '85,52'; $txtFrom.Size = '285,25'
$txtFrom.Text = ""                                  # <-- adresse expéditeur : VIERGE
$grpSmtp.Controls.Add($txtFrom)

# Authentification : Windows intégrée (défaut, adaptée à un relais AD interne)
# ou compte explicite (utilisateur + mot de passe masqué)
$chkAuth = New-Object System.Windows.Forms.CheckBox
$chkAuth.Text = "Compte explicite (sinon : authentification Windows intégrée / AD)"
$chkAuth.Location = '10,84'; $chkAuth.AutoSize = $true
$grpSmtp.Controls.Add($chkAuth)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = '10,108'; $txtUser.Size = '175,25'; $txtUser.Enabled = $false
$grpSmtp.Controls.Add($txtUser)
$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = '195,108'; $txtPass.Size = '175,25'; $txtPass.Enabled = $false
$txtPass.UseSystemPasswordChar = $true              # saisie masquée
$grpSmtp.Controls.Add($txtPass)
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "(utilisateur / mot de passe)"; $lblUser.Location = '10,136'; $lblUser.AutoSize = $true
$lblUser.ForeColor = [System.Drawing.Color]::Gray
$grpSmtp.Controls.Add($lblUser)

$chkAuth.Add_CheckedChanged({
    $txtUser.Enabled = $chkAuth.Checked
    $txtPass.Enabled = $chkAuth.Checked
})

# ------------------------------------------------------------- Prévisualisation (refactorisée)
$grpPrev = New-Object System.Windows.Forms.GroupBox
$grpPrev.Text = "4. Prévisualisation de tous les mails - sélectionne un destinataire dans la liste"
$grpPrev.Location = '15,222'; $grpPrev.Size = '865,330'
$form.Controls.Add($grpPrev)

# Bouton d'édition du modèle de mail (en haut à droite du groupe)
$btnModele = New-Object System.Windows.Forms.Button
$btnModele.Text = "Modifier le modèle..."
$btnModele.Location = '760,0'; $btnModele.Size = '100,22'
$btnModele.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpPrev.Controls.Add($btnModele)

# --- ListBox des destinataires (gauche) -----------------------------------------
$listbox = New-Object System.Windows.Forms.ListBox
$listbox.Location = '10,22'; $listbox.Size = '250,295'
$listbox.SelectionMode = 'One'
$listbox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$listbox.IntegralHeight = $false
$grpPrev.Controls.Add($listbox)

$listbox.Add_SelectedIndexChanged({
    if ($listbox.SelectedIndex -ge 0) {
        $script:index = $listbox.SelectedIndex
        Show-Apercu
    }
})

# Sauvegarder la référence de l'événement pour pouvoir le déconnecter dans Update-ListboxItemStatus
$script:listboxEventHandler = $listbox.SelectedIndexChanged.GetInvocationList()[-1]

# --- Préviz HTML du mail sélectionné (droite) ------------------------------------
$web = New-Object System.Windows.Forms.WebBrowser
$web.Location = '270,22'; $web.Size = '585,295'
$web.AllowNavigation = $false
$web.AllowWebBrowserDrop = $false
$web.IsWebBrowserContextMenuEnabled = $false
$web.WebBrowserShortcutsEnabled = $false
$web.ScriptErrorsSuppressed = $true
$grpPrev.Controls.Add($web)

# -------------------------------------------------------------- Boutons d'action (refactorisés)
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Charger les prévisualisations"
$btnStart.Location = '15,560'; $btnStart.Size = '230,36'
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 160)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = 'Flat'; $btnStart.Enabled = $false
$form.Controls.Add($btnStart)

$btnSendAll = New-Object System.Windows.Forms.Button
$btnSendAll.Text = "Envoyer tous les mails"
$btnSendAll.Location = '255,560'; $btnSendAll.Size = '200,36'
$btnSendAll.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 90)
$btnSendAll.ForeColor = [System.Drawing.Color]::White
$btnSendAll.FlatStyle = 'Flat'; $btnSendAll.Enabled = $false
$form.Controls.Add($btnSendAll)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Annuler"
$btnCancel.Location = '465,560'; $btnCancel.Size = '100,36'
$form.Controls.Add($btnCancel)
$btnCancel.Add_Click({ $form.Close() })

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = '575,567'; $progress.Size = '305,22'
$form.Controls.Add($progress)

# ----------------------------------------------------------------------- Journal
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = '15,604'; $txtLog.Size = '865,70'
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtLog)

function Write-Log {
    param([string]$Message)
    $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Message))
}

# ==============================================================================
# 3. ÉTAT DE LA CAMPAGNE
# ==============================================================================
$script:donnees         = $null    # lignes du CSV
$script:index           = -1       # index de la ligne en cours de prévisualisation
$script:stats           = @{ OK = 0; KO = 0; Passes = 0 }
$script:envoiEffectues  = @()      # liste des envois pour la preuve

# ==============================================================================
# 4. FONCTIONS DE PILOTAGE DE LA CAMPAGNE
# ==============================================================================

function Get-LigneCourante { return $script:donnees[$script:index] }

function Get-ChampsLigne {
    <# Extrait (email, nom, cheminQR) de la ligne courante selon le mapping GUI. #>
    $ligne = Get-LigneCourante
    return @{
        Email    = ([string]$ligne.($combos['Email'].SelectedItem)).Trim()
        Nom      = if ($combos['Nom'].SelectedItem) { ([string]$ligne.($combos['Nom'].SelectedItem)).Trim() } else { "" }
        CheminQR = ([string]$ligne.($combos['CheminQR'].SelectedItem)).Trim()
    }
}

function Show-Apercu {
    <#
        Affiche la prévisualisation du mail pour le destinataire à l'index $script:index.
        N'AFFECTE PAS le contenu du ListBox (pour éviter la récursion infinie via
        l'événement SelectedIndexChanged).
    #>
    if ($script:index -lt 0 -or $script:index -ge $script:donnees.Count) { return }

    $c = Get-ChampsLigne
    $num   = $script:index + 1

    # Vérification du QR
    $qrOk = (-not [string]::IsNullOrWhiteSpace($c.CheminQR)) -and (Test-Path $c.CheminQR)

    # Préviz : {QR} -> image encodée en BASE64 (data URI)
    $imgPrev = if ($qrOk) {
        try {
            $bytesQR = [IO.File]::ReadAllBytes($c.CheminQR)
            $hashQRPref = ([System.Security.Cryptography.SHA256]::Create().ComputeHash($bytesQR) | Select-Object -First 10) -join ""
            Write-Log "Aperçu QR [$num] : hash=$hashQRPref"
            
            $b64 = [Convert]::ToBase64String($bytesQR)
            "<img src=""data:image/png;base64,$b64"" style=""width:280px;height:280px;"" />"
        }
        catch {
            "<div style='color:red;border:2px dashed red;padding:30px;text-align:center;'>ERREUR LECTURE QR :<br/>$($_.Exception.Message)</div>"
        }
    } else {
        "<div style='color:red;border:2px dashed red;padding:30px;text-align:center;'>IMAGE QR INTROUVABLE :<br/>$($c.CheminQR)</div>"
    }

    # Préviz signature
    $sigOk = (-not [string]::IsNullOrWhiteSpace($CheminSignature)) -and (Test-Path $CheminSignature)
    $sigPrev = if ($sigOk) {
        try {
            $mimePrev = switch ([IO.Path]::GetExtension($CheminSignature).ToLower()) {
                '.jpg'  { 'image/jpeg' } '.jpeg' { 'image/jpeg' } '.gif' { 'image/gif' } default { 'image/png' }
            }
            $bytesSig = [IO.File]::ReadAllBytes($CheminSignature)
            $b64s = [Convert]::ToBase64String($bytesSig)
            "<img src=""data:$mimePrev;base64,$b64s"" style=""width:$($LargeurSignaturePx)px;"" />"
        }
        catch {
            "<div style='color:#b06000;'>ERREUR LECTURE SIGNATURE : $($_.Exception.Message)</div>"
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($CheminSignature)) {
        "<div style='color:#b06000;border:1px dashed #b06000;padding:8px;'>Signature configurée mais introuvable : $CheminSignature</div>"
    } else {
        ""
    }

    $html = (($ModeleCorps -replace '\{NOM\}', $c.Nom) -replace '\{QR\}', $imgPrev) -replace '\{SIGNATURE\}', $sigPrev
    $web.DocumentText = $html

    $progress.Maximum = $script:donnees.Count
    $progress.Value   = $num
}

function Update-ListboxItemStatus {
    <#
        Met à jour UNIQUEMENT le texte de l'item du ListBox à l'index donné,
        SANS déclencher l'événement SelectedIndexChanged (sinon boucle infinie).
    #>
    param([int]$Index, [string]$ItemText)
    
    # Déconnecter temporairement l'événement pour éviter la boucle infinie
    $listbox.Remove_SelectedIndexChanged($script:listboxEventHandler)
    
    if ($listbox.Items.Count -gt $Index) {
        $listbox.Items[$Index] = $ItemText
    }
    
    # Reconnecter l'événement
    $listbox.Add_SelectedIndexChanged($script:listboxEventHandler)
}

function Invoke-EnvoiCourant {
    <#
        Envoie le mail de la ligne courante.
        Retourne $true (succès) / $false (échec ou ligne invalide).
        Enregistre chaque tentative dans $script:envoiEffectues pour la preuve d'envoi.
        Toutes les erreurs sont journalisées, jamais bloquantes pour la campagne.
    #>
    $c = Get-ChampsLigne
    $num = $script:index + 1
    $timestampEnvoi = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $statut = "ERREUR"
    $message = ""

    # Garde-fous : ligne incomplète = échec contrôlé, pas d'envoi partiel
    if ([string]::IsNullOrWhiteSpace($c.Email)) {
        Write-Log "[$num] ÉCHEC : email vide."
        $message = "Email vide"
        $script:stats.KO++
        $script:envoiEffectues += [PSCustomObject]@{
            Timestamp    = $timestampEnvoi
            Destinataire = $c.Email
            Objet        = "N/A"
            Statut       = "ÉCHEC"
            Message      = $message
        }
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($c.CheminQR) -or -not (Test-Path $c.CheminQR)) {
        Write-Log "[$num] ÉCHEC : fichier QR introuvable pour $($c.Email)."
        $message = "Fichier QR introuvable : $($c.CheminQR)"
        $script:stats.KO++
        $script:envoiEffectues += [PSCustomObject]@{
            Timestamp    = $timestampEnvoi
            Destinataire = $c.Email
            Objet        = ($ModeleObjet -replace '\{NOM\}', $c.Nom)
            Statut       = "ÉCHEC"
            Message      = $message
        }
        return $false
    }

    # Identifiants : compte explicite si coché, sinon authentification Windows (AD)
    $cred = $null
    if ($chkAuth.Checked) {
        if ([string]::IsNullOrWhiteSpace($txtUser.Text)) {
            Write-Log "[$num] ÉCHEC : compte explicite coché mais utilisateur vide."
            $message = "Compte explicite coché mais utilisateur vide"
            $script:stats.KO++
            $script:envoiEffectues += [PSCustomObject]@{
                Timestamp    = $timestampEnvoi
                Destinataire = $c.Email
                Objet        = ($ModeleObjet -replace '\{NOM\}', $c.Nom)
                Statut       = "ÉCHEC"
                Message      = $message
            }
            return $false
        }
        $secure = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
        $cred   = New-Object System.Management.Automation.PSCredential($txtUser.Text, $secure)
    }

    $objet = ($ModeleObjet -replace '\{NOM\}', $c.Nom)
    try {
        Send-QRMail -Serveur     $txtSrv.Text.Trim() `
                    -Port        ([int]$txtPort.Text) `
                    -UtiliserSSL $chkSsl.Checked `
                    -Credential  $cred `
                    -Expediteur  $txtFrom.Text.Trim() `
                    -Destinataire $c.Email `
                    -Objet       $objet `
                    -CorpsHtml   ($ModeleCorps -replace '\{NOM\}', $c.Nom) `
                    -CheminQR    $c.CheminQR | Out-Null
        Write-Log "[$num] ENVOYÉ : $($c.Email)"
        $statut = "OK"
        $message = "Envoi réussi"
        $script:stats.OK++
        $succes = $true
    }
    catch {
        Write-Log "[$num] ÉCHEC SMTP pour $($c.Email) : $($_.Exception.Message)"
        $statut = "ERREUR SMTP"
        $message = $_.Exception.Message
        $script:stats.KO++
        $succes = $false
    }
    finally {
        # Enregistrer la preuve d'envoi pour CHAQUE tentative
        $script:envoiEffectues += [PSCustomObject]@{
            Timestamp    = $timestampEnvoi
            Destinataire = $c.Email
            Objet        = $objet
            Statut       = $statut
            Message      = $message
        }
    }

    return $succes
}

function Test-ParametresSmtp {
    <# Vérifie que serveur + expéditeur sont renseignés avant tout envoi. #>
    if ([string]::IsNullOrWhiteSpace($txtSrv.Text) -or [string]::IsNullOrWhiteSpace($txtFrom.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Renseignez le serveur SMTP et l'adresse expéditeur avant d'envoyer.",
            "Paramètres SMTP incomplets", 'OK', 'Warning') | Out-Null
        return $false
    }
    if (-not ($txtPort.Text -match '^\d+$')) {
        [System.Windows.Forms.MessageBox]::Show("Le port SMTP doit être numérique.", "Port invalide", 'OK', 'Warning') | Out-Null
        return $false
    }
    return $true
}

# ==============================================================================
# 5. ÉDITEUR DE MODÈLE (objet, corps HTML, signature)
# ==============================================================================
function Show-EditeurModele {
    <#
        Ouvre une fenêtre modale d'édition du modèle de mail :
        - Objet (avec variable {NOM})
        - Corps HTML (variables {NOM}, {QR} obligatoire, {SIGNATURE} optionnelle)
        - Chemin de l'image de signature (+ bouton Parcourir) et largeur en px
        Les modifications sont appliquées en mémoire ($script:...) :
        - immédiatement visibles dans la prévisualisation
        - valables pour toute la session ; le script lui-même n'est pas modifié.
    #>
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Édition du modèle de mail"
    $dlg.Size            = New-Object System.Drawing.Size(720, 560)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox     = $false
    $dlg.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

    # --- Objet -----------------------------------------------------------------
    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "Objet du mail  ({NOM} = nom du destinataire) :"
    $l1.Location = '15,12'; $l1.AutoSize = $true
    $dlg.Controls.Add($l1)

    $tObjet = New-Object System.Windows.Forms.TextBox
    $tObjet.Location = '15,33'; $tObjet.Size = '675,25'
    $tObjet.Text = $script:ModeleObjet
    $dlg.Controls.Add($tObjet)

    # --- Corps HTML --------------------------------------------------------------
    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = "Corps HTML  -  variables : {NOM}   {QR} (OBLIGATOIRE)   {SIGNATURE} (optionnelle) :"
    $l2.Location = '15,68'; $l2.AutoSize = $true
    $dlg.Controls.Add($l2)

    $tCorps = New-Object System.Windows.Forms.TextBox
    $tCorps.Location = '15,89'; $tCorps.Size = '675,290'
    $tCorps.Multiline = $true; $tCorps.ScrollBars = 'Vertical'
    $tCorps.AcceptsReturn = $true; $tCorps.WordWrap = $false
    $tCorps.Font = New-Object System.Drawing.Font("Consolas", 9)
    $tCorps.Text = $script:ModeleCorps
    $dlg.Controls.Add($tCorps)

    # --- Signature ------------------------------------------------------------------
    $l3 = New-Object System.Windows.Forms.Label
    $l3.Text = "Image de signature (vide = pas de signature) :"
    $l3.Location = '15,392'; $l3.AutoSize = $true
    $dlg.Controls.Add($l3)

    $tSig = New-Object System.Windows.Forms.TextBox
    $tSig.Location = '15,413'; $tSig.Size = '480,25'
    $tSig.Text = $script:CheminSignature
    $dlg.Controls.Add($tSig)

    $bSig = New-Object System.Windows.Forms.Button
    $bSig.Text = "Parcourir..."
    $bSig.Location = '505,411'; $bSig.Size = '90,27'
    $bSig.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Images (*.png;*.jpg;*.jpeg;*.gif)|*.png;*.jpg;*.jpeg;*.gif|Tous les fichiers (*.*)|*.*"
        $ofd.Title  = "Sélectionner l'image de signature"
        if ($ofd.ShowDialog() -eq 'OK') { $tSig.Text = $ofd.FileName }
    })
    $dlg.Controls.Add($bSig)

    $lLarg = New-Object System.Windows.Forms.Label
    $lLarg.Text = "Largeur (px) :"
    $lLarg.Location = '605,392'; $lLarg.AutoSize = $true
    $dlg.Controls.Add($lLarg)

    $numLarg = New-Object System.Windows.Forms.NumericUpDown
    $numLarg.Location = '605,413'; $numLarg.Size = '85,25'
    $numLarg.Minimum = 50; $numLarg.Maximum = 800
    $numLarg.Value = [Math]::Max(50, [Math]::Min(800, $script:LargeurSignaturePx))
    $dlg.Controls.Add($numLarg)

    # --- Boutons Enregistrer / Annuler ------------------------------------------------
    $bOk = New-Object System.Windows.Forms.Button
    $bOk.Text = "Enregistrer"
    $bOk.Location = '460,470'; $bOk.Size = '110,32'
    $bOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 90)
    $bOk.ForeColor = [System.Drawing.Color]::White
    $bOk.FlatStyle = 'Flat'
    $dlg.Controls.Add($bOk)

    $bCancel = New-Object System.Windows.Forms.Button
    $bCancel.Text = "Annuler"
    $bCancel.Location = '580,470'; $bCancel.Size = '110,32'
    $bCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($bCancel)
    $dlg.CancelButton = $bCancel

    $bOk.Add_Click({
        # Garde-fou : le marqueur {QR} est indispensable, sinon le token n'est jamais envoyé
        if ($tCorps.Text -notmatch '\{QR\}') {
            [System.Windows.Forms.MessageBox]::Show(
                "Le corps du mail doit contenir le marqueur {QR} : c'est lui qui reçoit le code QR du destinataire.",
                "Marqueur {QR} manquant", 'OK', 'Warning') | Out-Null
            return
        }
        # Avertissement non bloquant si la signature est renseignée mais introuvable
        if (-not [string]::IsNullOrWhiteSpace($tSig.Text) -and -not (Test-Path $tSig.Text.Trim())) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                "L'image de signature est introuvable :`n$($tSig.Text)`n`nEnregistrer quand même ? (les mails partiront sans signature)",
                "Signature introuvable", 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }

        # Application en mémoire (portée script -> utilisée par préviz ET envoi)
        $script:ModeleObjet        = $tObjet.Text
        $script:ModeleCorps        = $tCorps.Text
        $script:CheminSignature    = $tSig.Text.Trim()
        $script:LargeurSignaturePx = [int]$numLarg.Value
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    if ($dlg.ShowDialog($form) -eq 'OK') {
        Write-Log "Modèle de mail mis à jour (objet/corps/signature)."
        # Rafraîchit la prévisualisation en cours pour refléter le nouveau modèle
        if ($script:index -ge 0 -and $script:index -lt $script:donnees.Count) {
            Show-Apercu
        }
    }
    $dlg.Dispose()
}

$btnModele.Add_Click({ Show-EditeurModele })

# ==============================================================================
# 6. ÉVÉNEMENTS
# ==============================================================================

# --- Chargement du CSV enrichi -------------------------------------------------
$btnCsv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $ofd.Title  = "Sélectionner le CSV enrichi (avec colonne CheminQR)"
    if ($ofd.ShowDialog() -ne 'OK') { return }

    $txtCsv.Text = $ofd.FileName

    # Import en UTF-8, repli sur l'encodage par défaut (ANSI) si échec
    try   { $script:donnees = @(Import-Csv -Path $ofd.FileName -Delimiter ';' -Encoding UTF8) }
    catch { $script:donnees = @(Import-Csv -Path $ofd.FileName -Delimiter ';') }

    if (-not $script:donnees -or $script:donnees.Count -eq 0) {
        Write-Log "ERREUR : CSV vide ou illisible."
        return
    }

    $colonnes = @($script:donnees[0].PSObject.Properties.Name)
    foreach ($k in $combos.Keys) {
        $combos[$k].Items.Clear()
        $combos[$k].Items.AddRange($colonnes)
    }

    # Auto-détection des colonnes (mêmes conventions que le générateur)
    $autoMail = Find-Column $colonnes @('mail', 'courriel')
    $autoNom  = Find-Column $colonnes @('^nom$', 'nom', 'name', 'utilisateur', 'user')
    $autoQR   = Find-Column $colonnes @('cheminqr', 'qr', 'chemin', 'path')
    if ($autoMail) { $combos['Email'].SelectedItem    = $autoMail }
    if ($autoNom)  { $combos['Nom'].SelectedItem      = $autoNom }
    if ($autoQR)   { $combos['CheminQR'].SelectedItem = $autoQR }

    $lblCount.Text = "$($script:donnees.Count) ligne(s)"
    $lblCount.ForeColor = [System.Drawing.Color]::FromArgb(0,120,90)
    Write-Log "CSV chargé : $($script:donnees.Count) destinataire(s) potentiel(s)."
    $btnStart.Enabled = $true
})

# --- Charger les prévisualisations -----------------------------------------------
$btnStart.Add_Click({
    if (-not $combos['Email'].SelectedItem -or -not $combos['CheminQR'].SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show(
            "Les colonnes Email et Chemin QR doivent être mappées.",
            "Mapping incomplet", 'OK', 'Warning') | Out-Null
        return
    }

    $listbox.Items.Clear()
    $script:index              = -1
    $script:stats              = @{ OK = 0; KO = 0; Passes = 0 }
    $script:envoiEffectues     = @()    # réinitialise la liste des envois
    
    # Peuple la ListBox avec tous les destinataires (sans déclencher SelectedIndexChanged)
    for ($i = 0; $i -lt $script:donnees.Count; $i++) {
        $script:index = $i
        $c = Get-ChampsLigne
        $qrOk = (-not [string]::IsNullOrWhiteSpace($c.CheminQR)) -and (Test-Path $c.CheminQR)
        $etat = if ($qrOk) { "✓" } else { "✗" }
        $itemText = "[$($i+1)] $($c.Email) - $etat"
        $listbox.Items.Add($itemText)
    }

    Write-Log "--- Prévisualisations chargées : $($script:donnees.Count) destinataire(s) ---"
    $btnStart.Enabled = $false
    $btnSendAll.Enabled = $true
    
    # Affiche la première préviz
    $script:index = 0
    $listbox.SelectedIndex = 0
})

# --- Envoyer tous les mails -------------------------------------------------------
$btnSendAll.Add_Click({
    if (-not (Test-ParametresSmtp)) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Envoyer $($script:donnees.Count) mail(s) ?",
        "Confirmation envoi de masse", 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }

    $btnSendAll.Enabled = $false
    $btnStart.Enabled = $false
    $listbox.Enabled = $false
    $script:stats = @{ OK = 0; KO = 0; Passes = 0 }
    $script:envoiEffectues = @()    # réinitialise pour cette campagne d'envoi

    for ($i = 0; $i -lt $script:donnees.Count; $i++) {
        $script:index = $i
        $c = Get-ChampsLigne
        $num = $i + 1

        $progress.Maximum = $script:donnees.Count
        $progress.Value   = $num
        [System.Windows.Forms.Application]::DoEvents()

        Invoke-EnvoiCourant | Out-Null
        
        # Met à jour l'item de la liste avec le statut final (sans déclencher SelectedIndexChanged)
        $qrOk = (-not [string]::IsNullOrWhiteSpace($c.CheminQR)) -and (Test-Path $c.CheminQR)
        $etat = if ($qrOk) { "✓" } else { "✗" }
        $statusText = if ($c.CheminQR -and $qrOk) { "ENVOYÉ" } else { "ERREUR" }
        $itemText = "[$num] $($c.Email) - $etat - $statusText"
        Update-ListboxItemStatus -Index $i -ItemText $itemText
        
        Start-Sleep -Milliseconds $PauseEntreEnvoisMs
    }

    $btnSendAll.Enabled = $true
    $btnStart.Enabled = $true
    $listbox.Enabled = $true
    
    # --- Génération du fichier de preuve d'envoi ---
    $preuveFilename = "preuve_envoi_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $preuveFilepath = Join-Path $txtOut.Text $preuveFilename
    $script:envoiEffectues | Export-Csv -Path $preuveFilepath -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Write-Log "Preuve d'envoi générée : $preuveFilepath"
    
    $bilan = "Bilan final : $($script:stats.OK) envoyé(s), $($script:stats.KO) échec(s).`n`nPreuve d'envoi : $preuveFilename"
    Write-Log "--- $bilan ---"
    [System.Windows.Forms.MessageBox]::Show($bilan, "Campagne terminée", 'OK', 'Information') | Out-Null
})

# ==============================================================================
# 7. LANCEMENT
# ==============================================================================
Write-Log "Prêt. Chargez le CSV enrichi, renseignez le SMTP, puis démarrez la campagne."
[void]$form.ShowDialog()
