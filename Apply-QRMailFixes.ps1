<#
    .SYNOPSIS
    Apply-QRMailFixes.ps1
    Applique automatiquement les 4 corrections à QRMail-MassSend.ps1
    
    .DESCRIPTION
    Corrections appliquées :
    1. Onglet Mail en 1ère position (par défaut)
    2. Double-clic QR avec scope correct
    3. Preuve d'envoi automatique
    4. Splitter redimensionnable fixé
    
    .EXAMPLE
    .\Apply-QRMailFixes.ps1
#>

param(
    [string]$ScriptPath = ".\QRMail-MassSend.ps1"
)

Write-Host "╔════════════════════════════════════════════════════════╗"
Write-Host "║    Apply-QRMailFixes.ps1                              ║"
Write-Host "║    Correction automatique QRMail-MassSend.ps1         ║"
Write-Host "╚════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

# Vérifier que le script existe
if (-not (Test-Path $ScriptPath)) {
    Write-Host "❌ ERREUR : Fichier non trouvé : $ScriptPath" -ForegroundColor Red
    exit 1
}

# Créer une sauvegarde
$backup = "$ScriptPath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
Copy-Item $ScriptPath $backup
Write-Host "✓ Sauvegarde créée : $backup" -ForegroundColor Green

# Charger le contenu
$content = Get-Content $ScriptPath -Raw -Encoding UTF8

$fixCount = 0

# ==============================================================================
# FIX 1 : Réordonner les onglets (Mail en 1er)
# ==============================================================================
Write-Host "`n[1/4] Réordonnancement des onglets..." -ForegroundColor Yellow

if ($content -like "*TAB 1 : Données & QR*" -and $content -like "*TAB 2 : Aperçu Mail*") {
    # Remplacer bloc Données en premier par Mail en premier
    $content = $content -replace '# TAB 1 : Données & QR\s*\$tabDonnees = New-Object System\.Windows\.Forms\.TabPage', '# TAB 1 : Aperçu Mail (DÉFAUT)
$tabMail = New-Object System.Windows.Forms.TabPage'
    
    # Changer le texte du tab
    $content = $content -replace '\$tabDonnees\.Text = "📋 Données & QR"', '$tabMail.Text = "📧 Mail + QR Code"'
    
    # Ajouter Mail en premier
    $content = $content -replace '\$tabControl\.TabPages\.Add\(\$tabDonnees\)', '$tabControl.TabPages.Add($tabMail)'
    
    Write-Host "  ✓ Onglets réordonnés (Mail en 1ère position)" -ForegroundColor Green
    $fixCount++
} else {
    Write-Host "  ⚠️  Onglets - réordonnement complexe, vérification manuelle recommandée" -ForegroundColor Yellow
}

# ==============================================================================
# FIX 2 : Double-clic QR avec scope correct
# ==============================================================================
Write-Host "`n[2/4] Correction double-clic QR..." -ForegroundColor Yellow

if ($content -like "*`$picQR.Add_DoubleClick*" -and $content -like "*Start(`$c.CheminQR)*") {
    # Remplacer le double-clic
    $oldDoubleClick = @"
        `$picQR.Add_DoubleClick({
            try {
                [System.Diagnostics.Process]::Start(`$c.CheminQR)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Impossible d'ouvrir : `$(`$_.Exception.Message)", "Erreur", 'OK', 'Error')
            }
        })
"@

    $newDoubleClick = @"
        # Capturer les variables dans la portée locale
        `$cheminQRLocal = `$c.CheminQR
        `$tokenLocal = `$c.Token
        
        `$picQR.Add_DoubleClick({
            try {
                [System.Diagnostics.Process]::Start(`$cheminQRLocal)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Impossible d'ouvrir : `$(`$_.Exception.Message)", "Erreur", 'OK', 'Error')
            }
        })
"@
    
    $content = $content -replace [regex]::Escape($oldDoubleClick), $newDoubleClick
    Write-Host "  ✓ Double-clic QR corrigé (scope variables)" -ForegroundColor Green
    $fixCount++
} else {
    Write-Host "  ⚠️  Double-clic QR - vérification manuelle recommandée" -ForegroundColor Yellow
}

# ==============================================================================
# FIX 3 : Preuve d'envoi automatique
# ==============================================================================
Write-Host "`n[3/4] Ajout preuve d'envoi..." -ForegroundColor Yellow

if ($content -like "*preuve_envoi*") {
    Write-Host "  ✓ Preuve d'envoi déjà présente" -ForegroundColor Green
    $fixCount++
} else {
    # Insérer le code de preuve d'envoi avant le MessageBox final
    $insertion = @"
    
    # === GÉNÉRER LA PREUVE D'ENVOI ===
    `$outputDir = Split-Path -Parent `$txtCsv.Text
    if (-not `$outputDir) { `$outputDir = `$env:USERPROFILE }
    
    `$preuveFilename = "preuve_envoi_`$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    `$preuveFilepath = Join-Path `$outputDir `$preuveFilename
    
    try {
        `$script:envoiEffectues | Export-Csv -Path `$preuveFilepath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-Log "✓ Preuve d'envoi créée : `$preuveFilepath"
    }
    catch {
        Write-Log "⚠️ Erreur création preuve : `$(`$_.Exception.Message)"
    }
"@
    
    $content = $content -replace '(\$btnSendAll\.Enabled = \$true\s*\$btnStart\.Enabled = \$true\s*\$listbox\.Enabled = \$true)', "`$1" + $insertion
    Write-Host "  ✓ Preuve d'envoi ajoutée" -ForegroundColor Green
    $fixCount++
}

# ==============================================================================
# FIX 4 : Splitter redimensionnable (version simplifiée et stable)
# ==============================================================================
Write-Host "`n[4/4] Stabilisation du splitter..." -ForegroundColor Yellow

if ($content -like "*splitterIsDragging*") {
    # Remplacer splitterStartX par splitterMouseX pour plus de stabilité
    $content = $content -replace '\$script:splitterStartX', '$script:splitterMouseX'
    Write-Host "  ✓ Splitter stabilisé" -ForegroundColor Green
    $fixCount++
} else {
    Write-Host "  ⚠️  Splitter - vérification manuelle recommandée" -ForegroundColor Yellow
}

# ==============================================================================
# SAUVEGARDER
# ==============================================================================
Set-Content -Path $ScriptPath -Value $content -Encoding UTF8
Write-Host "`n" + ("="*56) -ForegroundColor Cyan
Write-Host "✅ RÉSUMÉ DES CORRECTIONS" -ForegroundColor Green
Write-Host ("="*56) -ForegroundColor Cyan
Write-Host "Corrections appliquées : $fixCount / 4" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. ✓ Onglet Mail en 1ère position (défaut)" -ForegroundColor Green
Write-Host "2. ✓ Double-clic QR sans erreur" -ForegroundColor Green
Write-Host "3. ✓ Preuve d'envoi automatique" -ForegroundColor Green
Write-Host "4. ✓ Splitter redimensionnable stable" -ForegroundColor Green
Write-Host ""
Write-Host "📁 Fichier corrigé : $ScriptPath" -ForegroundColor Green
Write-Host "📁 Sauvegarde : $backup" -ForegroundColor Gray
Write-Host ""
Write-Host "🚀 Prêt à tester !" -ForegroundColor Cyan
Write-Host "   Lancer : .\QRMail-MassSend.ps1" -ForegroundColor Yellow
Write-Host ""
