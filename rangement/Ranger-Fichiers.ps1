<#
  Ranger-Fichiers.ps1 - Range Téléchargements + Documents selon l'arborescence :

    01_Alpha-Omega-AI\Clients\<client>   \Offres-et-methodes   \Site-et-marque   \Gestion
    02_Softel-Stellantis
    03_DermaOxy-France
    04_Personnel\Administratif   \Parcours
    99_Archives\<mission terminée>
    00_A-trier                      (classement incertain + LISEZ-MOI avec une ligne par fichier)
    00_A-trier\_A-valider-suppression (doublons, installeurs, archives déjà extraites : RIEN n'est supprimé)

  Utilisation (PowerShell, OneDrive en pause, Word/Excel/PowerPoint fermés) :
    1) Simulation (ne touche à rien, produit le plan CSV) :
         powershell -ExecutionPolicy Bypass -File .\Ranger-Fichiers.ps1
    2) Exécution réelle :
         powershell -ExecutionPolicy Bypass -File .\Ranger-Fichiers.ps1 -Execute
    Annulation : lancer le fichier ANNULER_<date>.ps1 créé dans 00_A-trier\_journal_<date>\
#>
param([switch]$Execute)

$ErrorActionPreference = 'Stop'
# Sans OneDrive : tout est rangé dans les dossiers LOCAUX C:\Users\<vous>\Documents et \Downloads.
# Le contenu éventuellement resté dans C:\Users\<vous>\OneDrive*\Documents (ou Téléchargements) est rapatrié.
$Downloads = Join-Path $env:USERPROFILE 'Downloads'
$Documents = Join-Path $env:USERPROFILE 'Documents'
$Root      = $Documents
$Sources   = @($Downloads, $Documents) + @(Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -Filter 'OneDrive*' -ErrorAction SilentlyContinue |
    ForEach-Object { foreach ($s in 'Documents','Downloads','Téléchargements') { Join-Path $_.FullName $s } } |
    Where-Object { Test-Path -LiteralPath $_ -PathType Container })
foreach ($d in $Downloads, $Documents) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d | Out-Null } }

if (Get-Process -Name OneDrive -ErrorAction SilentlyContinue) {
    Write-Host 'OneDrive est encore ouvert. Dissociez ce PC de OneDrive et quittez OneDrive, puis relancez le script.' -ForegroundColor Red
    exit 1
}
# Fichiers « uniquement en ligne » (nuage) : ils ne sont pas sur le disque, les déplacer les casserait.
$EnLigne = @($Sources | ForEach-Object { Get-ChildItem -LiteralPath $_ -Recurse -File -Force -ErrorAction SilentlyContinue } |
    Where-Object { ($_.Attributes -band 0x400000) -or ($_.Attributes -band 0x1000) })
if ($EnLigne.Count -gt 0) {
    $liste = Join-Path $env:USERPROFILE "Desktop\fichiers_uniquement_en_ligne_$(Get-Date -Format yyyy-MM-dd_HHmm).txt"
    $EnLigne.FullName | Set-Content -Encoding UTF8 $liste
    Write-Host "$($EnLigne.Count) fichier(s) ne sont que dans le nuage OneDrive (liste : $liste)." -ForegroundColor Red
    Write-Host 'Rouvrez OneDrive, clic droit sur le dossier OneDrive > « Toujours conserver sur cet appareil », attendez la fin, puis dissociez et relancez.' -ForegroundColor Red
    exit 1
}
$Stamp     = Get-Date -Format 'yyyy-MM-dd_HHmm'
$ATrier    = '00_A-trier'
$Corbeille = '00_A-trier\_A-valider-suppression'

# ---------------------------------------------------------------------------
# CLIENTS (déduits des noms de fichiers du PC). Statut 'Actif' -> 01_Alpha-Omega-AI\Clients\<Nom>
#                                    Statut 'Termine' -> 99_Archives\<Nom>
# ---------------------------------------------------------------------------
$Clients = @(
    # En cours (fichiers 2025-2026)
    @{ Nom = 'Hotel-Concierge-WhatsApp'; Motif = 'hotel|concierge|check.?in|whatsapp.?(agent|concierge)'; Statut = 'Actif' }
    @{ Nom = 'Premodelisation-juridique-IA'; Motif = 'pre.?mod[eé]lisation';  Statut = 'Actif' }
    @{ Nom = 'MailSense';   Motif = 'mail.?sense';          Statut = 'Actif'   }
    @{ Nom = 'OWI';         Motif = '(^|[^a-z])owi([^a-z]|$)'; Statut = 'Actif' }
    @{ Nom = 'Knwler';      Motif = 'knwler';               Statut = 'Actif'   }
    @{ Nom = 'SPERO';       Motif = 'spero';                Statut = 'Actif'   }
    # Terminés
    @{ Nom = 'MGEN';        Motif = 'mgen';                 Statut = 'Termine' }
    @{ Nom = 'FoxVisit';    Motif = 'fox.?visit';           Statut = 'Termine' }
    @{ Nom = 'TFO';         Motif = '(^|[^a-z])tfo';        Statut = 'Termine' }
    @{ Nom = 'AFPA';        Motif = 'afpa';                 Statut = 'Termine' }
    @{ Nom = 'Mobilis';     Motif = 'mobilis';              Statut = 'Termine' }
    @{ Nom = 'Apivia';      Motif = 'apivia';               Statut = 'Termine' }
    @{ Nom = 'ThoughtSpot'; Motif = 'thought.?spot';        Statut = 'Termine' }
    @{ Nom = 'MoniA';       Motif = 'monia';                Statut = 'Termine' }
    @{ Nom = 'Governata';   Motif = 'governata|naima';      Statut = 'Termine' }
    @{ Nom = 'Finbursa';    Motif = 'finbursa';             Statut = 'Termine' }
    @{ Nom = 'CAFS';        Motif = 'cafs';                 Statut = 'Termine' }
    @{ Nom = 'Cybtech';     Motif = 'cybtech';              Statut = 'Termine' }
    @{ Nom = 'InGrav';      Motif = 'ingrav';               Statut = 'Termine' }
    @{ Nom = 'Visioneers';  Motif = 'visioneers';           Statut = 'Termine' }
)

# Règles thématiques, testées dans l'ordre sur le nom du fichier ou du dossier.
$AO = '(alpha.?omega|aomega|(^|[^a-z])ao[\s_-]?ai)'
$AvantClients = @(
    @{ M = 'softel|sofltel|stellantis|john.?paul|nexidia|speech.?analytics|isms|computer use agreement'; D = '02_Softel-Stellantis'; C = 'Softel-Stellantis' }
    @{ M = 'derma.?oxy';                                                               D = '03_DermaOxy-France';                     C = 'DermaOxy' }
    @{ M = "$AO.*(rib|iban|kbis|statut|factur|devis)|(rib|iban|kbis|statut).*$AO";     D = '01_Alpha-Omega-AI\Gestion';              C = 'Alpha-Omega-AI' }
    @{ M = '(^|[^a-z])(cockpit|radar|oii)([^a-z]|$)';                                  D = '01_Alpha-Omega-AI\Offres-et-methodes';   C = 'Alpha-Omega-AI' }
    @{ M = "$AO.*(logo|charte|site|brand|marque|banni|banner|favicon|chanel|\.html)|(logo|banni[eè]re|banner).*$AO|banni[eè]re linkedin|^(logo|charte graphique|favicon)|^alpha-omega-ai.*\.html$"; D = '01_Alpha-Omega-AI\Site-et-marque'; C = 'Alpha-Omega-AI' }
    @{ M = $AO;                                                                        D = '01_Alpha-Omega-AI\Offres-et-methodes';   C = 'Alpha-Omega-AI' }
    @{ M = 'edf|engie|loyer|quittance|charges|bouygues|(^|[^a-z])(free|sfr|orange)([^a-z]|$)|comptededepots|(^|[^a-z])(rib|iban)([^a-z_]|_0)|nickel|(^|[^a-z])sepa([^a-z]|$)'; D = '04_Personnel\Administratif'; C = 'Personnel' }
    @{ M = 'factur|invoice|devis|avoir|note de frais|urssaf|kbis|statuts|compta|bilan|(^|[^a-z])tva([^a-z]|$)|liasse'; D = '01_Alpha-Omega-AI\Gestion'; C = 'Alpha-Omega-AI' }
)
$ApresClients = @(
    @{ M = 'imp[oô]t|imposition|passeport|paseport|passport|(^|[^a-z])cni([^a-z]|$)|carte.?(vitale|identit)|(^|[^a-z])caf([^a-z]|$)|ameli|cpam|mutuelle|france.?travail|p[oô]le.?emploi|arr[eê]t|m[eé]decin|ordonnance|(^|[^a-z])bail([^a-z]|$)|banque|relev[eé] de compte|attestation|convocation|bulletin de (salaire|paie)|fiche de paie|payslip|indemnit|cerfa|souffrance au travail|performences? review|offer.?letter|maman|papa|parents|samira|mohamed|no[eé]mie|remboursement|syrine|ghalamallah_ilham|\d ghalamallah|^ao\d+_'; D = '04_Personnel\Administratif'; C = 'Personnel' }
    @{ M = '(^|[^a-z])(cv|resume|curriculum)([^a-z]|$)|résumé|dipl[oô]m|diploma|(^|[^a-z])phd([^a-z]|$)|th[eè]se|publication|certificat|relev[eé] de notes|ilh[eè]me.?.?ghalamallah|lettre de motivation|assignment|(^|[^a-z])(bio|profil)([^a-z]|$)'; D = '04_Personnel\Parcours'; C = 'Personnel' }
    @{ M = 'contrat|contract|(^|[^a-z])nda([^a-z]|$)|(^|[^a-z])cgv([^a-z]|$)|avenant|bon de commande|purchase order|juridique|mandat|(^|[^a-z])sow([^a-z]|$)|statement of work'; D = '01_Alpha-Omega-AI\Gestion'; C = 'Alpha-Omega-AI' }
    @{ M = 'appel.?d.?offre|tender|(^|[^a-z])rfp([^a-z]|$)|offer|offre|proposal|proposition|pitch|methodo|framework'; D = '01_Alpha-Omega-AI\Offres-et-methodes'; C = 'Alpha-Omega-AI' }
    @{ M = '(^|[^a-z])(ai|ia|genai|llm|gpt|agents?|agentic|rag)([^a-z]|$)|agentic|artificial intelligence|intelligence artificielle|state of ai|machine learning|knowledge|analytics|mckinsey|predictions|playbook|roadmap|blueprint|radar'; D = '01_Alpha-Omega-AI\Offres-et-methodes\Veille-IA'; C = 'Alpha-Omega-AI' }
)

$Cibles = @('01_Alpha-Omega-AI\Clients','01_Alpha-Omega-AI\Offres-et-methodes','01_Alpha-Omega-AI\Site-et-marque',
            '01_Alpha-Omega-AI\Gestion','02_Softel-Stellantis','03_DermaOxy-France','04_Personnel\Administratif',
            '04_Personnel\Parcours','99_Archives',$ATrier,$Corbeille)
$NomsCibles = '^(00_A-trier|01_Alpha-Omega-AI|02_Softel-Stellantis|03_DermaOxy-France|04_Personnel|99_Archives)$'
# Dossiers d'applications laissés en place dans Documents
$Laisser = '^(desktop\.ini|Custom Office Templates|Modèles Office personnalisés|Fichiers Outlook|Outlook Files|WindowsPowerShell|PowerShell|Zoom|My Music|My Pictures|My Videos|Ma musique|Mes images|Mes vidéos|Sound recordings|Enregistrements audio|Visual Studio.*|IISExpress|My Web Sites|Default\.rdp)$'
$Installeurs = '\.(exe|msi|msix|msixbundle|appx|appxbundle|dmg|pkg)$'
$Temporaires = '^(~\$.*|thumbs\.db|\.ds_store)$|\.(tmp|crdownload|part|partial)$'
$Archives    = '\.(zip|rar|7z)$'
$Illisible   = '^((export|file)_[0-9a-f\-]{8,}.*|doc-\d{8}-wa\d+_?.*|img[-_]\d{8}.*|screenshot_\d.*|sodapdf.*|winmail.*|report-\d+.*|(scan(ned)?|num[eé]risation|img|image|dsc|photo|document|doc|sans titre|untitled|nouveau document( texte)?|new document|fichier|file|download|t[eé]l[eé]chargement|capture( d.?[eé]cran)?|screenshot|copie de)[\s_\-\.\(\)\d]*|[a-z]{0,3}[\s_\-]?\d[\d\s_\-\.\(\)]*)$'
$ExtImages   = '\.(jpe?g|png|heic|gif|bmp|webp|tiff?)$'

# ---------------------------------------------------------------------------
$Plan     = New-Object System.Collections.Generic.List[object]
$Reserved = @{}

function Get-FreePath([string]$dir, [string]$name, [bool]$isDir) {
    if ($isDir) { $base = $name; $ext = '' }
    else { $base = [IO.Path]::GetFileNameWithoutExtension($name); $ext = [IO.Path]::GetExtension($name) }
    $p = Join-Path $dir $name; $i = 2
    while ((Test-Path -LiteralPath $p) -or $Reserved.ContainsKey($p.ToLower())) {
        $p = Join-Path $dir ('{0} ({1}){2}' -f $base, $i, $ext); $i++
    }
    $Reserved[$p.ToLower()] = $true
    return $p
}

function Get-Classement([string]$nom) {
    $n = $nom.ToLower()
    foreach ($r in $AvantClients) { if ($n -match $r.M) { return @{ D = $r.D; C = $r.C; R = "mot-clé « $($Matches[0].Trim()) »" } } }
    foreach ($c in $Clients) {
        if ($n -match $c.Motif) {
            if ($c.Statut -eq 'Actif') { $d = "01_Alpha-Omega-AI\Clients\$($c.Nom)" } else { $d = "99_Archives\$($c.Nom)" }
            return @{ D = $d; C = $c.Nom; R = "client $($c.Nom) ($($c.Statut))" }
        }
    }
    foreach ($r in $ApresClients) { if ($n -match $r.M) { return @{ D = $r.D; C = $r.C; R = "mot-clé « $($Matches[0].Trim()) »" } } }
    return $null
}

function New-NomLisible($item, [string]$client) {
    $base = [IO.Path]::GetFileNameWithoutExtension($item.Name)
    $date = $item.LastWriteTime
    if ($base -match '(20\d{2})[-_]?([01]\d)[-_]?([0-3]\d)') {
        try { $date = [datetime]::new([int]$Matches[1], [int]$Matches[2], [int]$Matches[3]) } catch {}
    }
    $b = $base.ToLower()
    if     ($b -match 'capture|screenshot')                  { $sujet = 'Capture-ecran' }
    elseif ($b -match 'scan|num' -or $b -match '^[a-z]{0,3}[\s_\-]?\d') { $sujet = 'Scan' }
    elseif ($item.Name -match $ExtImages -or $b -match 'img|dsc|photo|image') { $sujet = 'Photo' }
    else                                                     { $sujet = 'Document' }
    return '{0}_{1}_{2}{3}' -f $date.ToString('yyyy-MM-dd'), $client, $sujet, $item.Extension.ToLower()
}

function Add-Action($item, [string]$destRel, [string]$type, [string]$raison, [string]$nouveauNom) {
    $isDir = $item.PSIsContainer
    if (-not $nouveauNom) { $nouveauNom = $item.Name }
    $dest = Get-FreePath (Join-Path $Root $destRel) $nouveauNom $isDir
    $Plan.Add([pscustomobject]@{
        Type = $type; Source = $item.FullName; Destination = $dest
        Renomme = $(if ($nouveauNom -ne $item.Name) { 'oui' } else { '' })
        Raison = $raison; Statut = 'prévu'; Erreur = ''
    })
}

# ---------------------------------------------------------------------------
Write-Host "Téléchargements : $Downloads"
Write-Host "Documents       : $Documents"
Write-Host "Sources         : $($Sources -join ' | ')"
Write-Host ($(if ($Execute) { 'MODE EXÉCUTION' } else { 'MODE SIMULATION (aucun fichier touché)' })) -ForegroundColor Yellow

$Elements = @()
$Elements += Get-ChildItem -LiteralPath $Downloads -Force
$Elements += $Sources | Where-Object { $_ -ne $Downloads } | ForEach-Object { Get-ChildItem -LiteralPath $_ -Force } |
    Where-Object { $_.Name -notmatch $NomsCibles -and $_.Name -notmatch $Laisser }
# Dossiers « fourre-tout » de la société : on range leur contenu un par un au lieu de les déplacer en bloc
$Conteneurs = '^(aomega|aomega[\s_-]?ai|alpha[\s_-]?omega([\s_-]?ai)?|ao[\s_-]?ai)$'
$Elements = @($Elements | ForEach-Object {
    if ($_.PSIsContainer -and $_.Name -match $Conteneurs) { Get-ChildItem -LiteralPath $_.FullName -Force } else { $_ }
})
$Laisses  = $Sources | Where-Object { $_ -ne $Downloads } | ForEach-Object { Get-ChildItem -LiteralPath $_ -Force } | Where-Object { $_.Name -match $Laisser }

# --- Doublons : fichiers libres (racine Téléchargements/Documents) identiques à un autre fichier ---
Write-Host 'Recherche des doublons (empreinte SHA-256)...'
$Libres = @($Elements | Where-Object { -not $_.PSIsContainer })
$LibresSet = @{}; foreach ($f in $Libres) { $LibresSet[$f.FullName] = $true }
$Reference = @($Libres) + @($Sources | Where-Object { $_ -ne $Downloads } | ForEach-Object { Get-ChildItem -LiteralPath $_ -Recurse -File -Force -ErrorAction SilentlyContinue } |
    Where-Object { $_.FullName -notmatch '\\_ORGANIZED_' -and $Sources -notcontains $_.DirectoryName })
$TaillesLibres = @{}; foreach ($f in $Libres) { if ($f.Length -gt 0) { $TaillesLibres[$f.Length] = $true } }
$Doublons = @{}
$Reference | Where-Object { $TaillesLibres.ContainsKey($_.Length) } | Group-Object Length | Where-Object Count -gt 1 | ForEach-Object {
    $_.Group | ForEach-Object {
        try { [pscustomobject]@{ F = $_; H = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } } catch {}
    } | Group-Object H | Where-Object Count -gt 1 | ForEach-Object {
        # On garde en priorité un fichier déjà rangé dans un sous-dossier, sinon le nom le plus court puis le plus ancien
        $tri = $_.Group | Sort-Object @{ E = { if ($LibresSet.ContainsKey($_.F.FullName)) { 1 } else { 0 } } }, @{ E = { $_.F.Name.Length } }, @{ E = { $_.F.LastWriteTime } }
        $garde = $tri[0].F
        $tri | Select-Object -Skip 1 | Where-Object { $LibresSet.ContainsKey($_.F.FullName) } | ForEach-Object { $Doublons[$_.F.FullName] = $garde.FullName }
    }
}

# --- Plan de classement ---
foreach ($it in $Elements) {
    $nom = $it.Name

    if ($it.PSIsContainer -and $nom -match '^_ORGANIZED_') {
        Add-Action $it $Corbeille 'à valider' 'Squelette + 281 doublons de la tentative d''avril 2026 (originaux conservés ailleurs)'; continue
    }
    if ($it.PSIsContainer -and $nom -match '^bureau[\s_-]*\d{6,8}$') {
        Add-Action $it '99_Archives' 'déplacement' 'Ancienne sauvegarde de bureau, archivée telle quelle'; continue
    }
    if (-not $it.PSIsContainer) {
        if ($Doublons.ContainsKey($it.FullName)) { Add-Action $it $Corbeille 'à valider' "Doublon exact de : $($Doublons[$it.FullName])"; continue }
        if ($nom -match $Temporaires)  { Add-Action $it $Corbeille 'à valider' 'Fichier temporaire / verrou Office / téléchargement incomplet'; continue }
        if ($nom -match $Installeurs)  { Add-Action $it $Corbeille 'à valider' 'Installeur de logiciel'; continue }
        if ($nom -match $Archives) {
            $b = ([IO.Path]::GetFileNameWithoutExtension($nom)) -replace '\s*\(\d+\)$', ''
            $extrait = $Sources | ForEach-Object { Join-Path $_ $b } | Where-Object { Test-Path -LiteralPath $_ -PathType Container } | Select-Object -First 1
            if ($extrait) { Add-Action $it $Corbeille 'à valider' "Archive déjà extraite dans : $extrait"; continue }
        }
    }

    $cl = Get-Classement $nom
    $base = $(if ($it.PSIsContainer) { $nom } else { [IO.Path]::GetFileNameWithoutExtension($nom) })
    $illisible = (-not $it.PSIsContainer) -and ($base.ToLower() -match $Illisible)

    if ($cl) {
        $nn = $(if ($illisible) { New-NomLisible $it $cl.C } else { $null })
        Add-Action $it $cl.D 'déplacement' $cl.R $nn
    } else {
        if ($illisible) { $r = 'Nom illisible, aucun indice de client/sujet : ouvrir pour identifier'; $nn = New-NomLisible $it 'Inconnu' }
        elseif ($nom -match $ExtImages) { $r = 'Image sans contexte (perso ou pro ?)'; $nn = $null }
        elseif ($it.PSIsContainer) { $r = 'Dossier sans mot-clé client/thème reconnu'; $nn = $null }
        else { $r = 'Aucun mot-clé client/thème reconnu dans le nom'; $nn = $null }
        Add-Action $it $ATrier 'à trier' $r $nn
    }
}

# --- Rapport / exécution ---
$Journal = Join-Path $Root "$ATrier\_journal_$Stamp"
if (-not $Execute) { $Journal = Join-Path $env:TEMP "rangement_simulation_$Stamp" }
New-Item -ItemType Directory -Force -Path $Journal | Out-Null

if ($Execute) {
    foreach ($c in $Cibles) { New-Item -ItemType Directory -Force -Path (Join-Path $Root $c) | Out-Null }
    $i = 0
    foreach ($a in $Plan) {
        $i++; Write-Progress -Activity 'Rangement' -Status $a.Source -PercentComplete (100 * $i / $Plan.Count)
        try {
            New-Item -ItemType Directory -Force -Path (Split-Path $a.Destination) | Out-Null
            if (-not (Test-Path -LiteralPath $a.Source -PathType Container)) {
                $f = Get-Item -LiteralPath $a.Source -Force
                if ($f.IsReadOnly) { $f.IsReadOnly = $false }
            }
            Move-Item -LiteralPath $a.Source -Destination $a.Destination
            $a.Statut = 'fait'
        } catch { $a.Statut = 'ERREUR'; $a.Erreur = $_.Exception.Message }
    }
    Write-Progress -Activity 'Rangement' -Completed

    # Script d'annulation (ordre inverse)
    $annul = @('# Remet chaque élément à son emplacement d''origine', '$ErrorActionPreference = ''Continue''')
    foreach ($a in ($Plan | Where-Object Statut -eq 'fait' | Sort-Object { $Plan.IndexOf($_) } -Descending)) {
        $annul += "Move-Item -LiteralPath '$($a.Destination -replace "'", "''")' -Destination '$($a.Source -replace "'", "''")'"
    }
    $annul | Set-Content -Encoding UTF8 (Join-Path $Journal "ANNULER_$Stamp.ps1")

    # LISEZ-MOI de 00_A-trier : une ligne d'explication par fichier
    $lignes = @("Fichiers à classer à la main - $Stamp", '')
    $Plan | Where-Object { $_.Type -eq 'à trier' -and $_.Statut -eq 'fait' } | ForEach-Object {
        $lignes += '{0}  —  {1}  (origine : {2})' -f (Split-Path $_.Destination -Leaf), $_.Raison, $_.Source
    }
    $lignes | Set-Content -Encoding UTF8 (Join-Path $Root "$ATrier\LISEZ-MOI_A-trier.txt")

    # Liste des éléments à supprimer après accord
    $Plan | Where-Object { $_.Type -eq 'à valider' -and $_.Statut -eq 'fait' } |
        Select-Object @{ N = 'Element'; E = { Split-Path $_.Destination -Leaf } }, Raison, Source, Destination |
        Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' (Join-Path $Root "$Corbeille\LISTE_A-VALIDER-SUPPRESSION.csv")
}

$Plan | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' (Join-Path $Journal "plan_$Stamp.csv")
if ($Laisses) { $Laisses | Select-Object FullName | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' (Join-Path $Journal "laisses_en_place_$Stamp.csv") }

# --- Résumé ---
Write-Host ''
$Plan | Group-Object { ($_.Destination.Substring($Root.Length + 1) -split '\\')[0..1] -join '\' } | Sort-Object Name |
    ForEach-Object { '{0,5}  {1}' -f $_.Count, $_.Name } | Write-Host
Write-Host ''
Write-Host ("Renommés : {0}   À trier : {1}   À valider (suppression) : {2}   Erreurs : {3}" -f `
    @($Plan | Where-Object Renomme -eq 'oui').Count, @($Plan | Where-Object Type -eq 'à trier').Count,
    @($Plan | Where-Object Type -eq 'à valider').Count, @($Plan | Where-Object Statut -eq 'ERREUR').Count)
if ($Execute) {
    $reste = @(Get-ChildItem -LiteralPath $Downloads -Force)
    if ($reste.Count -eq 0) { Write-Host 'Téléchargements est vide.' -ForegroundColor Green }
    else { Write-Host "Il reste $($reste.Count) élément(s) dans Téléchargements (voir colonne Erreur du plan)." -ForegroundColor Red }
}
Write-Host "Plan détaillé : $Journal"
Invoke-Item $Journal
