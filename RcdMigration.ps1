<#
.SYNOPSIS
    Prépare le passage de Restricted SharePoint Search (RSS) à Restricted Content Discovery (RCD).

.DESCRIPTION
    Par défaut, l'outil est en LECTURE SEULE : il ne modifie rien dans le locataire.

    Il lit :
      1. Le mode de RSS (Get-SPOTenantRestrictedSearchMode) et sa liste autorisée
         (Get-SPOTenantRestrictedSearchAllowedList).
      2. Chaque site SharePoint (Get-SPOSite), puis l'état de RCD site par site
         (propriété RestrictContentOrgWideSearch de Get-SPOSite -Identity).

    Avec RSS activée, tout site absent de la liste autorisée est tenu à l'écart de la recherche à
    l'échelle de l'organisation et de Copilot. Après la retraite de RSS (31 janvier 2027), ces sites
    redeviennent détectables, car Microsoft ne migre pas les réglages. L'outil dresse donc la liste
    des sites candidats à RCD et classe chaque site :
      * À MASQUER : RSS les tient à l'écart et RCD n'y est pas active. Candidat à RCD.
      * DÉJÀ MASQUÉ : RCD est déjà active.
      * RSS AUTORISÉ : dans la liste autorisée de RSS, donc déjà détectable.
      * RSS INACTIVE : RSS est désactivée, rien à migrer pour ce site.
      * À VÉRIFIER : l'état de RCD n'a pas pu être lu.
      * EXCLU : écarté par -ExcludeUrlPattern ou -ExcludeTemplate.

    Avec -Apply, l'outil active RCD (Set-SPOSite -RestrictContentOrgWideSearch $true) sur les sites
    « À MASQUER », jusqu'à -MaxSites. Chaque modification passe par ShouldProcess : -WhatIf et
    -Confirm sont donc pris en charge. Microsoft Learn précise que -WhatIf ne fonctionne pas sur
    Set-SPOSite lui-même : l'outil le gère avant d'appeler la cmdlet.

    Learn demande d'utiliser RCD de façon sélective : trop de sites masqués réduisent la qualité des
    réponses de Copilot. Relisez la liste avant d'utiliser -Apply.

.PARAMETER AdminUrl
    URL du centre d'administration SharePoint, par exemple https://contoso-admin.sharepoint.com.

.PARAMETER SkipConnect
    N'ouvre pas de connexion (utile si Connect-SPOService est déjà fait).

.PARAMETER OutputFolder
    Dossier où écrire le CSV et le rapport HTML. Par défaut : le dossier courant.

.PARAMETER ExcludeUrlPattern
    Motifs génériques (-like) d'URL de sites à exclure, par exemple '*/sites/Archives*'.

.PARAMETER ExcludeTemplate
    Modèles de sites à exclure. Par défaut : modèles système (centre de recherche, hôte de sites
    personnels, catalogue d'applications, redirection). À ajuster selon votre locataire.

.PARAMETER Apply
    Active RCD sur les sites « À MASQUER ». Sans ce paramètre, rien n'est modifié.

.PARAMETER MaxSites
    Nombre maximal de sites modifiés par exécution avec -Apply. Par défaut : 25.

.PARAMETER PassThru
    Retourne aussi les objets dans le pipeline (sinon un résumé est écrit dans le flux d'information).

.EXAMPLE
    .\RcdMigration.ps1 -AdminUrl https://contoso-admin.sharepoint.com -OutputFolder .\rapports

    Inventaire en lecture seule avec CSV et rapport HTML.

.EXAMPLE
    .\RcdMigration.ps1 -SkipConnect -Apply -WhatIf

    Montre les sites où RCD serait activée, sans rien modifier.

.EXAMPLE
    .\RcdMigration.ps1 -SkipConnect -Apply -MaxSites 5 -ExcludeUrlPattern '*/sites/Projet*'

    Active RCD sur au plus cinq sites, après confirmation pour chacun.

.NOTES
    Prérequis : module Microsoft.Online.SharePoint.PowerShell à jour, rôle d'administrateur
    SharePoint, licence Microsoft Copilot et SharePoint Advanced Management pour RCD.
    Projet : https://github.com/CharlesTrivium365/RcdMigration
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Connect')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Connect')]
    [ValidatePattern('^https://[^/\s]+-admin\.sharepoint\.[a-z.]+/?$')]
    [string]$AdminUrl,

    [Parameter(Mandatory, ParameterSetName = 'Connected')]
    [switch]$SkipConnect,

    [ValidateScript({ Test-Path -Path $_ -PathType Container })]
    [string]$OutputFolder = (Get-Location).Path,

    [string[]]$ExcludeUrlPattern = @(),

    [string[]]$ExcludeTemplate = @('SRCHCEN#0', 'SPSMSITEHOST#0', 'APPCATALOG#0', 'REDIRECTSITE#0'),

    [switch]$Apply,

    [ValidateRange(1, 1000)]
    [int]$MaxSites = 25,

    [switch]$PassThru
)

#Requires -Modules Microsoft.Online.SharePoint.PowerShell

Set-StrictMode -Version 2.0

function ConvertTo-HtmlText {
    param([string]$Text)
    [System.Net.WebUtility]::HtmlEncode($Text)
}

function ConvertTo-SafeCsvCell {
    param([string]$Text)
    # Protège contre l'interprétation d'une formule par Excel.
    if ($Text -match '^[=+\-@]') { return "'" + $Text }
    $Text
}

function Get-NormalizedUrl {
    param([string]$Url)
    $Url.Trim().TrimEnd('/').ToLowerInvariant()
}

function Get-UrlFromItem {
    param($Item)
    if ($null -eq $Item) { return $null }
    if ($Item -is [string]) { return $Item }
    $property = $Item.PSObject.Properties['Url']
    if ($property) { return [string]$property.Value }
    [string]$Item
}

function Write-Utf8Bom {
    param([string]$Path, [string]$Text)
    [System.IO.File]::WriteAllText($Path, $Text, (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $true))
}

if (-not $SkipConnect) {
    Connect-SPOService -Url $AdminUrl
}

# 1. RSS : mode et liste autorisée
$rssMode = 'Inconnu'
$allowedUrls = @{}
try {
    $rssMode = [string](Get-SPOTenantRestrictedSearchMode -ErrorAction Stop)
}
catch {
    Write-Warning "Mode de RSS illisible (les cmdlets de RSS cessent de fonctionner le 28 février 2027) : $($_.Exception.Message)"
}
$rssEnabled = $rssMode -match 'Enabled'

if ($rssEnabled) {
    try {
        foreach ($item in @(Get-SPOTenantRestrictedSearchAllowedList -ErrorAction Stop)) {
            $url = Get-UrlFromItem -Item $item
            if ($url) { $allowedUrls[(Get-NormalizedUrl -Url $url)] = $true }
        }
    }
    catch {
        Write-Warning "Liste autorisée de RSS illisible : $($_.Exception.Message)"
    }
}
Write-Information -MessageData "RSS : mode $rssMode, $($allowedUrls.Count) site(s) dans la liste autorisée." -InformationAction Continue

# 2. Sites SharePoint (les sites OneDrive sont hors portée de RCD et exclus par défaut de Get-SPOSite)
$sites = @(Get-SPOSite -Limit All)
$rcdPropertyMissing = $false
$results = New-Object -TypeName System.Collections.Generic.List[object]
$index = 0

foreach ($site in $sites) {
    $index++
    $url = [string]$site.Url
    Write-Progress -Activity 'Lecture de l''état de RCD' -Status $url -PercentComplete ([int](100 * $index / $sites.Count))

    # Le site du centre d'administration (nom d'hôte en -admin.sharepoint.) n'est jamais un candidat.
    $excluded = ($url -match '^https://[^/]+-admin\.sharepoint\.') -or ($ExcludeTemplate -contains [string]$site.Template)
    foreach ($pattern in $ExcludeUrlPattern) { if ($url -like $pattern) { $excluded = $true } }

    $rcd = $null
    $status = 'EXCLU'
    if (-not $excluded) {
        # Get-SPOSite -Limit ne remplit pas toutes les propriétés : on relit chaque site par son URL.
        try {
            $detail = Get-SPOSite -Identity $url -ErrorAction Stop
            $property = $detail.PSObject.Properties['RestrictContentOrgWideSearch']
            if ($property) { $rcd = [bool]$property.Value } else { $rcdPropertyMissing = $true }
        }
        catch {
            Write-Warning "Lecture impossible pour $url : $($_.Exception.Message)"
        }

        $inAllowedList = $allowedUrls.ContainsKey((Get-NormalizedUrl -Url $url))
        if ($null -eq $rcd) { $status = 'À VÉRIFIER' }
        elseif ($rcd) { $status = 'DÉJÀ MASQUÉ' }
        elseif (-not $rssEnabled) { $status = 'RSS INACTIVE' }
        elseif ($inAllowedList) { $status = 'RSS AUTORISÉ' }
        else { $status = 'À MASQUER' }
    }

    $results.Add([pscustomobject]@{
            Statut      = $status
            Url         = $url
            Titre       = [string]$site.Title
            Modele      = [string]$site.Template
            RcdActive   = $rcd
            ListeRss    = $allowedUrls.ContainsKey((Get-NormalizedUrl -Url $url))
            Action      = ''
        })
}
Write-Progress -Activity 'Lecture de l''état de RCD' -Completed

if ($rcdPropertyMissing) {
    Write-Warning 'La propriété RestrictContentOrgWideSearch est absente : mettez à jour le module Microsoft.Online.SharePoint.PowerShell.'
}

# 3. Modification (seulement avec -Apply)
if ($Apply) {
    $done = 0
    foreach ($row in @($results | Where-Object -FilterScript { $_.Statut -eq 'À MASQUER' })) {
        if ($done -ge $MaxSites) { $row.Action = "Ignoré (limite de $MaxSites atteinte)"; continue }
        if ($PSCmdlet.ShouldProcess($row.Url, 'Activer Restricted Content Discovery')) {
            try {
                Set-SPOSite -Identity $row.Url -RestrictContentOrgWideSearch $true -ErrorAction Stop
                $row.Action = 'Activé (propagation en cours)'
                $row.RcdActive = $true
            }
            catch {
                $row.Action = "Échec : $($_.Exception.Message)"
            }
        }
        else {
            $row.Action = 'Simulé'
        }
        $done++
    }
}

# 4. Sorties : CSV et rapport HTML (UTF-8 avec BOM)
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath = Join-Path -Path $OutputFolder -ChildPath "RcdMigration-$stamp.csv"
$htmlPath = Join-Path -Path $OutputFolder -ChildPath "RcdMigration-$stamp.html"

$csvRows = $results | Select-Object -Property Statut, Url, @{ Name = 'Titre'; Expression = { ConvertTo-SafeCsvCell -Text $_.Titre } }, Modele, RcdActive, ListeRss, Action
Write-Utf8Bom -Path $csvPath -Text ((($csvRows | ConvertTo-Csv -NoTypeInformation) -join "`r`n") + "`r`n")

$order = 'À MASQUER', 'À VÉRIFIER', 'DÉJÀ MASQUÉ', 'RSS AUTORISÉ', 'RSS INACTIVE', 'EXCLU'
$counts = foreach ($name in $order) {
    $n = @($results | Where-Object -FilterScript { $_.Statut -eq $name }).Count
    "<li><strong>$(ConvertTo-HtmlText -Text $name)</strong> : $n</li>"
}
$body = foreach ($row in ($results | Sort-Object -Property @{ Expression = { $order.IndexOf($_.Statut) } }, Url)) {
    "<tr><td>$(ConvertTo-HtmlText -Text $row.Statut)</td><td>$(ConvertTo-HtmlText -Text $row.Url)</td><td>$(ConvertTo-HtmlText -Text $row.Titre)</td><td>$(ConvertTo-HtmlText -Text $row.Modele)</td><td>$(ConvertTo-HtmlText -Text $row.Action)</td></tr>"
}
$html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>RcdMigration : de RSS à RCD</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 2rem; color: #1b1b1b; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #c8c8c8; padding: 4px 8px; text-align: left; font-size: 0.9rem; }
th { background: #243a5e; color: #fff; }
</style>
</head>
<body>
<h1>RcdMigration : de Restricted SharePoint Search à Restricted Content Discovery</h1>
<p>Généré le $(Get-Date -Format 'yyyy-MM-dd HH:mm'). Mode de RSS : <strong>$(ConvertTo-HtmlText -Text $rssMode)</strong>. Sites dans la liste autorisée : $($allowedUrls.Count). Mode de l'outil : $(if ($Apply) { 'modification demandée (-Apply)' } else { 'lecture seule' }).</p>
<ul>
$($counts -join "`n")
</ul>
<table>
<tr><th>Statut</th><th>URL</th><th>Titre</th><th>Modèle</th><th>Action</th></tr>
$($body -join "`n")
</table>
</body>
</html>
"@
Write-Utf8Bom -Path $htmlPath -Text $html

Write-Information -MessageData "CSV : $csvPath" -InformationAction Continue
Write-Information -MessageData "HTML : $htmlPath" -InformationAction Continue

$summary = $results | Group-Object -Property Statut | Select-Object -Property Name, Count
foreach ($line in $summary) { Write-Information -MessageData "$($line.Name) : $($line.Count)" -InformationAction Continue }

if ($PassThru) { $results }
