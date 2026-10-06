# RcdMigration

Outil de préparation au passage de **Restricted SharePoint Search (RSS)** à **Restricted Content Discovery (RCD)** dans SharePoint Online.

RSS cesse de fonctionner le **31 janvier 2027** (nouvelles activations bloquées depuis le 31 juillet 2026, cmdlets PowerShell retirées le 28 février 2027) selon l’avis MC1395311 du Centre de messages. Microsoft **ne migre aucun réglage** : un site que RSS tenait à l’écart de la recherche et de Copilot redevient détectable après la retraite, sauf si vous le masquez vous-même avec RCD.

RcdMigration répond à une question : **quels sites faut-il examiner avant la fin de RSS ?**

📝 Article du blogue : [blog.trivium365.com](https://blog.trivium365.com) (l’article « Restricted SharePoint Search prend sa retraite » accompagne cet outil).

## Ce que fait l’outil

1. Lit le mode de RSS (`Get-SPOTenantRestrictedSearchMode`) et sa liste autorisée (`Get-SPOTenantRestrictedSearchAllowedList`).
2. Liste les sites (`Get-SPOSite -Limit All`) puis relit chaque site par son URL pour obtenir `RestrictContentOrgWideSearch`, l’état de RCD.
3. Classe chaque site :

| Statut | Signification |
|---|---|
| **À MASQUER** | RSS le tient à l’écart (absent de la liste autorisée) et RCD n’y est pas active : candidat à RCD |
| **DÉJÀ MASQUÉ** | RCD est déjà active |
| **RSS AUTORISÉ** | Dans la liste autorisée de RSS, donc déjà détectable |
| **RSS INACTIVE** | RSS est désactivée : rien à migrer pour ce site |
| **À VÉRIFIER** | L’état de RCD n’a pas pu être lu |
| **EXCLU** | Écarté par `-ExcludeUrlPattern` ou `-ExcludeTemplate` |

Il produit un CSV et un rapport HTML (UTF-8 avec BOM) et, avec `-PassThru`, des objets PowerShell.

**Par défaut, l’outil ne modifie rien.**

## Modification : le paramètre -Apply

Avec `-Apply`, l’outil exécute `Set-SPOSite -Identity <url> -RestrictContentOrgWideSearch $true` sur les sites « À MASQUER », au plus `-MaxSites` par exécution (25 par défaut). Chaque site passe par `ShouldProcess` : `-WhatIf` et `-Confirm` sont pris en charge.

Microsoft Learn précise que le commutateur `-WhatIf` ne fonctionne pas sur `Set-SPOSite` lui-même. L’outil décide donc de la simulation avant d’appeler la cmdlet : avec `-WhatIf`, `Set-SPOSite` n’est jamais appelée.

Learn demande aussi d’utiliser RCD **de façon sélective** : trop de sites masqués réduisent la qualité des réponses de Copilot. RCD est présenté comme un contrôle temporaire le temps de revoir les permissions. La liste « À MASQUER » est un point de départ, pas une décision. Relisez-la avec les propriétaires de sites.

Pour annuler sur un site : `Set-SPOSite -Identity <url> -RestrictContentOrgWideSearch $false`.

## Prérequis

- PowerShell 7 ou Windows PowerShell 5.1
- Module `Microsoft.Online.SharePoint.PowerShell` à jour (la propriété `RestrictContentOrgWideSearch` doit exister)
- Rôle d’administrateur SharePoint (ou équivalent)
- Pour RCD : licence Microsoft Copilot et SharePoint Advanced Management (voir les prérequis de SAM sur Learn)

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

## Utilisation

Inventaire en lecture seule :

```powershell
.\RcdMigration.ps1 -AdminUrl https://contoso-admin.sharepoint.com -OutputFolder .\rapports
```

Déjà connecté avec `Connect-SPOService` ? Ajoutez `-SkipConnect`.

Simuler l’activation de RCD :

```powershell
.\RcdMigration.ps1 -SkipConnect -Apply -WhatIf
```

Activer RCD sur cinq sites au plus, en excluant des sites de projet :

```powershell
.\RcdMigration.ps1 -SkipConnect -Apply -MaxSites 5 -ExcludeUrlPattern '*/sites/Projet*'
```

## Exemple de résultat

```
Statut        Url                                         Titre            Modele   RcdActive ListeRss
------        ---                                         -----            ------   --------- --------
À MASQUER     https://contoso.sharepoint.com/sites/RH     Ressources hum.  STS#3    False     False
À MASQUER     https://contoso.sharepoint.com/sites/Fin    Finances         GROUP#0  False     False
DÉJÀ MASQUÉ   https://contoso.sharepoint.com/sites/Dir    Direction        STS#3    True      False
RSS AUTORISÉ  https://contoso.sharepoint.com/sites/Vente  Ventes           GROUP#0  False     True
```

## Limites connues

- L’outil est livré sans avoir été exécuté contre un vrai locataire : validez-le d’abord sur un locataire de test ou avec `-WhatIf`.
- Le format exact de la sortie de `Get-SPOTenantRestrictedSearchAllowedList` n’est pas détaillé sur Learn : l’outil accepte des URL en texte ou des objets avec une propriété `Url`. Vérifiez le nombre de sites de la liste autorisée affiché au début de l’exécution.
- La relecture site par site est lente dans un grand locataire, car `Get-SPOSite -Limit` ne remplit pas toutes les propriétés.
- La liste des modèles exclus par défaut (`SRCHCEN#0`, `SPSMSITEHOST#0`, `APPCATALOG#0`, `REDIRECTSITE#0`) vise des sites système : ajustez-la selon votre locataire. Le site du centre d’administration est toujours exclu.
- RCD ne s’applique pas à OneDrive, ne change pas les permissions et peut prendre plus d’une semaine à se propager sur un site de plus de 500 000 éléments.
- Learn contient deux phrases qui se contredisent en partie : RCD vise « y compris les fichiers ayant récemment interagi », mais les utilisateurs peuvent encore découvrir le contenu qu’ils possèdent ou avec lequel ils ont récemment interagi. Testez avec un compte réel sur un site pilote.
- Les cmdlets de RSS cessent de fonctionner le 28 février 2027 : après cette date, l’outil signale que le mode de RSS est illisible.

## Sources

- [MC1395311 : retraite de Restricted SharePoint Search](https://mc.merill.net/message/MC1395311) (Centre de messages)
- [Restrict discovery of SharePoint sites and content](https://learn.microsoft.com/en-us/sharepoint/restricted-content-discovery) (Microsoft Learn)
- [Restricted SharePoint Search Admin PowerShell Scripts](https://learn.microsoft.com/en-us/sharepoint/restricted-sharepoint-search-admin-scripts) (Microsoft Learn)
- [Get-SPOSite](https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/get-sposite) et [Set-SPOSite](https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/set-sposite) (Microsoft Learn)
- [SharePoint Advanced Management features with a Copilot license](https://learn.microsoft.com/en-us/sharepoint/sharepoint-advanced-management-features-copilot-license) (Microsoft Learn)

## Licence

MIT. Publié par Charles Jenkins, [blog.trivium365.com](https://blog.trivium365.com).
