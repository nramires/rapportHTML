Import-Module .\rapportHTML.ps1

#--------------------------------
# CREATION D'UN OBJET RAPPORTHTML
#--------------------------------

$rapport = [rapportHTML]::new("Titre principal du rapport")


#--------------------------------------
# CREATION D'UN TABLEAU AU FORMAT LISTE
#--------------------------------------

# on prépare un pscustomobjet pour stocker les informations techniques concernant le script
# ... le nom des propriétés du pscustomobject deviendront le nom des colonnes d'un tableau du rapport
# ... les valeurs des propriétés deviendront les valeurs de la propriété TEXTE des cellules de ce tableau.
# ... le TEXTE d'une cellule sera interprété en html. On peut donc y insérer des balises.
# ...... la Cellule de la Colonne LOG est un exemple. Elle affichera le texte "Exemples.log" et sur un clic ouvrira un nouvel onglet affichant une log.
$logfilename = "Exemples.log"
$logfolder = ".\LOG"
$donnéesInformations = [pscustomobject]@{
    "SERVER" = $env:computername
    "SCRIPT" = $PSCommandPath
    "START"  = Get-Date -uformat "%H:%M:%S"
    "RUNAS"  = "$env:USERDOMAIN\$env:USERNAME"
    "LOG"    = "<a href=""$logfolder\$logfilename"" target=""_blank"">$logfilename</a>"}

# on veut faire apparaître une ligne supplémentaire d'information uniquement si on approche ou dépasse la date d'expiration d'un certificat
$dateCertificate = Get-Date (Get-ChildItem Cert:\CurrentUser\My\ | ? Thumbprint -like "64c3638d9c0d60058df73e17b1cd858561f20d72" | % NotAfter)
#$dateCertificate = (get-date).AddDays(30) # cette ligne sert uniquement à forcer le résultat pour faire des captures d'écrans pour la documentation
if ($dateCertificate -le (get-date).AddDays(30)) {
    $donnéesInformations | Add-Member "CERTIF" ""}

# création d'un tableau au format Liste (dernier param = 'liste')
# ... Premier param = le titre du tableau
# ... deuxième param = les noms des colonnes.
# ...... Si on passe un pscustomobject, les noms des propriétés seront considérés comme noms des colonnes.
# ...... Si on passe un ordered, les noms des clés seront considérés comme noms des colonnes.
# ... troisième param : un string 'colonne' ou 'liste' pour indiquer le format du tableau
$tabInformations = $rapport.CréerTableau("Titre d'un tableau au format Liste", $donnéesInformations, [FormatTableau]::Liste) # premier tableau créé

$ligne = $tabInformations.Lignes[0]

# scriptblock codant la mise en forme d'une cellule en fonction de la date d'expiration du certificat
$setFormatCertif = {
    if ($dateCertificate -le (Get-Date).AddDays(30)) {
        $cellule.Style = [StyleHTML]::INFO
        $cellule.Texte = "<div align=""center""><b>Le certificat expire le $(Get-Date $dateCertificate -Format "dd/MM/yyyy hh:mm:ss"). Penser à le remplacer.</b></div>"}
    if ($dateCertificate -le (Get-Date).AddDays(15)) {
        $cellule.Style = [StyleHTML]::ALERTE
        $cellule.Texte = "<div align=""center""><b>/!\ Le certificat expire le $(Get-Date $dateCertificate -Format "dd/MM/yyyy hh:mm:ss"). Remplacez-le !</b></div>"}
    if ($dateCertificate -le (Get-Date)) {
        $cellule.Style = [StyleHTML]::ERREUR
        $cellule.Texte = "<div align=""center""><b>Le certificat est expiré. Connexion Exchange Online impossible.</b></div>"}
}

# on récupère l'instance de la [Cellule] de la colonne CERTIF de la première ligne du premier tableau stocké dans notre instance [rapportHTML]
$cellule = $rapport.Tableaux[0].Lignes[0].CERTIF # Lignes[0] renvoie un [pscustomobject] dont les propriétés sont les noms de colonnes, et les valeurs des [Cellule]
$cellule = $tabInformations.Lignes[0].CERTIF # équivalent à la ligne prédédente, si on a pris soin de stocker la tableau au préalable
$cellule = $ligne.CERTIF # équivalent à la ligne précédente, si on a pris soin de stocker la ligne au préalable

# on exécute la mise en forme de la cellule
Invoke-Command -ScriptBlock $setFormatCertif


#----------------------------------------
# CREATION D'UN TABLEAU AU FORMAT COLONNE
#----------------------------------------

# avec les mêmes données on va créer un tableau cette fois au format Colonne
$tabDonnées = $rapport.CréerTableau("Titre d'un tableau au format Colonne", $donnéesInformations, [FormatTableau]::Colonne) # deuxième tableau créé
$cellule = $tabDonnées.Lignes[0].CERTIF # première (et unique) ligne du DEUXIEME tableau
Invoke-Command -ScriptBlock $setFormatCertif

#---------------------------------------------------------------
# ILLUSTRATION STYLES PREDEFINIS & ALIGNEMENT HORIZONTAL CELLULE
#---------------------------------------------------------------

# on crée un tableau vide (sans ligne) car on ne passe en 2ème paramètre que des chaînes de caractères.
$tabStyles = $rapport.CréerTableau("Illustration des styles prédéfinis de format de cellule, et des trois alignements horizontaux de texte",[StyleHTML]::GetNames([StyleHTML]),"Colonne")
$alignements = [enum]::GetValues([alignement]) # alignement est une enum fourni par le module, et vaut : left,center, right
for ($i=0; $i -lt $alignements.count; $i++) {
    $ligne = $tabStyles.CréerLigne() # création d'une ligne vide
    foreach ($nom in $tabStyles.NomsColonnes) {
        $cellule = $ligne.$nom # rappel, on accède à une cellule de ligne à partir de son nom de colonne
        $cellule.Texte = 'NA' 
        $cellule.Style = [StyleHTML]::$nom
        $cellule.Alignement = $alignements[$i] # alignement horizontal du texte
    }
}

#-----------------------------------------
# ILLUSTRATION SURTITRES & CELLULE ETENDUE
#-----------------------------------------

$tabSurtitres = $rapport.CréerTableau("Illustration des surtitres et d'une cellule étendue", @("info1","info2","info3","info4","autre1","autre2","résultat1","résultat2"), "colonne")
$tabSurtitres.CréerSurtitres("données en entrée",1,4)
$tabSurtitres.CréerSurtitres("données en sortie",7,2)
$tabSurtitres.CréerLigne(@{
    # cette fois on ne crée pas une ligne vide, mais on passe en paramètre une structure permettant d'avoir des noms de colonnes et des valeurs associées
    # contrairement à la création de Tableau, la création de Ligne accepte aussi les Hastable, car la présence et l'ordre d'apparition des colonnes ne sont pas importants
    "info1" = "Cellule 'INFO1' centrée et étendue sur toutes les colonnes, avec style 'INFO'"}) | Out-Null
$tabSurtitres.Lignes.info1.Alignement = "center"
$tabSurtitres.Lignes.info1.Style = "info"
$tabSurtitres.Lignes.info1.Etendue = $tabSurtitres.NombreColonnes

#---------------------------------------------------------
# ILLUSTRATION ALIGNEMENT HORIZONTAL DU TEXTE VIA COLONNES
#---------------------------------------------------------

$tabAlignement = $rapport.CréerTableau("Illustration de l'alignement horizontal du texte des cellules via les colonnes", @("explication","colonne 1", "colonne 2", "colonne 3"), "Colonne")

$ligne = $tabAlignement.CréerLigne(@{
    "explication" = "Ligne créée sans aucune modification de l'alignement"
    "colonne 1" = "texte"
    "colonne 2" = "texte"
    "colonne 3" = "texte"})

# on peut modifier l'alignement par défaut d'une ou de plusieurs colonnes par leurs numéros d'ordre ou leurs noms
$tabAlignement.ModifierAlignementParDéfautDesColonnes([alignement]::right, (3,4))

$ligne = $tabAlignement.CréerLigne(@{
    "explication" = "Ligne créée :<br> - APRES avoir modifié l'alignement PAR DEFAUT des 2 dernières colonnes en 'right', mais ... <br> - AVANT de centrer la dernière colonne<br>-> on voit quand même l'effet du centrage sur la dernière colonne"
    "colonne 1" = "texte"
    "colonne 2" = "texte"
    "colonne 3" = "texte"})

# on modifie l'alignement horizontal du texte des cellules existantes de la colonne 4
$tabAlignement.ModifierAlignementDesColonnes("center",4)

$ligne = $tabAlignement.CréerLigne(@{
    "explication" = "Ligne créée APRES avoir centré la dernière colonne<br> -> on retrouve l'alignement par défaut sur la dernière colonne"
    "colonne 1" = "texte"
    "colonne 2" = "texte"
    "colonne 3" = "texte"})


#----------------------------------
# EXEMPLES DE STYLES PERSONNALISES
#----------------------------------

$tabPerso = $rapport.CréerTableau("Exemples de styles personnalisés de format de cellule", @("Colonne 1","Colonne 2"), "Colonne")
$tabPerso.CréerLigne([ordered]@{
    "Colonne 1" = "style perso 1"
    "Colonne 2" = "style perso 2"}) | Out-Null

# on défini un style personnalisé de format cellule à l'aide de la classe [StylePerso]
$datastyle = [stylePerso]::new()
$datastyle.background = "#0000FF" # bleu
$datastyle.color = "#FFFFFF" 
$datastyle.'font-weight' = "bold"

# on peut ensuite l'ajouter dans les styles disponibles dans le rapport, et on récupère le nom du nouveau style dans la variable $monstyle_1
$monstyle_1 = $rapport.AjouterStylePerso($datastyle)

# on crée un deuxième style perso en modifiant légèrement le précédent modèle
$datastyle.color = "#FF00FF" # rose
$monstyle_2 = $rapport.AjouterStylePerso($datastyle)

# on applique nos styles persos sur des cellules
$tabPerso.Lignes[0].'Colonne 1'.Style = $monstyle_1
$tabPerso.Lignes[0].'Colonne 2'.Style = $monstyle_2

#---------------------------
# GENERATION DU FICHIER HTML
#---------------------------

$rapport.GénérerRapport() | Out-File ($PSCommandPath -replace ".ps1",".html")


# SIG # Begin signature block
# MIIFowYJKoZIhvcNAQcCoIIFlDCCBZACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUi/EzgRvQIuDcZ+Uwjlv6Ltt5
# tSqgggMsMIIDKDCCAhCgAwIBAgIQFpiwcPSA9plJJVYjwMrpxzANBgkqhkiG9w0B
# AQsFADAsMSowKAYDVQQDDCFQb3dlclNoZWxsIENvZGUgU2VsZiBTaWduaW5nIENl
# cnQwHhcNMjMwNjIzMDYzNDQwWhcNMjQwNjIzMDY1NDQwWjAsMSowKAYDVQQDDCFQ
# b3dlclNoZWxsIENvZGUgU2VsZiBTaWduaW5nIENlcnQwggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQDG20g7ZP23CxaqY0TA5jv0KUIS+jbXA7n1nwoQGHg/
# n4Gq5QNYwIzP6OSEtnLBuxFO/c4Y+XLOtCiMoFu/JPgf+wCtDwxvv1EMcvwCS7EG
# EDl4bH80k0LkKtmHJho+cbPlA7ZznIWqxICwY9wDRMDu4yZvfcW1uHsA+/2uWWuz
# hXkdD9m5T1Pg2TQ4nGjYRgNTACDchloWd+FwpSCdXUt0HJ1tE06MDOWYe1BJcBDZ
# 6n1Ul3M02MXv+cP4dv2IM44j8xbqh6tiVxBK1aFWtTkcFN6ZYBH9h5ReShhg11CF
# RC7lAMdEfV57KMvNgXIrdDGZaIJC7NX6HA+ACzIeTd7BAgMBAAGjRjBEMA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUm3xCA6yP
# qjsvPDcrwIYZnfxob6MwDQYJKoZIhvcNAQELBQADggEBADWO0HX0MWILYmLhCCa0
# PJ5eYERGDc36YvwNIMaV0EFf1oCFq0tl6OZ0J5yFx43qvClL27zkujKKHHna/fzG
# Vv6gKrr/vNIEYHRJqVxeYP7rcbffXo22yvjEaibq+DtGdMIpphQviWUl5x7Gh2sM
# 2nyLjQd+JbAJvoTb21VdsVqxDQkqox6bMXv3EflIZWU8AT7Z+yuwwvI/x+vfo2tm
# KyP1EEuKoNwm6pM9Igv1shSJzt1/5GdMNIy5kWEjASHDgDzcF31xXqVFSrMwhS58
# 2qOgfV3SeanhyrxKatSj28zHtzThVOeHsomokcLRepswkGTKhhi8dR6dHDjxRivp
# AlgxggHhMIIB3QIBATBAMCwxKjAoBgNVBAMMIVBvd2VyU2hlbGwgQ29kZSBTZWxm
# IFNpZ25pbmcgQ2VydAIQFpiwcPSA9plJJVYjwMrpxzAJBgUrDgMCGgUAoHgwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
# 9vkolXlw+fsDNOUzd8sJy813YM0wDQYJKoZIhvcNAQEBBQAEggEAcXUiDo7a+EXo
# nnQ8nrnl3T9+Zd99P2ELpKNcaiAxieMiQEAnahlRqzLsP+PwHciRiHbQ3LBAKGCg
# 3RKVexIXvxGKlBD6Qs6Guplq8dTDipO9A7NCIr+i8rG+5hPkJRaCFb6ks8MnJLjA
# cAPK+/LKJX3Rga+85iOFf44FObphtUkDhHAyVGxYaykufol5huhZCNFCbV1vCkha
# xVG6d74//khWSbUHJBYvPq8koNWHkjwp/g3W0PTVMAyuwyhKheJpQSw4In7hTv5N
# DGWk5c3agSU7Imt2+Dt7tLAsyCdb54dyKuF9B7vaLINkXfxI4ISwqPrq8ykJw2i+
# yr+M4ZTSnQ==
# SIG # End signature block
