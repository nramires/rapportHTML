<#
    .SYNOPSIS
        Librairie d'aide à la représentation au format HTML d'un ou de plusieurs tableaux de données sous forme de rapport.
    .DESCRIPTION
        Permet de :
            - créer un objet [rapportHTML] qui est à la base de votre gestion du futur rapport html.
            - créer des objets [Tableau], représentant les tableaux de données à présenter dans le rapport html.
            - choisir de présenter le tableau au format colonne ou liste.
            - créer des surtitres qui peuvent être étendus sur plusieurs colonnes.
            - créer des lignes vides, partiellement ou totalement renseignées.
            - utiliser des styles prédéfinis ou personnalisés pour :
                - modifier l'apparence d'une cellule de données individuellement, ou pour une ou plusieurs colonnes d'un coup.
                - modifier l'apparence par défaut des cellules d'une ou de plusieurs colonnes d'un coup.
                - modifier l'apparence des cellules d'entête de colonnes ou de surtitres.
            - modifier l'alignement horizontal du texte d'une cellule de données, individuellement ou pour une ou plusieurs colonne d'un coup.
            - modifier l'alignement horizontal par défaut du texte d'une cellule de données pour une ou plusieurs colonnes d'un coup.
            - modifier l'alignement horizontal du texte des cellules d'entête de colonnes ou de surtitres.
            - modifier individuellement une cellule de données : son texte, l'alignement horizontal du texte, son style, son étendue éventuelle sur plusieurs colonne de la même ligne.
            - générer automatiquement le fichier html correspondant à l'objet [rapportHTML]
        Avantages :
            - uniformisez l'apparence de vos rapports.
            - à partir de données brutes, créez facilement un ou plusieurs tableaux.
            - ajoutez, supprimez ou changez l'ordre de vos colonnes facilement pendant votre phase de développement de votre script sans effets de bords.
            - utilisez des balises html dans le texte de vos cellules de données pour aller encore plus loin dans la personnalisation de vos rapports.
    .EXAMPLE
        Import-Module .\rapportHTML.ps1

        Commande préalable à l'utilisation de cette librairie.
    .EXAMPLE
        $r = [rapporthtml]::new("Nouveau rapport")

        Crée un nouveau [rapportHTML] dont le titre général est 'Nouveau rapport'.
        On stocke ce rapport dans la variable $r pour la suite.
    .EXAMPLE
        $t = $r.CréerTableau("tableau Liste avec surtitre", [ordered]@{"colonne 1"="val1";"colonne 2"="val2";"colonne 3"="val3";"colonne 4"="val4"}, "liste")
        Crée un nouveau [Tableau] dont le titre spécifique est 'tableau Liste avec surtitre'.
        On pourrait passer en deuxième paramètre uniquement le nom des colonnes que l'on veut pour le tableau. Mais ici on passe une structure associant des valeurs à des noms de colonne. Une ligne sera automatiquement créée en même temps que le tableau.
        On stocke ce tableau dans la variable $t pour la suite.

        >$t.CréerSurtitres("surtitre",2,2)
        On crée un surtitre.

        >$t.StyleSurtitre = [styleHTML]::TITRE_Dégradé_Bleu_Gauche
        >$t.StyleTitre = [styleHTML]::TITRE_Dégradé_Bleu_Droite
        Du fait de l'ajout d'un surtitre, nous ne sommes plus satisfait de l'apparence par défaut de notre tableau.
        On modifie le style des surtitres et des titres.

        >$t.Lignes[0].'Colonne 2'.Style = [StyleHTML]::OK
        On accède à la première ligne (Lignes[0]) du tableau, et dans cette ligne on accède à la cellule de données de la colonne 'Colonne 2'. On modifie le style de la cellule par le style prédéfini OK.
    .EXAMPLE
        $r.GénérerRapport() | Out-File test.html

        On génère le fichier html et on l'enregistre.

#>

Enum StyleHTML {
    NEUTRE
    OK
    INFO
    ALERTE
    ERREUR
    GRAS
    ITALIQUE
    TITRE_Dégradé_Bleu_Descendant
    TITRE_Dégradé_Bleu_Milieu
    TITRE_Dégradé_Bleu_Montant
    TITRE_Dégradé_Bleu_Droite
    TITRE_Dégradé_Bleu_Gauche
    TITRE_Gras
}

Enum Alignement {
    left
    center
    right
}

Enum FormatTableau {
    Liste
    Colonne
}

Class StylePerso {
    # remarque : on ne définit pas d'alignement horizontal car il entrerait en conflit avec celui défini directement dans la cellule.
    [string] $background
    [string] $color
    [string] ${font-family}
    [string] ${fontsize}
    [string] ${text-transform}
    [string] $padding
    [string] ${vertical-align}
    [string] ${font-weight}
}

Class RapportHTML {
    hidden $_HTML_fr = (New-Object system.globalization.cultureinfo(“fr-FR”))
    hidden $_nbrStylePerso = 0
    hidden $_style = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }

    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }

    
    
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }
    


    #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }


    .TITRE_BLANC {
        background: #fff;
        color: #fff;
	}

    .TITRE_DÉGRADÉ_BLEU_MILIEU {
        background: linear-gradient(to right, #49708f, #293f50, #49708f);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
		font-weight: bold;
	}

    .TITRE_DÉGRADÉ_BLEU_DESCENDANT {
        background: linear-gradient(#293f50, #49708f);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
        font-weight: bold;
	}

    .TITRE_DÉGRADÉ_BLEU_MONTANT {
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
        font-weight: bold;
	}

    .TITRE_DÉGRADÉ_BLEU_DROITE {
        background: linear-gradient(to right, #293f50, #49708f);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
        font-weight: bold;
	}

    .TITRE_DÉGRADÉ_BLEU_GAUCHE {
        background: linear-gradient(to right, #49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
        font-weight: bold;
	}

    .TITRE_GRAS {
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
        font-weight: bold;
	}


    .NEUTRE {

        color: #000000;
    }
    
  
    .OK {

        background: #008000;
        color: #FFFFFF;
        font-weight: bold;
    }
    
  
    .INFO {

        background: #FFFF00;
        color: #000000;
    }
    
  
    .ALERTE {

        background: #FF8C00;
        color: #FFFFFF;
        font-weight: bold;
    }
    
  
    .ERREUR {

        background: #FF0000;
        color: #FFFFFF;
        font-weight: bold;
    }
    
    .GRAS {

        font-weight: bold;
    }
    
    .ITALIQUE {

        font-style: italic;
    }
    
</style>
"@

    [string] $Titre
    [Tableau[]] $Tableaux

    # CONSTRUCTEURS
    RapportHTML ([string] $Titre) {
        $this.Titre = $Titre}

    # METHODES
    [Tableau] CréerTableau ([string] $Titre, $Tableau, [FormatTableau] $FormatTableau) {
        $tableau = [Tableau]::new($Titre, $Tableau, $this, $FormatTableau)
        $this.Tableaux += $tableau
        return $tableau
    }

    [string[]] GénérerRapport () {
        [string[]] $rapport = @()
        $rapport += '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">'
        $rapport += '<html xmlns="http://www.w3.org/1999/xhtml">'
        $rapport += '<head>'
        $rapport += $this._style
        $rapport += '</head><body>'
        $rapport += "<h1>" + $this.Titre + "</h1>"
        foreach ($tableau in $this.Tableaux) {$rapport += $this.ConvertirTableau($tableau)}
        $rapport += "<p id='CreationDate'>Génération du rapport: $(Get-Date -Format ($this._HTML_fr.DateTimeFormat.FullDateTimePattern))</p>"
        $rapport += '</body></html>'
        return $rapport}

    hidden [string] ConvertirCellulesFormatColonnes ([pscustomobject] $psCellules, [boolean] $EstSurtitres) {
        $tableau = $psCellules.psobject.properties.value[0].Colonne.Tableau
        $NomsColonnes = @($tableau.NomsColonnes)
        $ligne = "<tr>"
        if ($EstSurtitres) { $ligne += " <!--SURTITRES-->"} # pour lisibilité html uniquement
        if ($EstSurtitres) {$ligne += "`n"} # pour lisibilité html uniquement
        for($i=0; $i -lt $tableau.NombreColonnes; $i++) {
            $nomColonne = $NomsColonnes[$i]
            if ($EstSurtitres) {$ligne += "    "} # pour lisibilité html uniquement
            $ligne += "<td"
            if ($EstSurtitres -and $psCellules.$nomColonne.Texte) {
                $ligne += ' class="' + $tableau.StyleSurtitre + '"'
                $ligne += ' Style="padding-top:5px;padding-bottom:5px"'
                $ligne += ' align="center"'}
            else {
                $ligne += ' class="' + $psCellules.$nomColonne.Style + '"'
                $ligne += ' align="' + $psCellules.$nomColonne.Alignement + '"'}
            
            $ligne += ' colspan="' + $psCellules.$nomColonne.Etendue + '"'
            $ligne += ">"
            $ligne += $psCellules.$nomColonne.Texte
            $ligne += "</td>"
            if ($EstSurtitres) {$ligne += "`n"} # pour lisibilité html uniquement
            $i += $psCellules.$nomColonne.Etendue - 1}
        
        $ligne += "</tr>"
        return $ligne}

    hidden [string[]] ConvertirTableau ([Tableau] $Tableau) {
        if ($Tableau._FormatTableau -eq [FormatTableau]::Colonne -and ! $Tableau.Lignes) { # si tableau vide on en ajoute une pour qu'il soit affiché
            $ligne = $Tableau.CréerLigne()
            $nomPremièreColonne = $Tableau._NomsColonnes[0]
            $ligne.$nomPremièreColonne.Texte = "Tableau vide"
            $ligne.$nomPremièreColonne.style = [styleHTML]::INFO
            $ligne.$nomPremièreColonne.Alignement = [alignement]::center
            $ligne.$nomPremièreColonne.Etendue = $Tableau.NombreColonnes}
        [string[]] $tableauConverti = @()
        $tableauConverti += "<h2>$($Tableau.Titre)</h2>"
        if ($Tableau.Lignes) {
            if ($Tableau._FormatTableau -eq [FormatTableau]::Liste) {
                $tableauConverti += "<table> <!--Format LISTE-->" # pour lisibilité html uniquement"
                $ligne = @()
                foreach ($psCellules in $Tableau.Lignes) {
                    $etendueEnCours = 0
                    foreach ($nomColonne in $Tableau.NomsColonnes) {
                        $td_surtitre =  ""
                        if ($Tableau._surtitres) {
                            if($Tableau._surtitres.$nomColonne.Texte) {
                                $etendueEnCours = $Tableau._surtitres.$nomColonne.Etendue
                                $td_surtitre =  '<td style="padding-top:0px;padding-bottom:0px" class="' + $Tableau.StyleSurtitre + '" align="center" rowspan="' + $Tableau._surtitres.$nomColonne.Etendue + '">' + $Tableau._surtitres.$nomColonne.Texte + '</td>'
                            }
                            elseif ($etendueEnCours -gt 0) {
                                $td_surtitre =  ''
                            }
                            else {
                                $td_surtitre =  '<td class="TITRE_BLANC"></td>'
                            }
                        }
                        $td_titre =  '<td style="padding-top:0px;padding-bottom:0px" class="' + $Tableau.StyleTitre + '" align="' + $Tableau.AlignementTitre + '">' + $nomColonne + '</td>'
                        $td_valeur = '<td style="padding-top:0px;padding-bottom:0px" class="' + $psCellules.$nomColonne.Style + '">' +$psCellules.$nomColonne.Texte +"</td>"
                        $ligne += "<tr>" + $td_surtitre + $td_titre + $td_valeur + "</tr>"
                        $etendueEnCours--}
                    $tableauConverti += $ligne}}
            elseif ($Tableau._FormatTableau -eq [FormatTableau]::Colonne){
                $tableauConverti += "<table> <!--Format COLONNE-->" # pour lisibilité html uniquement"
                $tableauConverti += "<colgroup>$("<col/>"*($Tableau.NomsColonnes.count))</colgroup>"
                if ($Tableau._surtitres) {
                    #$tableauConverti += "<!--SURTITRES-->" # pour lisibilité html uniquement
                    $tableauConverti += $this.ConvertirCellulesFormatColonnes($Tableau._surtitres, $true)}
                $tableauConverti += '<tr><th align="' + $Tableau.AlignementTitre + '">' + ($Tableau.NomsColonnes -join '</th><th align="' + $Tableau.AlignementTitre + '">') + '</th></tr>'
                foreach ($psCellules in $Tableau.Lignes) {
                    $tableauConverti += $this.ConvertirCellulesFormatColonnes($psCellules, $false)}}
            else {
                Throw "Format de Tableau non traité"}
            $tableauConverti += '</table>'}
        return $tableauConverti}

    [string] AjouterStylePerso ([StylePerso] $StylePerso) {
        $this._nbrStylePerso++
        $stylename = "PERSO" + $this._nbrStylePerso
        $stylecontent = "    .$stylename {`n"

        # seules les propriétés renseignées de $styleperso seront ajoutés dans la propriété _style de la classe [rapportHTML])
        foreach ($format in $StylePerso.psobject.properties.name) {
            if ($StylePerso.$format) {$stylecontent += "        $($format): $($StylePerso.$format);`n"}}

        $stylecontent += "    }`n"
        $this._style = $this._style -replace "</style>", "$stylecontent</style>"
        return $stylename
    }
}

Class Tableau {
    [string]  $Titre
    [Alignement]   $AlignementTitre

    hidden [FormatTableau] $_FormatTableau
    hidden [string[]] $_nomsColonnes
    hidden [int] $_nombreColonnes
    hidden [RapportHTML] $_rapport
    hidden [pscustomobject] $_colonnes
    hidden [Ligne[]] $_lignes
    hidden [pscustomobject] $_surtitres
    hidden [string] $_StyleTitre
    hidden [string] $_StyleSurtitre

    # CONSTRUCTEURS
    hidden Tableau ([string] $Titre, $Données, [RapportHTML] $RapportHTML, $FormatTableau) {

        #-------------------------------------------
        # contrôle de validité du paramètre $Données
        #-------------------------------------------
        
        if ($Données -is [Array]) {
            if ($Données.count -eq 0) {Throw "Le tableau transmis est vide."}
            $typename = $Données[0].Gettype().Name
            if (("string","pscustomobject","OrderedDictionary") -notcontains $typename) {Throw "Le tableau passé en paramètre doit contenir un seul type (string, pscustomobject ou ordered"}
            foreach ($item in $Données) {
                if ($item.Gettype().Name -notlike $typename) {Throw "Le tableau ne contient pas uniquement des objets de même type."}
            }
            switch ($typename) {
                String {$this._nomsColonnes = [string[]] $Données; break}
                PSCustomObject {$this._nomsColonnes = [string[]] $Données[0].psobject.Properties.Name; break}
                OrderedDictionary {$this._nomsColonnes = [string[]] $Données[0].Keys; break}
                Default {Throw "Le paramètre 'Données' doit être le ou les noms des colonnes, ou bien de type [pscustomobject] ou [ordered] s'il y a des valeurs."}} 
        }
        else {
            switch ($Données.Gettype().Name) {
                String {$this._nomsColonnes = [string[]] $Données; break}
                PSCustomObject {$this._nomsColonnes = [string[]] $Données.psobject.Properties.Name; break}
                OrderedDictionary {$this._nomsColonnes = [string[]] $Données.Keys; break}
                Default {Throw "Le paramètre 'Données' doit être le ou les noms des colonnes, ou bien  de type [pscustomobject] ou [ordered] s'il y a des valeurs."}} 
        }
        if ($this._nomsColonnes.Count -ne ($this._nomsColonnes | Sort-Object -Unique).Count) {Throw "Doublon dans les noms de colonnes !"}

        #----------------------------------------------
        # Param $Données OK : poursuite du constructeur
        #----------------------------------------------
        
        $this.Titre = $Titre
        $this._FormatTableau = $FormatTableau
        $this | Add-Member ScriptProperty 'NomsColonnes' `            {$this._nomsColonnes}
        $this._nombreColonnes = $this._nomsColonnes.Count
        $this | Add-Member ScriptProperty 'NombreColonnes' `            {$this._nombreColonnes}
        $this | Add-Member -MemberType ScriptProperty -Name StyleTitre -Value `            {$this._StyleTitre} `
            {
                Param ([string] $style)
                $this._StyleTitre = $style.ToUpper()
            }
        $this | Add-Member -MemberType ScriptProperty -Name StyleSurtitre -Value `            {$this._StyleSurtitre} `
            {
                Param ([string] $style)
                $this._StyleSurtitre = $style.ToUpper()
            }
        if ($this._FormatTableau -eq [FormatTableau]::Liste) {
            # format des titres par défaut pour les tableaux au format Liste
            $this.StyleTitre = [StyleHTML]::TITRE_Gras
            $this.AlignementTitre = [Alignement]::left
            $this.StyleSurtitre = [StyleHTML]::TITRE_Gras}
        else {
            # format des titres par défaut pour les tableaux au format Colonne
            $this.StyleTitre = [StyleHTML]::TITRE_Dégradé_Bleu_Montant
            $this.AlignementTitre = [Alignement]::center
            $this.StyleSurtitre = [StyleHTML]::TITRE_Dégradé_Bleu_Descendant}
        $numColonne = 1
        $this._colonnes = [pscustomobject]@{}
        foreach ($nomColonne in $this._nomsColonnes) {
            $this._colonnes | Add-Member $nomColonne ([Colonne]::new($numColonne, $this))
            $numColonne++}
        $this | Add-Member ScriptProperty 'Lignes' `            {$this._lignes}

        # création des lignes si besoin        
        if ($Données -is [Array]) {
            if ($Données[0].Gettype().Name  -notlike "string") { # on exclu le cas où seuls les noms de colonnes ont été passés
                foreach ($ligne in $Données) {
                    $this.CréerLigne($ligne)}}}
        else {
            if ($Données.Gettype().Name -notlike "string") {
                $this.CréerLigne($Données)}}
    } # fin constructeur

    # METHODES
    ModifierAlignementDesColonnes ([Alignement] $Alignement) {
            foreach ($nomColonne in $this._nomsColonnes) {
                    foreach ($cellule in $this._Colonnes.$nomColonne._cellules) {
                        $cellule.Alignement = $Alignement}}}

    ModifierAlignementDesColonnes ([Alignement] $Alignement, [int[]] $NumérosColonnes) {
            foreach ($numéroColonne in $NumérosColonnes) {
                if ($numéroColonne -gt 0 -and $numéroColonne -le ($this.NombreColonnes)) {
                    $this.ModifierAlignementDesColonnes($Alignement, $this._nomsColonnes[$numéroColonne - 1])}
                else {
                    Throw "La colonne n° $NumérosColonnes n'existe pas."}}}

    ModifierAlignementDesColonnes ([Alignement] $Alignement, [string[]] $NomsColonnes) {
            foreach ($nomColonne in $NomsColonnes) {
                if ($this._nomsColonnes -contains $nomColonne){
                    foreach ($cellule in $this._Colonnes.$nomColonne._cellules) {
                        $cellule.Alignement = $Alignement}}
                else {
                    Throw "Tentative d'alignement du texte d'une colonne inexistante ($nomColonne)."}}}

    ModifierAlignementParDéfautDesColonnes ([Alignement] $Alignement) {
            foreach ($nomColonne in $this._nomsColonnes) {
                $this._Colonnes.$nomColonne._AlignementParDéfaut = $Alignement}}

    ModifierAlignementParDéfautDesColonnes ([Alignement] $Alignement, [int[]] $NumérosColonnes) {
            foreach ($numéroColonne in $NumérosColonnes) {
                if ($numéroColonne -gt 0 -and $numéroColonne -le ($this.NombreColonnes)) {
                    $this.ModifierAlignementParDéfautDesColonnes($Alignement, $this._nomsColonnes[$numéroColonne - 1])}
                else {
                    Throw "La colonne n° $NumérosColonnes n'existe pas."}}}

    ModifierAlignementParDéfautDesColonnes ([Alignement] $Alignement, [string[]] $NomsColonnes) {
            foreach ($nomColonne in $NomsColonnes) {
                if ($this._nomsColonnes -contains $nomColonne){
                    $this._Colonnes.$nomColonne._AlignementParDéfaut = $Alignement}
                else {
                    Throw "La colonne '$nomColonne' n'existe pas."}}}

    ModifierStyleDesColonnes ([string] $style) {
            foreach ($nomColonne in $this._nomsColonnes) {
                    foreach ($cellule in $this._Colonnes.$nomColonne._cellules) {
                        $cellule.Style = $Style}}}

    ModifierStyleDesColonnes ([string] $style, [int[]] $NumérosColonnes) {
            foreach ($numéroColonne in $NumérosColonnes) {
                if ($numéroColonne -gt 0 -and $numéroColonne -le ($this.NombreColonnes)) {
                    $this.ModifierStyleDesColonnes($style, $this._nomsColonnes[$numéroColonne - 1])}
                else {
                    Throw "Tentative de formatage d'une colonne inexistante (numéro d'ordre = $NumérosColonnes)."}}}

    ModifierStyleDesColonnes ([string] $style, [string[]] $NomsColonnes) {
            foreach ($nomColonne in $NomsColonnes) {
                if ($this._nomsColonnes -contains $nomColonne){
                    foreach ($cellule in $this._Colonnes.$nomColonne._cellules) {
                        $cellule.Style = $Style}}
                else {
                    Throw "Tentative de formatage d'une colonne inexistante ($nomColonne)."}}}

    ModifierStyleParDéfautDesColonnes ([string] $style) {
            foreach ($nomColonne in $this._nomsColonnes) {
                $this._Colonnes.$nomColonne._StyleParDéfaut = $style}}

    ModifierStyleParDéfautDesColonnes ([string] $style, [int[]] $NumérosColonnes) {
            foreach ($numéroColonne in $NumérosColonnes) {
                if ($numéroColonne -gt 0 -and $numéroColonne -le ($this.NombreColonnes)) {
                    $this.ModifierStyleParDéfautDesColonnes($style, $this._nomsColonnes[$numéroColonne - 1])}
                else {
                    Throw "La colonne n° $NumérosColonnes n'existe pas."}}}

    ModifierStyleParDéfautDesColonnes ([string] $style, [string[]] $NomsColonnes) {
            foreach ($nomColonne in $NomsColonnes) {
                if ($this._nomsColonnes -contains $nomColonne){
                    $this._Colonnes.$nomColonne._StyleParDéfaut = $style}
                else {
                    Throw "La colonne '$nomColonne' n'existe pas."}}}

    hidden TesterAjoutLigne() {
        # empêche de créer plus d'une ligne dans un tableau au format Liste
        if ($this.FormatTableau -like "Liste") {
            if ($this.Lignes.count -eq 1) {Throw "Un tableau au format Liste n'accepte qu'une seule ligne."}}}

    [Ligne] CréerLigne () {
        $this.TesterAjoutLigne()
        return [Ligne]::new($this)}

    [Ligne] CréerLigne ($Données) {
        $this.TesterAjoutLigne()
        return [Ligne]::new($this, $Données)}


    CréerSurtitres ([string]$Titre, [int]$NuméroColonne, [int] $Etendue) {
        if ($NuméroColonne -lt 1 -or $NuméroColonne -gt $this._nombreColonnes) {
            Write-Host "[CréerSurtitre] numéro de colonne incorrect." -ForegroundColor Red
            Throw "[CréerSurtitre] Le numéro de la colonne de départ est en dehors de la plage possible : 1 <= numéro <= $($this._nombreColonnes)."}
        if (($NuméroColonne-1+$Etendue) -gt $this._nombreColonnes) {
            Write-Host "[CréerSurtitre] étendue trop grande." -ForegroundColor Red
            Throw "[CréerSurtitre] L'étendue va au-delà de la plage possible."}
        if ($Etendue -lt 1) {
            Write-Host "[CréerSurtitre] étendue trop petite." -ForegroundColor Red
            Throw "[CréerSurtitre] L'étendue est inférieure à 1."}
        # si un surtitre existe déjà, on reprend l'existant, sinon on part de zéro
        if ($this._surtitres) {
            $ligne = $this._surtitres}
        else {
            $ligne = [pscustomobject]@{}}
        $EtendueEnCours = 1
        foreach ($nomColonne in $this._nomsColonnes) {
            # si le membre n'existe pas déjà on le crée (cas où on part de zéro)
            if (!$ligne.$nomColonne) {
                $ligne | Add-Member $nomColonne ([Cellule]::new("", $this._Colonnes.$nomColonne))}
            if ($this._nomsColonnes[$NuméroColonne - 1] -eq $nomColonne) {
                if ($ligne.$nomColonne.Texte -or $EtendueEnCours -gt 1) {
                    Write-Host "[CréerSurtitre] Chevauchement de surtitre." -ForegroundColor Red
                    Throw "[CréerSurtitre] Vous commencez votre surtitre sur un surtitre existant."}
                $ligne.$nomColonne.Texte = $Titre
                $ligne.$nomColonne.Etendue = $Etendue
                $EtendueEnCours = $Etendue+1}
            else {
                if ($ligne.$nomColonne.Texte -and $EtendueEnCours -gt 1) {
                    Write-Host "[CréerSurtitre] chevauchement de surtitre." -ForegroundColor Red
                    Throw "[CréerSurtitre] le surtitre déborde sur un surtitre existant."}
                else {
                    $EtendueEnCours = $ligne.$nomColonne.Etendue+1}}
            if ($EtendueEnCours -gt 1) {$EtendueEnCours--}}
        $this._surtitres = $ligne}
}

Class Colonne {
    hidden [int] $_numéro
    hidden [Tableau] $_tableau
    hidden [Cellule[]] $_cellules = @()
    hidden [Alignement] $_AlignementParDéfaut = "left"
    hidden [string] $_StyleParDéfaut = [styleHTML]::NEUTRE

    # CONSTRUCTEURS
    hidden Colonne ([int] $Numéro, [Tableau] $Tableau) {
        $this._numéro = $Numéro        $this | Add-Member ScriptProperty 'Numéro' `            {$this._numéro}
        $this._tableau = $Tableau        $this | Add-Member ScriptProperty 'Tableau' `            {$this._tableau}
        $this | Add-Member ScriptProperty 'Cellules' `            {if ($this._cellules) {$this._cellules} else {"Aucune Cellule"}} `            {
                Param ([Cellule] $Cellule)
                if ($this.Equals($Cellule.Colonne) -and $this._cellules -notcontains $Cellule) {
                    $this._cellules += $Cellule}
                else {Throw "Cet objet n'a pas été ajouté aux cellules de la colonne $($this._tableau.NomsColonnes[$this._numéro])."}
            }}
}

Class Cellule {
    [string] $Texte
    [Alignement] $Alignement

    hidden [Colonne] $_colonne
    hidden [int] $_etendue = 1
    hidden [string] $_Style = ""

    # CONSTRUCTEURS
    Cellule ([string] $Texte, [Colonne] $Colonne) {
        $this.Texte = $Texte
        $this._colonne = $Colonne
        $this | Add-Member ScriptProperty 'Colonne' `            {$this._colonne}
        $this | Add-Member ScriptProperty 'Etendue' `            {$this._etendue} `            {                Param ($NombreCellules)                $this.Etendre($NombreCellules)            }
        $this | Add-Member ScriptProperty 'Style' `            {$this._Style} `            {                Param ([string] $Style)                $this._Style = $Style.ToUpper()            }
        $this._colonne._cellules += $this
        $this.Alignement = $this._colonne._AlignementParDéfaut
        $this.Style = $this._colonne._StyleParDéfaut}

    # METHODES
    hidden Etendre () {
        $numéroColonne = $this.Colonne.Numéro
        $totalColonnes = $this.Colonne.Tableau.NombreColonnes
        $valMax = $totalColonnes - $numéroColonne + 1
        $NombreCellules = $valMax
        $this._etendue = $NombreCellules}
    
    hidden Etendre ($NombreCellules) {
        $numéroColonne = $this.Colonne.Numéro
        $totalColonnes = $this.Colonne.Tableau.NombreColonnes
        $valMax = $totalColonnes - $numéroColonne + 1
        if ($NombreCellules -eq $null) {$NombreCellules=$valMax}
        if ($NombreCellules -is [int] -and $NombreCellules -gt 0 -and $NombreCellules -le $valMax) {
            $this._etendue = $NombreCellules}
        else {
            Throw "La propriété Etendue doit être comprise entre 1 (pas d'extension) et $valMax (toutes les cellules restantes)."}    }
}

Class Ligne {
    hidden [pscustomobject] $_Cellules
    Ligne ([tableau]$tableau) {
        foreach ($nomColonne in $tableau._nomsColonnes) {
            $this | Add-Member $nomColonne ([Cellule]::new("", $tableau._Colonnes.$nomColonne))}
        $tableau._lignes += $this}

    Ligne ([Tableau]$tableau, $Données) {
        [string[]] $nomsColonnes = @()
        switch($Données.GetType().Name) {
            "Hashtable" {$nomsColonnes = $Données.Keys;break}
            "OrderedDictionary" {$nomsColonnes = $Données.Keys;break}
            "PSCustomObject" {$nomsColonnes = $Données.psobject.Properties.Name;break}
            default {Throw "Le constructeur de [Ligne] n'accepte que les types suivants : [Hashtable], [OrderedDictionary], [PSCustomObject]"}}

        foreach ($nomColonne in $nomsColonnes) {
            # on vérifie que les noms de colonnes transmis existent bien
            if ($tableau._nomsColonnes -notcontains $nomColonne) {
                Write-Host "Données non conformes pour créer une ligne" -ForegroundColor Red
                Throw "Données non conformes pour créer une ligne"}}

        $this._Cellules = @{}
        foreach ($nomColonne in $tableau._nomsColonnes) {
            # si cette colonne fait partie des données transmises on crée une cellule avec le champ Texte renseigné
            if ($nomsColonnes -contains $nomColonne) {
                $this._Cellules | Add-Member $nomColonne ([Cellule]::new($Données.$nomColonne, $tableau._Colonnes.$nomColonne))}
            # sinon on crée une cellule avec le champ Texte vide
            else {
                $this._Cellules | Add-Member $nomColonne ([Cellule]::new("", $tableau._Colonnes.$nomColonne))}
            # on crée la propriété publique renvoyant la Cellule, avec uniquement un accesseur get
            $this | Add-Member -MemberType ScriptProperty -Name $nomColonne -Value `                {$this._Cellules.$nomColonne}.GetNewClosure()}
                <# on peut aussi créer la propriété publique de la façon ci-dessous. C'est moins beau, mais c'est peut-être un peu moins gourmand en mémoire. J'ignore quelle est la meilleure solution.
                    $command = "
                    `$this | Add-Member -MemberType ScriptProperty -Name '$nomColonne' -Value ``                        {`$this._Cellules.'$nomColonne'}"
                    Invoke-Expression $command
                #>
        $tableau._lignes += $this}
}

<#

SEQUENCE DE TEST MINIMALE

Import-Module .\rapportHTML.ps1

$r = [rapporthtml]::new("Nouveau rapport")
$r.CréerTableau("Tableau vide, une seule colonne", "Une colonne", "colonne") | out-null
$r.CréerTableau("Une seule ligne [ordered], une seule colonne", ([ordered]@{"une colonne"="une valeur"}), "colonne") | out-null
$r.CréerTableau("Tableau vide, plrs colonnes", ("colonne 1","colonne 2"), "colonne") | out-null

$d=[pscustomobject[]]@()
$i=(1..4); $i | % { $p=[pscustomobject]@{}; $j=$_; $i | % {$p | add-member "colonne$_" "val$j"};$d+=$p}
$r.CréerTableau("Tableau de psobject", $d, "colonne") | out-null
$t = $r.CréerTableau("Tableau de psobject avec surtitre", $d, "colonne")
$t.CréerSurtitres("surtitre",2,2)

$r.CréerTableau("Un seul psobject au format Liste", $d[0], "liste") | out-null
	
$t = $r.CréerTableau("tableau vide puis ligne créée", ("colonne 1","colonne 2"), "colonne")
$t.CréerLigne(@{"colonne 1"="val1";"colonne 2"="val2"}) | out-null

$r.CréerTableau("tableau Liste", [ordered]@{"colonne 1"="val1";"colonne 2"="val2";"colonne 3"="val3";"colonne 4"="val4"}, "liste") | out-null

$t = $r.CréerTableau("tableau Liste avec surtitre", [ordered]@{"colonne 1"="val1";"colonne 2"="val2";"colonne 3"="val3";"colonne 4"="val4"}, "liste")
$t.CréerSurtitres("surtitre",2,2)

$r.GénérerRapport() | Out-File test.html


#>
# SIG # Begin signature block
# MIIFowYJKoZIhvcNAQcCoIIFlDCCBZACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU094bn1S0rVFu4VgZqdcrMPpM
# 4XmgggMsMIIDKDCCAhCgAwIBAgIQFpiwcPSA9plJJVYjwMrpxzANBgkqhkiG9w0B
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
# SKDQFYpaE/7yvinFD45qx5BAqVEwDQYJKoZIhvcNAQEBBQAEggEAPP2DCtJojV7W
# nDNQS1RM8En8OfInFiM4/5tSXjYkhlVg+yg3FhfR7xh7+VKTPIMfWoX7LMH6W3Un
# BvbYeMEpjuR0BGixjYZU3MQiOFGLONZOpAYWfdkajLd1CglpfuUhk5dQhk9Nz4kZ
# Mvvd0+LDD6Y4+HVUJ3N+QoZqyWN3r9GUsaxXvkMpc++7wjeAOed3KzgaGqFi/OT3
# n+7y9PV6ue10HkRFvPyHGsBWKaza5x5b5FQCooLbqyuWAlc2SGepRQmzER/rOspz
# 10b1GwdSi8SvbiAHPTsfiSIwIzGxLwTS/ajrXUrGSL93oBEMvrgAmYO3O9VG6R3f
# pIKPATLuVg==
# SIG # End signature block
