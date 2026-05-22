Clear-Host
#script signez, seule se type de script pour etre utiliser sur ce domaine (lab)
#crée a ville
#date :  jj/mm/yyyy
#par : admin-name
#but : créé rapidement des utilisateurs, ordinateurs, groupe, dans l'active directory à partir d'un fichier csv, ainsi que crée des boites au lettres exchange pour ces utilisateurs
#la parti sur les boite au lettres est à éxécuter sur le serveur Exchange ou sur exchange magement shell

#variable recupérant distinguishedname et le nom de domaine 
#variable pour compter le nombre d'utilisateurs/DL/ordinateurs/boite au lettre crées
$dndomain = (Get-ADDomain).distinguishedname
$dnsroot = (Get-ADDomain).dnsroot
$compteur_user = 0
$compteur_ordi = 0
$compteur_DL = 0
$compteur_mail = 0

# /!\ specifié les chemins des fichiers csv dans les variable en dessous 
$Dossiercsv = "D:\admin\Scripts\csv\"
$chemin_fichier_user = Join-Path -Path $Dossiercsv -ChildPath "utilisateurs.csv"
$chemin_fichier_DL = Join-Path -Path $Dossiercsv -ChildPath "DL.csv"
$chemin_fichier_ordinateur = Join-Path -Path $Dossiercsv -ChildPath "ordinateur.csv"


# Initialisation du compteur de fichiers valides
$CompteurFichiersCsv = 0
#Test de présence de chaque fichier

$test_User = Test-Path -Path $chemin_fichier_user
if ($test_User) { $CompteurFichiersCsv++ }

$test_DL = Test-Path -Path $chemin_fichier_DL
if ($test_DL) { $CompteurFichiersCsv++ }

$Test_Ordi = Test-Path -Path $chemin_fichier_ordinateur
if ($Test_Ordi) { $CompteurFichiersCsv++ }


# Vérification: si aucun fichier n'est présent (compteur à 0), fin du script.
if ($CompteurFichiersTrouves -eq 0) {
    Write-Error "aucun fichier csv trouver dans le dossier $Dossiercsv, verifier le nom et l'emplacement des fichiers. arret du script"
    Exit
}

# Traitement et importation selon les resultats des tests

#  Section UTILISATEURS 
if ($Test_User) {
    Write-Host "encodage et import du fichier utilisateurs.csv..." -ForegroundColor Green
    $content = Get-Content -Path $chemin_fichier_user -Raw
    Set-Content -Path $chemin_fichier_user -Value $content -Encoding UTF8
    $importuser = Import-Csv -Path $chemin_fichier_user -Delimiter ";"
    $import_mail = $importuser
} else {
    Write-Warning "Le fichier utilisateurs.csv est absent."
    $importuser = $null
    $import_mail = $null
}

#  Section DL 
if ($Test_DL) {
    Write-Host "encodage et import du fichier DL.csv..." -ForegroundColor Green
    $content = Get-Content -Path $chemin_fichier_DL -Raw
    Set-Content -Path $chemin_fichier_DL -Value $content -Encoding UTF8
    $importDL = Import-Csv -Path $chemin_fichier_DL -Delimiter ";"
} else {
    Write-Warning "Le fichier DL.csv est absent."
    $importDL = $null
}

#  Section ORDINATEURS 
if ($Test_Ordi) {
    Write-Host "encodage et import du fichier ordinateur.csv..." -ForegroundColor Green
    $content = Get-Content -Path $chemin_fichier_ordinateur -Raw
    Set-Content -Path $chemin_fichier_ordinateur -Value $content -Encoding UTF8
    $import_ordi = Import-Csv -Path $chemin_fichier_ordinateur -Delimiter ";"
} else {
    Write-Warning "Le fichier ordinateur.csv est absent."
    $import_ordi = $null
}


#mise en place d'un menu 
write-host ""
do
{

Write-Host "===================================================================================="
Write-Host "========================= Menu de gestion des utilisateurs ========================="
Write-Host "===================================================================================="
Write-Host "              1: Saisissez '1' création d'utilisateurs"
Write-Host "              2: Saisissez '2' création de DLs"
Write-Host "              3: Saisissez '3' création d'ordinateurs"
Write-Host "              4: Saisissez '4' création de boites mails"
Write-Host "              5: Saisissez '5' pour quitter"
Write-Host "===================================================================================="

Write-Host ""
$choix = Read-Host " Veuillez sélectionner un choix de menu (1 .. 5) "
Write-Host ""
    
    Switch ($choix)
    {
        "1"
        {
            #parti sur la création des utilisateurs, $item nom donner arbitrairement
            foreach ($item in $importuser)
            {
            $ou = $item.ou
            #on recupere le DN de l'OU dans laquelle on veut crée notre utilisateurs (/!\ renseigner dans la colone OU dans le fichier csv)
            $user_chemin = (Get-ADOrganizationalUnit -Filter { Name -eq $ou } -SearchScope Subtree).distinguishedName

            # test si l'OU "users" existe dans l'OU qu'on à rechercher juste avant (si elle existe on ne fait rien, si elle n'existe pas on la crée
            $test = Get-ADObject -Filter * |?{$_.name -like 'users'}|?{$_.DistinguishedName -like "ou=users,$user_chemin" }

                if ($test -like "ou*")
                {
  
                }
                else
                {
                 New-ADOrganizationalUnit -Name users -path $user_chemin
                }

              #declaration de toute nos variables, qui prendrons les valeurs se trouvant dans le fichier csv
              $nom = $item.nom
              $prenom = $item.prenom
              $grade = $item.grade
              $display = "$grade $nom $prenom"
              $service = $item.service
              $entite = $item.Entite
              $fonction = $item.fonction
              $tph = $item.tph_bureau
              $login = $item.login  
              $path = "ou=users,$user_chemin"
              $password = $item.password
              $expiration = $item.expiration
              $pwd_user = ConvertTo-SecureString -AsPlainText "$password" -Force
              $groupe = $item.groupe
              $alias = "$prenom"+"."+"$nom" 
              $email = "$alias"+"@"+"$dnsroot" -replace ' ','-' -replace "é","e" -replace "ù","u" -replace "à","a" -replace "ë","e" -replace "ô","o" -replace "ç","c"
              $chef = $item.chef

              New-ADUser -Name $nom `
                -Surname $nom `
                -GivenName $prenom `
                -DisplayName $display `
                -AccountPassword $pwd_user `
                -SamAccountName $login `
                -UserPrincipalName $login@$dnsroot `
                -Path $path `
                -Enabled $true `
                -Title "$fonction"`
                -MobilePhone $tph `
                -AccountExpirationDate "$expiration"`
                -Company "$entite" `
                -Department "$service"`
                -Initials "$grade" `
                -EmailAddress "$email" 
  
              #mise en place d'un compteur pour verifier si tous les utilisateurs du csv on bien été créés
              $compteur_user ++

              #condition si la colone chef dans le csv est a 1 ou 0, place les utilisateur "chef" dans le groupe chef plutot que le groupe utilisateurs
              if ($chef -eq "1")
              {
              $groupechef = "$groupe"+"_chef"
              $dn_groupe = (Get-ADGroup -Filter { Name -like $groupechef }-SearchScope Subtree).distinguishedName
              Add-ADGroupMember -Identity $dn_groupe -Members $login 
              }
              else
                {
                $dn_groupe = (Get-ADGroup -Filter { Name -like $groupe }-SearchScope Subtree).distinguishedName
                Add-ADGroupMember -Identity $dn_groupe -Members $login 
                }
         }
         #affichage du nombre d'utilisateurs créés par le script
         Write-Host " vous avez crée $compteur_user compte utilisateurs"
          }

        "2"
         {
         #parti pour la création de DL 

         #test de l'existence de l'OU DL et création si necessaire 

         $test_dl = Get-ADObject -Filter * |?{$_.name -like 'DL'}|?{$_.DistinguishedName -like "ou=DL,$dndomain" }
         if ($test_dl -like "ou*")
         {
  
         }

         else
            {
            New-ADOrganizationalUnit -name DL -path $dndomain 
            }
            #declaration du chemin de l'OU DL
            $dlpath = "ou=dl,$dndomain"
            foreach ($item_DL in $importDL)       #$item nom donner arbitrairement
            {
            #on recupere la valeur de dl dans le fichier csv, puis 2 variable pour avec un dl en R l'autre en Rw
            $dl = $item_DL.dl
            $dl_R = "DL_"+$dl+"_R"
            $dl_RW = "DL_"+$dl+"_RW"
            #creation d'un tableau avec les 2 nom en R et RW
            $tableau = @($dl_r,$dl_rw)
            foreach ($item_tab in $tableau)
            {
            #creation des dl pour chaque valeur du tableau 
            New-ADGroup -Name $item_tab -GroupCategory Security -GroupScope DomainLocal -Path $dlpath
            
            #on incremente le compteur pour savoir combien de dl on été créés
            $compteur_DL ++
            }
            Write-Host " vous avez crée $compteur_DL DL"
            }
         }
             
        "3"
         {
         #parti sur les ordinateurs
            foreach ($item_ordi in $import_ordi)
            {
            #declaration des variable pour la création des compte ordinateur
            $groupe_ordi = $item_ordi.groupe
            $ordi = $item_ordi.ou
            $ordi_chemin = (Get-ADOrganizationalUnit -Filter { Name -eq $ordi } -SearchScope Subtree).distinguishedName
            $test_ordi = Get-ADObject -Filter * |?{$_.name -like 'ordi*'}|?{$_.DistinguishedName -like "ou=ordinateurs,$ordi_chemin" }
            
            if ($test_ordi -like "ou*")
            {
   
            }
            else
            {
             New-ADOrganizationalUnit -Name ordinateurs -path $ordi_chemin
            }
            $nomordi = $item_ordi.ordi
            $pathordi = "ou=ordinateurs,$ordi_chemin"

            $dn_ordi = (Get-ADGroup -Filter { Name -like $groupe_ordi }-SearchScope Subtree).distinguishedName
 
            if ($dn_ordi -like "CN*")
            {

             }
            else
            {
            New-ADGroup -Name "GG_$groupe_ordi" -Path $ordi_chemin -GroupScope Global -GroupCategory Security 
            }

            New-ADComputer -Name $nomordi -path $pathordi
            Add-ADGroupMember -Identity $dn_ordi -Members "cn=$nomordi,$pathordi"
            Add-ADGroupMember -Identity "CN=GG_ordinateurs,OU=ordinateurs,OU=Administration,$dndomain" -Members "$dn_ordi"
            $compteur_ordi ++

             }
             Write-Host "vous avez crée $compteur_ordi ordinateurs"
             }

        "4"
        {
        # permet de pouvoir utiliser les commandes exchange sur ise
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
            foreach ( $i in $import_mail) 
            { 
            $user = $i.login
            $domain = $i.domain
            $login_mail = "$user@$domain" 
            $alias_mail = "$($i.prenom).$($i.nom)"-replace ' ','-' -replace "é","e" -replace "ù","u" -replace "à","a" -replace "ë","e" -replace "ô","o" -replace "ç","c" 
            $bdd = "BDD_MailBox_$($i.BaL)s"
            $grade_mail = $i.grade
            $fonction_mail = $i.Fonction
            Enable-Mailbox -Identity $login_mail -Alias $alias_mail -Database $bdd
            #ajout des attribut au utilisateurs crée , attribut 1 : grade , atribut 2: leurs fonctions
            Set-Mailbox -identity $login_mail -customAttribute1 "$grade_mail" -customAttribute2 "$fonction_mail"

 
        #On compte le nombre d'utilisateurs créés 
        $compteur_mail ++
        } 
        # On affiche le nombre d'utilisateurs créés
        Write-host "vous avez créé $compteur_mail boite mails "

        }

        "5"
        {
        #choix 5 pour quitter le script
        Clear-Host
        Write-Host
        Write-Host  "Vous avez choisi de quitter le menu ..."
        Write-Host
        Write-Host
        pause
        exit
        }
        #default si une valeur non attendu est entré 
        Default
        {
        Write-host " Vous n'avez pas sélectionné une des valeurs attendus..."
        }
    }

Write-Host ""
Write-Host ""
pause
Clear-Host
}

while ($choix -ne '5')

write-host
write-host " #####################"
write-host "#                     #"
Write-host "# ! Fin du script !   #"
write-host "#                     #"
write-host " #####################"
write-host
#ajoutez signature du script juste après 
