cls
#se script est un script signer, seule ce type de script peut etre utiliser sur ce domaine, pour savoir comment signer un script referer vous au cahier de consigne 
#crée a 
#date : 
#par : 
#but : créé rapidement des utilisateurs, ordinateurs, groupe, dans l'active directory à partir d'un fichier csv, ainsi que crée des boites au lettres pour ces utilisateurs
#la parti sur les boite au lettres est à éxécuter sur le serveur Exchange

#déclaration d'un compteur pour compter le nombre d'utilisateurs/DL/ordinateurs/boite au lettre crées
$dndomain = (Get-ADDomain).distinguishedname
$dnsroot = (Get-ADDomain).dnsroot
$compteur_user = 0
$compteur_ordi = 0
$compteur_DL = 0
$compteur_mail = 0

# /!\ specifié les chemins des fichiers csv dans les variable en dessous 
$chemin_fichier_user = "\\admin.lab\script\crea_user.csv"
$chemin_fichier_DL = "\\admin.lab\script\\DL.csv"
$chemin_fichier_ordinateur = "\\admin.lab\script\ordinateur.csv"


# on encode le fichier csv en utf8 pour ne pas avoir de probleme avec les accents...
$UTF8 = @($chemin_fichier_user,$chemin_fichier_DL,$chemin_fichier_ordinateur)
foreach ($fichier in $UTF8)
{
$content = Get-Content -Path $fichier
Set-Content -path $fichier -Value $content -Encoding UTF8
}

#import des fichiers csv necessaire pour le script
$importuser = import-csv -Path "$chemin_fichier_user" -Delimiter ";"
$importDL = import-Csv -Path "$chemin_fichier_DL" -delimiter ";"
$import_ordi = import-Csv -Path "$chemin_fichier_ordinateur" -Delimiter ";"
$import_mail = import-Csv -Path "$chemin_fichier_user" -Delimiter ";"


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
              $secure_pwd = ConvertTo-SecureString -AsPlainText "$password" -Force
              $groupe = $item.groupe
              $alias = "$prenom"+"."+"$nom" 
              $email = "$alias"+"@"+"$dnsroot" -replace ' ','-' -replace "é","e" -replace "ù","u" -replace "à","a" -replace "ë","e" -replace "ô","o" -replace "ç","c"
              $chef = $item.chef

              New-ADUser -Name $nom `
                -Surname $nom `
                -GivenName $prenom `
                -DisplayName $display `
                -AccountPassword $secure_pwd `
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
            $cn_ordi = $item_ordi.cn
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
            Add-ADGroupMember -Identity "$cn_ordi"+"",$dndomain" -Members "$dn_ordi"
            $compteur_ordi ++

             }
             Write-Host "vous avez crée $compteur_ordi ordinateurs"
             }

        "4"
        {
        # permet de pouvoir utiliser les commandes exchange sur PowerShell
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
            foreach ( $i in $import_mail) 
            { 
            $user = $i.login
            $nom_domain = $i.domain
            $login_mail = "$user@$nom_domain" 
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
        cls
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
cls
}

while ($choix -ne '5')

write-host
write-host " #####################"
write-host "#                     #"
Write-host "# ! Fin du script !   #"
write-host "#                     #"
write-host " #####################"
write-host
#singature du script retirer

