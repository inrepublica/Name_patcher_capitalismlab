#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ico\write.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#Include <Array.au3>
#include<File.au3>
#include<functions\csv2array.au3>

;Renseignement sur le programme
$boite = MsgBox(1, "Name patcher write", "Place CAPMOD.RES, corporation.csv and person.csv in the same folder of this tool and click OK.")
If $boite = 2 Then Exit

;Création de la fenêtre de progression
ProgressOn("Name patcher write", "wait few seconds...")
ProgressSet(0, "Open CAPMOD.RES")

;Ouverture du fichier et mise dans un tableau
$fichier_a_ouvrir = "CAPMOD.RES"
$fichier = FileOpen($fichier_a_ouvrir, 16)
$tableau_fichier = StringRegExp(FileRead($fichier), '[[:xdigit:]]{2}', 3)
FileClose($fichier)
ProgressSet(50, 0 & "Processing corporation")

;Définition des variables pour corporation.csv
$debut_nom = 14307
$longueur_fullname = 39
$longueur_shortname = 16
$longueur_logofile = 7
$nombre_entreprise = 94

;Ouverture et chargement du fichier corporation.csv dans un tableau
$tableau_corporationcsv = _CSV2Array("corporation.csv", ";")

;Mise en place du pointeur
$pointeur = $debut_nom

;Boucle sur le nombre d'entreprises
For $x = 1 to $nombre_entreprise
	$corporation_name = $tableau_corporationcsv[$x][0]
	$corporation_shortname = $tableau_corporationcsv[$x][1]
	$corporation_logofile = $tableau_corporationcsv[$x][2]

	;Découpage des noms en tableau de caractères
	$t_corporation_name = StringSplit($corporation_name, "")
	$t_corporation_shortname = StringSplit($corporation_shortname, "")
	$t_corporation_logofile = StringSplit($corporation_logofile, "")

	;Modification du FULL NAME
	For $i = 1 to $t_corporation_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_corporation_name[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	For $i = 1 to $longueur_fullname - $t_corporation_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary(" "), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 1

	;Modification du SHORT NAME
	For $i = 1 to $t_corporation_shortname[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_corporation_shortname[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	For $i = 1 to $longueur_shortname - $t_corporation_shortname[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary(" "), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 2

	;Modification du LOGOFILE
	For $i = 1 to $t_corporation_logofile[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_corporation_logofile[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 23
Next

ProgressSet(70, 0 & "Processing person")

;Définition des variables pour person.csv
$debut_nom = 25155
$longueur_name = 29
$longueur_code = 3
$nombre_personne = 64

;Ouverture et chargement du fichier person.csv dans un tableau
$tableau_personcsv = _CSV2Array("person.csv",";")

;Mise en place du pointeur
$pointeur = $debut_nom

;Boucle sur le nombre de personnes
For $x = 1 to $nombre_personne

	;Découpage en tableau de caractères
	$t_person_name = StringSplit($tableau_personcsv[$x][0], "")
	$t_person_code = StringSplit($tableau_personcsv[$x][1], "")
	$t_person_training = StringSplit(prepare_expertise($tableau_personcsv[$x][2]), "")
	$t_person_advertise = StringSplit(prepare_expertise($tableau_personcsv[$x][3]), "")
	$t_person_rd = StringSplit(prepare_expertise($tableau_personcsv[$x][4]), "")
	$t_person_farming = StringSplit(prepare_expertise($tableau_personcsv[$x][5]), "")
	$t_person_manufacturing = StringSplit(prepare_expertise($tableau_personcsv[$x][6]), "")
	$t_person_retail = StringSplit(prepare_expertise($tableau_personcsv[$x][7]), "")
	$t_person_raw = StringSplit(prepare_expertise($tableau_personcsv[$x][8]), "")

	;Modification du NAME
	For $i = 1 to $t_person_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_name[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	For $i = 1 to $longueur_name - $t_person_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary(" "), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 1
	;Modification du CODE
	For $i = 1 to $t_person_code[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_code[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 73
	;Modification du TRAINING
	For $i = 1 to $t_person_training[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_training[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	;Modification du ADVERTISE
	For $i = 1 to $t_person_advertise[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_advertise[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	;Modification du R&D
	For $i = 1 to $t_person_rd[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_rd[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	;Modification du FARMING
	For $i = 1 to $t_person_farming[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_farming[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	;Modification du MANUFACTURING
	For $i = 1 to $t_person_manufacturing[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_manufacturing[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	;Modification du RETAIL
	For $i = 1 to $t_person_retail[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_retail[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	;Modification du RAW
	For $i = 1 to $t_person_raw[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_raw[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 1
Next

ProgressSet(90, "Writing CAPMOD.RES")

;Ecriture du fichier CAPMOD.RES
$fichier_a_ecrire = ""
For $i = 0 to UBound($tableau_fichier)-1
  $fichier_a_ecrire &= $tableau_fichier[$i]
Next
$fichier = FileOpen($fichier_a_ouvrir, 18)
FileWrite($fichier, "0x" & $fichier_a_ecrire)

;Fin de la fenêtre de progression
ProgressSet(100, "Done", "Complete")
Sleep(500)
ProgressOff()

;Message de sortie
MsgBox(0, "Write Done", "File CAPMOD.RES has been patched with yours csv files.")

;Fonctions
Func prepare_expertise ($expertise)
    If (StringLen($expertise) = 1) Then
		$expertise = "  "&$expertise
	ElseIf (StringLen($expertise) = 2) Then
		$expertise = " "&$expertise
	Else
		$expertise = $expertise
	EndIf
	Return $expertise
EndFunc