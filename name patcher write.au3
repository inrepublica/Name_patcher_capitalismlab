#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ico\write.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Include <Array.au3>
#include<File.au3>
#include<functions\csv2array.au3>

;Ouverture du fichier et mise dans un tableau
$fichier_a_ouvrir = "1STD.SET"
$fichier = FileOpen($fichier_a_ouvrir, 16)
$tableau_fichier = StringRegExp(FileRead($fichier), '[[:xdigit:]]{2}', 3)
FileClose($fichier)

;Définition des variables pour corporation.csv
$debut_nom = 17712
$longueur_fullname = 40
$longueur_shortname = 16
$nombre_entreprise = 35

;Ouverture et chargement du fichier corporation.csv dans un tableau
$tableau_corporationcsv = _CSV2Array("corporation.csv", ";")

;Mise en place du pointeur
$pointeur = $debut_nom

;Boucle sur le nombre d'entreprises
For $x = 1 to $nombre_entreprise
	$corporation_name = $tableau_corporationcsv[$x][0]
	$corporation_shortname = $tableau_corporationcsv[$x][1]
	$corporation_color = $tableau_corporationcsv[$x][2]
	;Test si corporation color est vide
	If ($corporation_color = "") Then
		$corporation_color = " "
	EndIf

	;Découpage des noms en tableau de caractères
	$t_corporation_name = StringSplit($corporation_name, "")
	$t_corporation_shortname = StringSplit($corporation_shortname, "")

	;Modification du FULL NAME
	For $i = 1 to $t_corporation_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_corporation_name[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	For $i = 1 to $longueur_fullname - $t_corporation_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary(" "), 2)
		$pointeur = $pointeur + 1
	Next

	;Modification du SHORT NAME
	For $i = 1 to $t_corporation_shortname[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_corporation_shortname[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	For $i = 1 to $longueur_shortname - $t_corporation_shortname[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary(" "), 2)
		$pointeur = $pointeur + 1
	Next

	;Modification du COLOR ID
	$tableau_fichier[$pointeur] = StringTrimLeft(Binary($corporation_color), 2)
	$pointeur = $pointeur + 11
Next

;Définition des variables pour person.csv
$debut_nom = 54387
$longueur_name = 30
$longueur_sex = 8
$longueur_code = 3
$nombre_personne = 64

;Ouverture et chargement du fichier corporation.csv dans un tableau
$tableau_personcsv = _CSV2Array("person.csv",";")

;Mise en place du pointeur
$pointeur = $debut_nom

;Boucle sur le nombre de personnes
For $x = 1 to $nombre_personne

	;Découpage en tableau de caractères
	$t_person_name = StringSplit($tableau_personcsv[$x][0], "")
	$t_person_sex = StringSplit($tableau_personcsv[$x][1], "")
	$t_person_code = StringSplit($tableau_personcsv[$x][2], "")
	$t_person_training = StringSplit(prepare_expertise($tableau_personcsv[$x][3]), "")
	$t_person_advertise = StringSplit(prepare_expertise($tableau_personcsv[$x][4]), "")
	$t_person_rd = StringSplit(prepare_expertise($tableau_personcsv[$x][5]), "")
	$t_person_farming = StringSplit(prepare_expertise($tableau_personcsv[$x][6]), "")
	$t_person_manufacturing = StringSplit(prepare_expertise($tableau_personcsv[$x][7]), "")
	$t_person_retail = StringSplit(prepare_expertise($tableau_personcsv[$x][8]), "")
	$t_person_raw = StringSplit(prepare_expertise($tableau_personcsv[$x][9]), "")

	;Modification du NAME
	For $i = 1 to $t_person_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_name[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	For $i = 1 to $longueur_name - $t_person_name[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary(" "), 2)
		$pointeur = $pointeur + 1
	Next
	;Modification du SEX
	For $i = 1 to $t_person_sex[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_sex[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 3
	;Modification du CODE
	For $i = 1 to $t_person_code[0]
		$tableau_fichier[$pointeur] = StringTrimLeft(Binary($t_person_code[$i]), 2)
		$pointeur = $pointeur + 1
	Next
	$pointeur = $pointeur + 51
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

;Ecriture du fichier 1STD.SET
$fichier_a_ecrire = ""
For $i = 0 to UBound($tableau_fichier)-1
  $fichier_a_ecrire &= $tableau_fichier[$i]
Next
$fichier = FileOpen($fichier_a_ouvrir, 18)
FileWrite($fichier, "0x" & $fichier_a_ecrire)

;Message de sortie
MsgBox(0, "Write Done", "File 1STD.SET has been patched with yours csv files.")

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