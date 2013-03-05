#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ico\write.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Include <Array.au3>
#include<File.au3>
#include<functions\csv2array.au3>

;Définition des variables pour corporation.csv
$debut_nom = 17712
$longueur_fullname = 40
$longueur_shortname = 16
$nombre_entreprise = 35

;Ouverture du fichier et mise dans un tableau
$fichier_a_ouvrir = "1STD.SET"
$fichier = FileOpen($fichier_a_ouvrir, 16)
$tableau_fichier = StringRegExp(FileRead($fichier), '[[:xdigit:]]{2}', 3)
FileClose($fichier)

;Ouverture et chargement du fichier corporation.csv dans un tableau
$tableau_corporationcsv = _CSV2Array("corporation.csv", ";")

;Mise en place du pointeur
$pointeur = $debut_nom

;Boucle sur le nombre d'entreprises
For $x = 1 to $nombre_entreprise
	$corporation_name = $tableau_corporationcsv[$x][0]
	$corporation_shortname = $tableau_corporationcsv[$x][1]
	$corporation_color = $tableau_corporationcsv[$x][2]

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
	$person_name = $tableau_personcsv[$x][0]

	$person_sex = $tableau_personcsv[$x][1]

	$person_code = $tableau_personcsv[$x][2]

	$person_training = $tableau_personcsv[$x][3]
	If (StringLen($person_training) = 1) Then
		$person_training = "  "&$person_training
	ElseIf (StringLen($person_training) = 2) Then
		$person_training = " "&$person_training
	Else
		$person_training = $person_training
	EndIf

	$person_advertise = $tableau_personcsv[$x][4]
	If (StringLen($person_advertise) = 1) Then
		$person_advertise = "  "&$person_advertise
	ElseIf (StringLen($person_advertise) = 2) Then
		$person_advertise = " "&$person_advertise
	Else
		$person_advertise = $person_advertise
	EndIf

	$person_rd = $tableau_personcsv[$x][5]
	If (StringLen($person_rd) = 1) Then
		$person_rd = "  "&$person_rd
	ElseIf (StringLen($person_rd) = 2) Then
		$person_rd = " "&$person_rd
	Else
		$person_rd = $person_rd
	EndIf

	$person_farming = $tableau_personcsv[$x][6]
	If (StringLen($person_farming) = 1) Then
		$person_farming = "  "&$person_farming
	ElseIf (StringLen($person_farming) = 2) Then
		$person_farming = " "&$person_farming
	Else
		$person_farming = $person_farming
	EndIf

	$person_manufacturing = $tableau_personcsv[$x][7]
	If (StringLen($person_manufacturing) = 1) Then
		$person_manufacturing = "  "&$person_manufacturing
	ElseIf (StringLen($person_manufacturing) = 2) Then
		$person_manufacturing = " "&$person_manufacturing
	Else
		$person_manufacturing = $person_manufacturing
	EndIf

	$person_retail = $tableau_personcsv[$x][8]
	If (StringLen($person_retail) = 1) Then
		$person_retail = "  "&$person_retail
	ElseIf (StringLen($person_retail) = 2) Then
		$person_retail = " "&$person_retail
	Else
		$person_retail = $person_retail
	EndIf

	$person_raw = $tableau_personcsv[$x][9]
	If (StringLen($person_raw) = 1) Then
		$person_raw = "  "&$person_raw
	ElseIf (StringLen($person_raw) = 2) Then
		$person_raw = " "&$person_raw
	Else
		$person_raw = $person_raw
	EndIf

	;Découpage en tableau de caractères
	$t_person_name = StringSplit($person_name, "")
	$t_person_sex = StringSplit($person_sex, "")
	$t_person_code = StringSplit($person_code, "")
	$t_person_training = StringSplit($person_training, "")
	$t_person_advertise = StringSplit($person_advertise, "")
	$t_person_rd = StringSplit($person_rd, "")
	$t_person_farming = StringSplit($person_farming, "")
	$t_person_manufacturing = StringSplit($person_manufacturing, "")
	$t_person_retail = StringSplit($person_retail, "")
	$t_person_raw = StringSplit($person_raw, "")

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

	$pointeur = $pointeur + 19
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