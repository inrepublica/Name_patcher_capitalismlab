#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ico\read.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Include <Array.au3>
#include<File.au3>
;Ouverture du fichier et mise dans un tableau
$fichier_a_ouvrir = "1STD.SET"
$fichier = FileOpen($fichier_a_ouvrir, 16)
$tableau_fichier = StringRegExp(FileRead($fichier), '[[:xdigit:]]{2}', 3)
FileClose($fichier)

;Declaration des variables pour les noms des entreprises
$debut_nom = 17712
$longueur_fullname = 39
$longueur_shortname = 15
$nombre_entreprise = 35
$longueur_csv = $nombre_entreprise + 1
Dim $csv[$longueur_csv] = ["FULL NAME (max 39 characters);SHORT NAME (max 15 characters);COLOR ID (1 to 9 for color, leave a space for random)"]

;Assignation de $i au début des noms d'entreprises
$i = $debut_nom

;Récupération du noms des entreprises
For $y = 1 to $nombre_entreprise
	$ligne = ""
	;Récupération du FULLNAME
	For $i = $i to $i + $longueur_fullname
		$ligne &= Chr("0x" & $tableau_fichier[$i])
	Next
	$ligne = StringStripWS($ligne, 2)
	$ligne &= ";"
	;Récupération du SHORTNAME
	For $i = $i to $i + $longueur_shortname
		$ligne &= Chr("0x" & $tableau_fichier[$i])
	Next
	$ligne = StringStripWS($ligne, 2)
	$ligne &= ";"
	;Récupération du Color_ID
	$ligne &= Chr("0x" & $tableau_fichier[$i])
	$i = $i + 11

	;Assignation du mot dans un tableau pour export en csv
	$csv[$y] = $ligne
Next

;Exportation du tableau dans un fichier csv
_FileWriteFromArray("corporation.csv", $csv)

;Supression de la dernière ligne vide du csv
$csv = FileRead("corporation.csv")
$fichier = FileOpen("corporation.csv", 2)
FileWrite($fichier, StringTrimRight($csv, 2))
FileClose($fichier)

;Declaration des variables pour les noms des personnes
$debut_nom = 54387
$longueur_name = 29
$longueur_sex = 7
$longueur_code = 2
$nombre_personne = 64
$longueur_csv = $nombre_personne + 1
Dim $csv[$longueur_csv] = ["NAME = max 30 characters;SEX = MWestern for man FWestern for woman;CODE = F01 to F26 for female picture and M01 to M38 for male picture;TRAINING = 0 to 100;ADVERTISE = 0 to 100;R&D = 0 to 100;FARMING = 0 to 100;MANUFACTURING = 0 to 100;RETAIL = 0 to 100;RAW = 0 to 100"]

;Assignation de $i au début des noms de personne
$i = $debut_nom

;Récupération du noms des personnes
For $y = 1 to $nombre_personne
	$ligne = ""
	;Récupération du NAME
	For $i = $i to $i + $longueur_name
		$ligne &= Chr("0x" & $tableau_fichier[$i])
	Next
	$ligne = StringStripWS($ligne, 2)
	$ligne &= ";"
	;Récupération du SEX
	For $i = $i to $i + $longueur_sex
		$ligne &= Chr("0x" & $tableau_fichier[$i])
	Next
	$i = $i + 3
	$ligne &= ";"
	;Récupération du CODE
	For $i = $i to $i + $longueur_code
		$ligne &= Chr("0x" & $tableau_fichier[$i])
	Next
	$ligne &= ";"
	$i = $i + 51
	;Récupération du TRAINING
	$valeur = ""
	For $i = $i to $i + 2
		$valeur &= Chr("0x" & $tableau_fichier[$i])
	Next
	$valeur = StringStripWS($valeur, 1)
	$ligne &= $valeur
	$ligne &= ";"
	;Récupération du ADVERTISE
	$valeur = ""
	For $i = $i to $i + 2
		$valeur &= Chr("0x" & $tableau_fichier[$i])
	Next
	$valeur = StringStripWS($valeur, 1)
	$ligne &= $valeur
	$ligne &= ";"
	;Récupération du R&D
	$valeur = ""
	For $i = $i to $i + 2
		$valeur &= Chr("0x" & $tableau_fichier[$i])
	Next
	$valeur = StringStripWS($valeur, 1)
	$ligne &= $valeur
	$ligne &= ";"
	;Récupération du FARMING
	$valeur = ""
	For $i = $i to $i + 2
		$valeur &= Chr("0x" & $tableau_fichier[$i])
	Next
	$valeur = StringStripWS($valeur, 1)
	$ligne &= $valeur
	$ligne &= ";"
	;Récupération du MANUFACTURING
	$valeur = ""
	For $i = $i to $i + 2
		$valeur &= Chr("0x" & $tableau_fichier[$i])
	Next
	$valeur = StringStripWS($valeur, 1)
	$ligne &= $valeur
	$ligne &= ";"
	;Récupération du RETAIL
	$valeur = ""
	For $i = $i to $i + 2
		$valeur &= Chr("0x" & $tableau_fichier[$i])
	Next
	$valeur = StringStripWS($valeur, 1)
	$ligne &= $valeur
	$ligne &= ";"
	;Récupération du RAW
	$valeur = ""
	For $i = $i to $i + 2
		$valeur &= Chr("0x" & $tableau_fichier[$i])
	Next
	$valeur = StringStripWS($valeur, 1)
	$ligne &= $valeur

	$i = $i + 1

	;Assignation du mot dans un tableau pour export en csv
	$csv[$y] = $ligne
Next

;Exportation du tableau dans un fichier csv
_FileWriteFromArray("person.csv", $csv)

;Supression de la dernière ligne vide du csv
$csv = FileRead("person.csv")
$fichier = FileOpen("person.csv", 2)
FileWrite($fichier, StringTrimRight($csv, 2))
FileClose($fichier)

;Message de réussite
MsgBox(0, "Export Done", "Export of corporation.csv and person.csv successfull.")