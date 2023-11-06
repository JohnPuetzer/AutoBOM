 ;Load all subfunctions
(setq LibLocation (strcat (findfile "AutoBOM.lsp") "\\lib"))

(VL-LOAD-ALL (findfile "SUBFUNCTIONS.lsp")) ; VL-LOAD-ALL
(VL-LOAD-ALL (findfile  "CountPipes.LSP")) ; VL-LOAD-ALL
(VL-LOAD-ALL (findfile "MakePipes.lsp")) ; VL-LOAD-ALL


(defun AddTable (lst / pt1 rows columns rowheight colwidth ent doc curspace obj objtable)
	 (setq doc (vla-get-activedocument (vlax-get-acad-object)))
	 (setq curspace (vla-get-paperspace doc)) ; now do table 
	 (setq pt1 (vlax-3d-point (getpoint "\nPick point for top left hand of table:  ")))
	 (setq rows (length lst)) ; TAGS LENGTH
	 (setq columns 7)
	 (setq rowheight 0.15)
	 (setq colwidth 2)
	 (setq objtable (vla-addtable curspace pt1 (+ rows 2) columns rowheight colwidth) ; vla-addtable
	 ) ;_ setq
	 (vla-SetText objtable 0 0 (strcat "%<\\AcVar ctab \\f \"%tc1\">%" " MATERIAL TAKEOFF")) ; vla-SetText
	 (foreach	i '(0 1 2 3 4 5 6)
		  (vla-SetText	objtable
					1
					I
					(NTH	I
						(LIST "REF#" "PART" "DESCRIPTION" "QTY." "UOM." "GROUP ID" "SHIP TO") ; LIST
					) ; NTH
		  ) ; vla-SetText
	 ) ; foreach
	 (setq i 0) ; ROW COUNTER
	 (while (< i rows) ; LOOP THROUGH TAG ROWS
		  (setq j	   0
			   lstj nil
		  ) ; setq
		  (while (< j columns) ; LOOP THROUGH TAG COLUMNS
			   (vla-settext objtable (+ i 2) j (cdr (nth j (nth i lst))))
			   (setq j (1+ j))
		  ) ; while
		  (setq i (1+ i))
	 ) ; while
	 (princ LSTI)
	 (vlax-release-object objtable)
	 (princ)
) ; defun

(defun GetTable ()
	 (setq table (vlax-ename->vla-object (car (entsel))))
	 (setq columns (vlax-get-property table 'Columns))
	 (setq rows (vlax-get-property table 'rows))
	 (setq lsti nil
		  i	  0
	 ) ; setq
	 (while (< i rows)
		  (setq j	   0
			   lstj nil
		  ) ; setq
		  (while (< j columns)
			   (setq lstj (append lstj (list (vlax-invoke table 'GetCellValue i j)))) ; setq
			   (setq j (1+ j))
		  ) ; while
		  (setq lsti (append lsti (list lstj)))
		  (setq i (1+ i))
	 ) ; while
	 (princ LSTI)
) ; defun


(setq RefLookupVals	'("A" "B"	"C" "D" "E" "F" "G"	"H" "J" "K" "L" "M"	"N" "P" "Q" "R" "S"	"T" "U" "V" "W" "X"	"Y" "Z" "AA" "AB" "AC" "AD" "AE" "AF"
				  "AG" "AH" "AJ" "AK" "AL" "AM" "AN" "AP" "AQ" "AR" "AS" "AT" "AU" "AV" "AW" "AX" "AY" "AZ" "BA" "BB" "BC" "BD" "BE" "BF" "BG"
				  "BH" "BJ" "BK" "BL" "BM" "BN" "BP" "BQ" "BR" "BS" "BT" "BU" "BV" "BW" "BX" "BY" "BZ" "CA" "CB" "CC" "CD" "CE" "CF" "CG" "CH"
				  "CJ" "CK" "CL" "CM" "CN" "CP" "CQ" "CR" "CS" "CT" "CU" "CV" "CW" "CX" "CY" "CZ" "DA" "DB" "DC" "DD" "DE" "DF" "DG" "DH" "DJ"
				  "DK" "DL" "DM" "DN" "DP" "DQ" "DR" "DS" "DT" "DU" "DV" "DW" "DX" "DY" "DZ"
				 )
) ; setq




(defun TagRefNumberLookup (i /) ; this is dumb but its easier than writing code ti skip I,J,AI,AJ, etc. only goes to 120   
	 (if	(= (type i) 'INT)
		  (nth i RefLookupVals) ;straight lookup - pass int to get char
		  (- (length RefLookupVals) (length (member (strcase i) RefLookupVals))) ;reverse lookup. pass char to get index in array
	 ) ;if
) ;defun

;;; (vla-put-color  (vlax-ename->vla-object (car e)) 5)
;;; (vla-get-color (vlax-ename->vla-object (car e)))



(setq partData	nil
	 pipeData	nil
) ; setq

(defun GetPartAttFromModel (allS / blockName)
	 (foreach	xx allS
		  (cond ((assoc 1
					 (tblsearch "BLOCK" (setq blockName (cdr (assoc 2 (entget xx))))) ; tblsearch
			    ) ; assoc
			    (GetPartAttFromModel
					(mapcar '(lambda (xx) (cdr (assoc -1 xx))) (JP:GetBlockEnts blockname "3dsolid,insert"))
			    ) ; GetPartAttFromModel
			   ) ;  if block is a xref, get sub entities and find parts/pipes
			   ((setq partDataxx (getxdata xx "PART")) ; if block has part attribute, pull its part data
			    (setq	partData (append partData
								  (list (list (cons "PART" partDataxx)
										    (cons "DESCRIPTION" (getxdata xx "DESCRIPTION"))
										    (cons "UOM" (getxdata xx "UOM"))
										    (cons "GROUP_ID" "")
										    (cons "SHIP_TO_LOC" "")
										    (cons "ENAME" xx)
										    (cons "Layer" (cdr (assoc 8 (entget xx))))
										    (cons "Position" (GetCentroid xx)) ; cons
										    (cons "Material" (getxdata XX "MATERIAL")) ; cons
									   ) ; list
								  ) ; list
						    ) ; append
			    ) ; setq
			   )
			   ((getxdata XX "PIPE") ;if it is a pipe, pull its pipe data
			    (setq	pipeData (append pipeData
								  (list (append (pipe_geom xx)
											 (list (cons "ENAME" xx)
												  (cons "Layer" (cdr (assoc 8 (entget xx)))) ; cons
												  (cons "Material" (getxdata XX "MATERIAL"))
											 ) ; list
									   ) ; append
								  ) ; list
						    ) ; append
			    ) ; setq
			   )
		  ) ;cond
	 ) ;foreach 
) ; defun



(defun GetTagDataFromLayout (flag /)
	 (cond ((= flag "Select")
		   (setq ssl (ssget '((0 . "~multileader") (0 . "~viewport") (0 . "~mtext"))) ; ssget
		   ) ; ssget
		  )
		  ((= flag "All")
		   (setq ssl (getblockselection "BOM_ITEM_TAG_V1")) ; ssget
		  )
	 ) ; cond
	 (setq TagData	'()
		  i 0
	 ) ; setq
	 (repeat (sslength ssl)
		  (setq TagData (cons (append (LM:getattributevalues (ssname ssl i)) (list (cons "ENAME" (ssname ssl i)))) ; append
						  TagData
					 ) ; cons
		  ) ; setq
		  (setq i (1+ i))
	 ) ;repeat
	 (setq TagData (reverse TagData)) ;RETURN LIST OF BLOCK ATTRIBUTES
) ; defun


(defun PutTagAtBlock (TagData VPsize / CurrentLayer CurrentPoint CurrentPart tag)
	 (setq CurrentPoint	(GetAV "PS_Position" TagData) ;Tag point
		  CurrentPart	(GetAV "PART" TagData)
		  CurrentDesc	(GetAV "DESCRIPTION" TagData)
		  CurrentUOM	(GetAV "UOM" TagData)
	 ) ;setq
	 (setq ptx (car CurrentPoint)
		  pty (cadr CurrentPoint)
		  lx	 (car (car VPsize))
		  ly	 (cadr (car VPsize))
		  hx	 (car (cadr VPsize))
		  hy	 (cadr (cadr VPsize))
	 ) ; setq
	 (if	(not (or (< ptx lx) (> ptx hx) (< pty ly) (> pty hy)))
		  (command "-insert" (strcat LibLocation "\\" "BOM_ITEM_TAG_V1.dwg" )
			   	CurrentPoint;insert point	     
				 "" ;scale
				 "" ;rotation
				 "" ;ref#
				 CurrentPart ;part
				 CurrentDesc ;description
				 "" ;qty
				 CurrentUOM ;uom
				 "" ;group id
				 "" ;ship to loc
				) 
	 )
	 ) 

(defun SortAndCheckDuplicateTags (lst / i j k jtag ctag)
 ;takes a list from exported tag data  lst = (GetTagDataFromLayout) abd returns list with REF# sorted and block entity names "ENAME" aappended to each tag
	 (setq i 0) ; Tag number counter
	 (setq k 0) ; Refence # counter
	 (while (< i (length lst))
		  (setq ctag (nth i lst))
		  (setq j 0) ;old tag number counter
		  (if (= i 0)
			   (progn ;if first in list
				    (setq	lst (SubstNth2 (list i 0) (cons "REF#" (TagRefNumberLookup k)) lst) ; SubstNth2
				    ) ; setq
				    (setq k (1+ k))
			   ) ;else after first
			   (progn
				    (while (< j i)
						(setq jtag (nth j lst))
						(if (and (= (cdr (assoc "PART" ctag)) (cdr (assoc "PART" jtag))) ; =
							    (= (cdr (assoc "SHIP_TO_LOC" ctag)) (cdr (assoc "SHIP_TO_LOC" jtag))) ; =
							    (= (cdr (assoc "GROUP_ID" ctag)) (cdr (assoc "GROUP_ID" jtag))) ; =
						    ) ; and
							 (progn (setq lst (SubstNth2 (list i 0) (car jtag) lst)) ; if duplicate
								   (setq j i)
							 ) ;progn
							 (if	(= j (1- i)) ;if next on list and already looped through all previous
								  (progn (setq	lst (SubstNth2 (list i 0) (cons "REF#" (TagRefNumberLookup k)) lst) ; SubstNth2
									    ) ; setq
									    (setq k (1+ k))
								  ) ;progn
							 ) ;if
						) ;if
						(setq j (1+ j))
				    ) ;while
			   ) ;progn
		  ) ;if
		  (setq i (1+ i))
	 ) ;while
	 LST
) ;defun



(defun LabelEachTag	(lstA / i) ;(setq lstA (GetTagDataFromLayout))
	 (setq lstA (SortAndCheckDuplicateTags lstA))
	 (setq i 0)
	 (foreach	tag lstA
		  (LM:setattributevalue (cdr (assoc "ENAME" tag)) "REF#" (cdr (assoc "REF#" tag))) ; LM:setattributevalue
		  (LM:setdynpropvalue (cdr (assoc "ENAME" tag)) "PAGE" i) ;CHANGE "PAGE" TO "ORDER"
		  (setq i (1+ i))
	 ) ;foreach
) ;defun




(defun InsertTagInOrder ()
	 (setq ssl (ssget '((0 . "~multileader") (0 . "~viewport") (0 . "~mtext"))) ; ssget
	 ) ;tags to insert
	 (setq st	(ssget ":S" '((0 . "~multileader") (0 . "~viewport") (0 . "~mtext"))) ; ssget
	 ) ;tag to put before/after
	 (setq Inserti (LM:getdynpropvalue (ssname st 0) "PAGE")) ;get location
	 (setq alltags (getblockselection "BOM_ITEM_TAG_V1"))
	 (setq list1 (LoopSS ssl)) ;get all selcted tags enames
	 (setq list2 (LoopSS alltags))
	 (setq listS (if (< (length list1) (length list2))
				    list1
				    list2
			   ) ; the Shorter list
		  list3 (car (vl-remove listS (list list1 list2))) ; the other one, to start
	 ) ; setq
	 (foreach x listS (setq list3 (vl-remove x list3))) ;remove inserted tags from all tags
	 (setq ll	nil
		  j	0
	 ) ; setq
	 (while (< j (length list3))
		  (setq tg (nth j list3))
		  (setq ll (cons (append	(LM:getattributevalues tg)
							(list (cons "PAGE" (LM:getdynpropvalue tg "PAGE")) (cons "ENAME" (nth j list3))) ; list
					  ) ; append
					  ll
				 ) ; cons
		  ) ; setq
		  (setq j (1+ j))
	 ) ; while
	 (foreach	tagll (if	(= (cdr (assoc "REF#" tag)) "")
					  (setq ll (vl-remove tag ll))
				 ) ; if
	 ) ; foreach
	 (setq ll	(vl-sort ll '(lambda (a b) (< (cdr (nth 7 a)) (cdr (nth 7 b))))) ; vl-sort
	 ) ;sort un selected tags
	 (setq ll2 nil
		  j	 0
	 ) ; setq
	 (while (< j (length list1))
		  (setq tg (nth j list1))
		  (setq ll2 (cons (append (LM:getattributevalues tg)
							 (list (cons "PAGE" (LM:getdynpropvalue tg "PAGE")) (cons "ENAME" (nth j list1))) ; list
					   ) ; append
					   ll2
				  ) ; cons
		  ) ; setq
		  (setq j (1+ j))
	 ) ; while
	 (setq all (append (append (LM:sublst ll 0 (1+ (fix Inserti))) ll2) ; append
				    (LM:sublst ll (1+ (fix Inserti)) (- (length ll) (1+ (fix Inserti)))) ; LM:sublst
			 ) ; append
	 ) ; setq
	 (setq all (SortAndCheckDuplicateTags all))
	 (LabelEachTag all)
) ;_ defun


(defun PlaceTagsOnViewport (/ x CurrentDPoint omode attdiamode attreqmode 3domode)
	 (setq partData nil
		  pipeData nil
	 ) ; setq
	 (GetPartAttFromModel (loopss (ssget "A" '((410 . "Model") (0 . "Insert,3DSolid")))) ; loopss
	 ) ; GetPartAttFromModel
 ;export all blocks from modelspace / xref
 ; read part data from xlsx and get x,y point on paper space
 ; remove frozen objects from data set
	 (command "MSPACE")
	 (foreach	x PartData ;loop through all individual parts
		  (if (= "Frozen" (LayerFreezeCheck (GetAV "Layer" x))) ;check if part is frozen 
			   (setq PartData (vl-remove x PartData)) ;remove if frozen true
			   (setq PartData (subst	(append x
									   (list (cons "PS_Position" (UCS2PS (GetAV "Position" x))) ; cons
									   ) ; list
								) ; append
								x
								PartData
						   ) ; subst
			   ) ; setq
		  ) ;if
	 ) ;foreach
	 (setq vpdisp (viewportDisp)) ; GET VIEWPORT WIDTH/HIEGHT
	 (command "PSPACE") ; activate paper space
	 (setq omode	   (getvar "osmode")
		  attdiamode (getvar "attdia")
		  attreqmode (getvar "attreq")
		  3domode	   (getvar "3dosmode")
	 ) ;save setting
	 (setvar "osmode" 0) ;turn off 2d osnap
	 (setvar "3dosmode" 0) ;turn off 3dosnap
	 (setvar "attdia" 0) ;turn off block attribute dialogue
	 (setvar "attreq" 1) ;turn on block attribute input
	 (foreach tag PartData (PutTagAtBlock tag vpdisp)) ; foreach
	 (setvar "osmode" omode) ;return osnap to previous settings
	 (setvar "attdia" attdiamode)
	 (setvar "attreq" attreqmode)
	 (setvar "3dosmode" 3domode)
) ;defun

(defun BOM (/ l2 l3)
	 (setq all (GetTagDataFromLayout "All"))
	 (setq all (vl-sort	all
					'(lambda (a b)
						  (< (TagRefNumberLookup (cdr (assoc "REF#" a))) (TagRefNumberLookup (cdr (assoc "REF#" b)))) ; <
					 ) ; lambda
			 ) ; vl-sort
	 ) ;sort un selected tags
	 (setq dmode (getvar "dimzin"))
	 (setvar "dimzin" 12)
	 (setq i	1 ; skip first entry. 
		  l2	(list (car all)) ;get first entry
	 ) ; setq
	 (while (< i (length all))
		  (setq l3 (nth i all))
		  (if (= (cdr (assoc "REF#" (nth i all))) (cdr (assoc "REF#" (nth (1- i) all)))) ; IF THE NEXT TAG MATCHES THE PREVIOUS
			   (progn ;TRUE
				    (setq	pq (atof (cdr (assoc "QTY" (last L2)))) ;previous qty. INCLUDES PREVIOUSLY COUNTED QTY
						cq (atof (cdr (assoc "QTY" (nth i all)))) ;current qty
				    ) ; setq
				    (setq	l2 (SubstNth2 (list (1- (length l2)) 3) (cons "QTY" (rtos (+ pq cq))) l2) ; SubstNth2
				    ) ; setq
			   ) ;true
			   (setq l2 (append l2 (list l3))) ;FALSE not duplicate. add new entry to total list
		  ) ; if
		  (setq i (1+ i))
	 ) ; while
	 (setvar "dimzin" dmode)
	 (AddTable l2)
) ; defun


(defun c:AUTOBOM ()
	 (initget "PutTags LabelTags UpdateTags BOM")
	 (setq psize (getkword (strcat "\nAUTOBOM: [PutTags/LabelTags/UpdateTags/BOM]:")) ; getkword
	 ) ; setq
	 (cond ((= psize "PutTags") ((PROGN (PlaceTagsOnViewport) (C:AUTOBOM))))
		  ((= psize "LabelTags") (LabelEachTag (GetTagDataFromLayout "Select")))
		  ((= psize "UpdateTags") (InsertTagInOrder))
		  ((= psize "BOM") (BOM))
	 ) ; cond
 ;(C:AUTOBOM)
) ; defun

(defun c:displayPartData	()
	 (while (not (setq cPart (car (entsel)))))
	 (if	(not (setq partDataxx (getxdata cPart "PART")))
		  (progn (princ "\nnot a valid part!") (exit))
	 ) ; if
 ;(if (setq partDataxx (getxdata cPart "PIPE"))
	 (setq popupTitle "AutoBOM"
		  popupMsg   (strcat "Data for part selected:"
						 "\nPart: "
						 partDataxx
						 "\nDESCRIPTION: "
						 (getxdata cPart "DESCRIPTION")
						 "\nUOM: "
						 (getxdata cPart "UOM") ;"\n GROUP_ID"
						 ;;(getxdata cPart "GROUP_ID")
						 ;;"\n SHIP_TO_LOC"
						 ;;(getxdata cPart "SHIP_TO_LOC")
				   ) ; strcat
	 ) ; setq
	 (LM:popup popupTitle popupMsg (+ 0 64 4096))
) ; defun

(defun c:autoPartReactor
  (getPartDataFromExcel)

  (if (not myReactor)
        (setq myReactor (vlr-editor-reactor nil '((:VLR-commandEnded . VTF2)))) ;_ end of setq
        (princ "already created!")
  ) ;_ end of if

  (defun VTF2 (CALL CALLBACK)
        (if (= (car CALLBACK) "-INSERT")
        (progn (setq lastPartEname (entlast))
        (setq partLookupVal (assoc (strcat (cdr (assoc 2 (entget lastPartEname))) ".dwg") partDataAll)) 
        (attachXdata lastPartEname "PART" (nth 1 partLookupVal))
        (attachXdata lastPartEname "DESCRIPTION" (nth 2 partLookupVal))
        (attachXdata lastPartEname "UOM" (nth 3 partLookupVal))
        (attachXdata lastPartEname "MATERIAL" "FS")
        ;;(alert (cdr (assoc 2 (entget(entlast)))))
        ) ;_ end of progn
        ) ;_ end of if
  ) ;_ end of defun
)