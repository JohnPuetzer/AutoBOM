(setq partNamesLookupExcelFile (strcat (vl-filename-directory (findfile "AutoBOM.lsp")) "\\lib\\FITTINGS PARTS.xlsx")) ;_ end of setq

(vl-load-com)

(defun getCellsFunction (fileName sheetName cellName / myXL myBook mySheet myRange cellValue)
  
	 (setq myXL (vlax-create-object "Excel.Application"))
	 (vla-put-visible myXL :vlax-false)
	 (vlax-put-property myXL 'DisplayAlerts :vlax-false)
	 (setq myBook (vl-catch-all-apply 'vla-open (list (vlax-get-property myXL "WorkBooks") fileName)) ;_ end of vl-catch-all-apply
	 ) ;_ end of setq
	 (setq mySheet	(vl-catch-all-apply	'vlax-get-property
								(list (vlax-get-property myBook "Sheets") "Item" sheetName) ;_ end of list
				) ;_ end of vl-catch-all-apply
	 ) ;_ end of setq
	 (vlax-invoke-method mySheet "Activate")
	 (setq myRange	(vlax-get-property (vlax-get-property mySheet 'Cells) "Range" cellName) ;_ end of vlax-get-property
	 ) ;_ end of setq
	 (setq cellValue (vlax-variant-value (vlax-get-property myRange 'Value2))) ;_ end of setq
	 (vl-catch-all-apply 'vlax-invoke-method (list myBook "Close")) ;_ end of vl-catch-all-apply
	 
	 (if	(not (vlax-object-released-p myRange))
		  (progn (vlax-release-object myRange) (setq myRange nil))
	 ) ;_ end of if
	 (if	(not (vlax-object-released-p mySheet))
		  (progn (vlax-release-object mySheet) (setq mySheet nil))
	 ) ;_ end of if
	 (if	(not (vlax-object-released-p myBook))
		  (progn (vlax-release-object myBook) (setq myBook nil))
	 ) ;_ end of if
	 (if	(not (vlax-object-released-p myXL))
		  (progn (vlax-release-object myXL) (setq myXL nil))
	 ) ;_ end of if
	 (if	(= 'safearray (type cellValue))
		  (progn
			   (setq cellValue (mapcar '(lambda (x) (mapcar 'vlax-variant-value x)) (vlax-safearray->list cellValue)))
		  ) ;_ end of progn
	 ) ;_ end of if
 ;    cellValue
) ;_ end of defun

(defun _GetClipBoardText	(/ htmlfile result)
	 ;;  Attribution: Reformatted version of
	 ;;  post by Patrick_35 at theswamp.org.
	 ;;
	 ;;  See http://tinyurl.com/2ngf4r.
	 (setq result (vlax-invoke (vlax-get (vlax-get (setq htmlfile (vlax-create-object "htmlfile")) 'ParentWindow) 'ClipBoardData)
						  'GetData
						  "Text"
			    ) ;_ vlax-invoke
	 ) ;_ setq
	 (vlax-release-object htmlfile)
	 result
) ;_ defun
(defun GetClipFromExcel ()
	 (setq L1 (_GetClipBoardText))
	 (setq L2 (LM:StringSubst "" "\n" L1))
	 (SETQ L3 (LM:str->lst L2 "\r"))
	 (SETQ L4 (vl-remove "" l3))
	 (setq i 0)
	 (while (< i (length l4))
		  (setq l4 (LM:SubstNth (LM:str->lst (nth i l4) "\t") i l4)
			   i	 (1+ i)
		  ) ;_ setq
	 ) ;_ while
	 (setq blank nil)
	 (repeat (length (car l4)) (setq blank (append '("") blank))) ;_ end of repeat
	 (setq l4 (vl-remove blank l4))
) ;_ defun

(defun LM:str->lst (str del / len lst pos)
	 (setq len (1+ (strlen del)))
	 (while (setq pos (vl-string-search del str))
		  (setq lst (cons (substr str 1 pos) lst)
			   str (substr str (+ pos len))
		  ) ;_ setq
	 ) ;_ while
	 (reverse (cons str lst))
) ;_ defun

(defun LM:StringSubst (new old str / inc len)
	 (setq len (strlen new)
		  inc 0
	 ) ;_ setq
	 (while (setq inc (vl-string-search old str inc))
		  (setq str (vl-string-subst new old str inc)
			   inc (+ inc len)
		  ) ;_ setq
	 ) ;_ while
	 str
) ;_ defun

(defun LM:SubstNth (a n l / i)
	 (setq i -1)
	 (mapcar '(lambda (x)
				 (if	(= (setq i (1+ i)) n)
					  a
					  x
				 ) ;_ if
			) ;_ lambda
		    l
	 ) ;_ mapcar
) ;_ defun

 ; SET DYNAMIC BLOCK ATTRIBUTE VALUE (CUSTOM)
(defun LM:setattributevalue (blk tag val / end enx)
	 (while (and (null end) (setq blk (entnext blk)) (= "ATTRIB" (cdr (assoc 0 (setq enx (entget blk)))))) ;_ end of and
		  (if (= (strcase tag) (strcase (cdr (assoc 2 enx))))
			   (if (entmod (subst (cons 1 val) (assoc 1 (reverse enx)) enx))
				    (progn (entupd blk) (setq end val)) ;_ end of progn
			   ) ;_ end of if
		  ) ;_ end of if
	 ) ;_ end of while
) ;_ end of defun

 ; GET DYNAMIC BLOCK ATTRIBUTE VALUES (CUSTOM)
(defun LM:getattributevalues (blk / enx lst)
	 (while (and (setq blk (entnext blk)) (= "ATTRIB" (cdr (assoc 0 (setq enx (entget blk)))))) ;_ end of and
		  (setq lst (cons (cons (cdr (assoc 2 enx)) (cdr (assoc 1 (reverse enx)))) ;_ end of cons
					   lst
				  ) ;_ end of cons
		  ) ;_ end of setq
	 ) ;_ end of while
	 (reverse lst)
) ;_ end of defun

 ; GET DYNAMIC BLOCK PROPERTERY VALUE (CUSTOM)
(defun LM:getdynpropvalue (blk prp)
	 (setq blk (vlax-ename->vla-object blk)) ;adapted for enames
	 (setq prp (strcase prp))
	 (vl-some	'(lambda (x)
				  (if (= prp (strcase (vla-get-propertyname x)))
					   (vlax-get x 'value)
				  ) ;_ end of if
			 ) ;_ end of lambda
			(vlax-invoke blk 'getdynamicblockproperties)
	 ) ;_ end of vl-some
) ;_ end of defun

 ; SET DYNAMIC BLOCK PROPERTERY VALUE (CUSTOM)
(defun LM:setdynpropvalue (blk prp val)
	 (setq blk (vlax-ename->vla-object blk)) ;adapted for enames
	 (setq prp (strcase prp))
	 (vl-some	'(lambda (x)
				  (if (= prp (strcase (vla-get-propertyname x)))
					   (progn	(vla-put-value	x
										(vlax-make-variant val (vlax-variant-type (vla-get-value x))) ;_ end of vlax-make-variant
							) ;_ end of vla-put-value
							(cond (val)
								 (t)
							) ;_ end of cond
					   ) ;_ end of progn
				  ) ;_ end of if
			 ) ;_ end of lambda
			(vlax-invoke blk 'getdynamicblockproperties)
	 ) ;_ end of vl-some
) ;_ end of defun

 ; ---------------------------------------------------------------------------------

 ; SUBSTITUTE VALUE IN LIST
 ; a - val
 ; n - position
 ; l - list
(defun LM:SubstNth (a n l / i)
	 (setq i -1)
	 (mapcar '(lambda (x)
				 (if	(= (setq i (1+ i)) n)
					  a
					  x
				 ) ;_ end of if
			) ;_ end of lambda
		    l
	 ) ;_ end of mapcar
) ;_ end of defun

 ;SUBSTITUTE VALUE IN 2D LIST
(defun SubstNth2 (ij val lst /) ;for modifying lists
 ;ij - list of location in 2d array '(i j) or '(i "j") to replace dotted list
 ;val - what to replace at that location
 ;lst - list to work on
	 (LM:SubstNth (LM:SubstNth val (cadr ij) (nth (car ij) lst)) (car ij) lst) ;_ end of LM:SubstNth
) ;_ end of defun

 ; GET SLICE OF LIST
(defun LM:sublst (lst idx len / rtn)
	 (setq len (if	len
				  (min len (- (length lst) idx))
				  (- (length lst) idx)
			 ) ;_ end of if
		  idx (+ idx len)
	 ) ;_ end of setq
	 (repeat len (setq rtn (cons (nth (setq idx (1- idx)) lst) rtn))) ;_ end of repeat
) ;_ end of defun

 ; GET ANONYMOUS BLOCK (CHANGED DYANMIC BLOCKS)
(defun LM:getanonymousreferences (blk / ano def lst rec ref)
	 (setq blk (strcase blk))
	 (while (setq def (tblnext "block" (null def)))
		  (if (and (= 1 (logand 1 (cdr (assoc 70 def))))
				 (setq rec (entget (cdr (assoc 330
										 (entget (tblobjname "block" (setq ano (cdr (assoc 2 def)))) ;_ end of tblobjname
										 ) ;_ end of entget
								    ) ;_ end of assoc
							    ) ;_ end of cdr
						 ) ;_ end of entget
				 ) ;_ end of setq
			 ) ;_ end of and
			   (while	(and (not (member ano lst)) (setq ref (assoc 331 rec))) ;_ end of and
				    (if (and (entget (cdr ref)) (wcmatch (strcase (LM:al-effectivename (cdr ref))) blk)) ;_ end of and
						(setq lst (cons ano lst))
				    ) ;_ end of if
				    (setq rec (cdr (member (assoc 331 rec) rec)))
			   ) ;_ end of while
		  ) ;_ end of if
	 ) ;_ end of while
	 (reverse lst)
) ;_ end of defun

;; Effective Block Name  -  Lee Mac
;; ent - [ent] Block Reference entity
(defun LM:al-effectivename (ent / blk rep)
	 (if	(wcmatch (setq blk (cdr (assoc 2 (entget ent)))) "`**")
		  (if (and (setq rep (cdadr (assoc	-3
									(entget (cdr (assoc	330
													(entget (tblobjname "block" blk)) ;_ end of entget
											   ) ;_ end of assoc
										   ) ;_ end of cdr
										   '("acdbblockrepbtag")
									) ;_ end of entget
							   ) ;_ end of assoc
						 ) ;_ end of cdadr
				 ) ;_ end of setq
				 (setq rep (handent (cdr (assoc 1005 rep))))
			 ) ;_ end of and
			   (setq blk (cdr (assoc 2 (entget rep))))
		  ) ;_ end of if
	 ) ;_ end of if
	 blk
) ;_ end of defun

;;; get selection of all dynamic blocks, which are anyonomous with no block name, just handle names
(defun getblockselection	(blk)
	 (ssget "ALL"
		   (list '(0 . "INSERT")
			    (cons 410 (getvar "CTAB"))
			    (cons	2
					(apply 'strcat
						  (cons blk
							   (mapcar '(lambda (x) (strcat ",`" x)) (LM:getanonymousreferences blk)) ;_ end of mapcar
						  ) ;_ end of cons
					) ;_ end of apply
			    ) ;_ end of cons
		   ) ;_ end of list
	 ) ;_ end of ssget
) ;_ end of defun

;; Popup  -  Lee Mac
;; A wrapper for the WSH popup method to display a message box prompting the user.
;; ttl - [str] Text to be displayed in the pop-up title bar
;; msg - [str] Text content of the message box
;; bit - [int] Bit-coded integer indicating icon & button appearance
;; Returns: [int] Integer indicating the button pressed to exit

(defun LM:popup (ttl msg bit / wsh rtn)
	 (if	(setq wsh (vlax-create-object "wscript.shell"))
		  (progn (setq rtn (vl-catch-all-apply 'vlax-invoke-method (list wsh 'popup msg 0 ttl bit)))
			    (vlax-release-object wsh)
			    (if (not (vl-catch-all-error-p rtn))
					rtn
			    ) ;_ if
		  ) ;_ progn
	 ) ;_ if
) ;_ defun


 ;---------------------------------------------------------------------------------------  END OF LEE MAC CODE


;;; checks if layer is currently frozen in model or viewport. layerToCheck is a str name of the layer
(defun LayerFreezeCheck (layerToCheck / x a b) ;get entity name of desired layer and check if it exists [layer 0 is the example]
	 (if	(setq a (tblobjname "LAYER" layerToCheck))
		  (progn ;one method to get viewport entity data and create a list of vpfrozen layer entities [DFX 331]
			   (setq b nil)
			   (foreach x (entget
							(ssname (ssget	"X"
										(list (cons 0 "VIEWPORT") (cons 410 (getvar "CTAB")) (cons 69 (getvar "CVPORT"))) ;_ end of list
								   ) ;_ end of ssget
								   0
							) ;_ end of ssname
					    ) ;_ end of entget
				    (if (= (car x) 331)
						(setq b (cons (cdr x) b))
				    ) ;if
			   ) ;foreach
 ;test if desired layer is in entity list
			   (cond ((member a b) "Frozen") ; layer is frozen
				    (T "Unfrozen") ; layer is unfrozen
			   ) ;cond
		  ) ;progn
	 ) ;if
) ;_ end of defun



;;;transform UCS point to Paper Space Point. Must have Model space active in active Viewport
(defun UCS2PS (PT /) (trans (trans PT 0 2) 2 3)) ;_ end of defun

 ;defun

;;;Get viewport location as 2 corners
;;;returns list of ((lower corner 2d point)(upper corner 2d point))
;;;must have model space active in required viewport

(defun viewportDisp	()
	 (setq o (vlax-safearray->list
				(vlax-variant-value
					 (vla-get-center (vla-get-activepviewport (vla-get-activedocument (vlax-get-acad-object)) ;_ end of vla-get-activedocument
								  ) ;_ end of vla-get-activepviewport
					 ) ;_ end of vla-get-center
				) ;_ end of vlax-variant-value
		    ) ;_ end of vlax-safearray->list
	 ) ;_ end of setq
	 (setq x (vla-get-width (vla-get-activepviewport (vla-get-activedocument (vlax-get-acad-object)) ;_ end of vla-get-activedocument
					    ) ;_ end of vla-get-activepviewport
		    ) ;_ end of vla-get-width
	 ) ;_ end of setq
	 (setq y (vla-get-height	(vla-get-activepviewport	(vla-get-activedocument (vlax-get-acad-object)) ;_ end of vla-get-activedocument
						) ;_ end of vla-get-activepviewport
		    ) ;_ end of vla-get-height
	 ) ;_ end of setq
	 (setq rlist (list (list (- (car o) (/ x 2)) (- (cadr o) (/ y 2))) ;_ end of list
				    (list (+ (car o) (/ x 2)) (+ (cadr o) (/ y 2))) ;_ end of list
			   ) ;_ end of list
	 ) ;_ end of setq
) ;_ end of defun

;;;Loop through Selection set
;;;lst - selection set
;;;returns list of ent names

(defun LoopSS (lst / j rlst)
	 (setq j	  0 ;get all tags enames
		  rlst nil
	 ) ;_ end of setq
	 (while (< j (sslength lst)) (setq rlst (cons (ssname lst j) rlst)) (setq j (1+ j))) ;_ end of while
	 rlst
) ;_ end of defun





;; Selection Set Bounding Box  -  Lee Mac - tweaked by JTP for getting single block centroid
;; Returns the centroid in WCS coordinates 
;; sel - [sel] single ename 

(defun GetCentroid (sel / idx llp ls1 ls2 obj urp)
	 (setq obj (vlax-ename->vla-object sel))
	 (if	(and	(vlax-method-applicable-p obj 'getboundingbox)
			(not (vl-catch-all-error-p (vl-catch-all-apply 'vla-getboundingbox (list obj 'llp 'urp))))
		) ;_ and
		  (setq ls1 (mapcar	'min
						(vlax-safearray->list llp)
						(cond (ls1)
							 ((vlax-safearray->list llp))
						) ;_ cond
				  ) ;_ mapcar
			   ls2 (mapcar	'max
						(vlax-safearray->list urp)
						(cond (ls2)
							 ((vlax-safearray->list urp))
						) ;_ cond
				  ) ;_ mapcar
		  ) ;_ setq
	 ) ;_ if
	 (if	(and ls1 ls2)
		  (progn (setq a (list ls1 ls2))
			    (setq cent (mapcar '(lambda (x) (/ x 2)) (mapcar '+ (car a) (cadr a))))
		  ) ;_ progn
	 ) ;_ if
) ;_ defun

 ; get value in assocation list.
 ; val = string. Case Sensitive Association lookup
 ; lst = lst to act on
(defun GetAV (val lst /) (cdr (assoc val lst)))

 ; set value in assocation list.
 ; val = string. Case Sensitive Association lookup
 ; a = value to replace
 ; lst = lst to act on
(defun SetAV (val a lst /) (subst (cons val s) (assoc val lst) lst))

(defun attachXdata (ename regappName xdata /)
	 (if	(not (tblsearch "APPID" regappName)) ;check if the application has been registered
		  (regapp regappName)
	 ) ;_ if
	 (entmod (append (entget ename) (LIST (list -3 (LIST regappName (cons 1000 xdata))))) ;_ end of append
	 ) ;_ end of entmod
) ;_ defun
(defun getxdata (ename xdataName /)
	 (cdr (car (cdr (car (cdr (assoc -3
							   (entget ename (LIST xdataName)) ;_ end of entget
						 ) ;_ end of assoc
					 ) ;_ end of cdr
				 ) ;_ end of car
			 ) ;_ end of cdr
		 ) ;_ end of car
	 ) ;_ cdr
) ;_ defun
 ;(PRINT (getxdata (car (entsel)) "MATERIAL"))



(defun isDictPartListCurrentDate (SavedDate excelFile /)
	 (if	(= SavedDate (VL-FILE-SYSTIME excelFile))
    T
		  nil
	 ) ;_ end of if
) ;_ end of defun

(defun getDict (dictName /) (vlax-ldata-get "dict" dictName)) ;_ end of defun
(defun setDict (dictName lst /) (vlax-ldata-put "dict" dictName lst)) ;_ end of defun


 ;TODO
 ;Check if part list exist local and up to date
 ;Check if dict part list exist and up to date
 ;
(defun getPartDataFromExcel ()
	 (setq dictPartSaveDate (getDict "PartDataAllMasterLastSave")
		  partDataAll	    (getDict "PartDataAllMaster")
	 ) ;_ end of setq
	 (if	(or 
          (not partDataAll) ;_ end of not
		      (not (isDictPartListCurrentDate dictPartSaveDate partNamesLookupExcelFile) ;_ end of isDictPartListCurrentDate
		      ) ;_ end of not
		) ;_ end of or
		  (progn (princ "getting part data, please wait ......\n")
			    (setq	partDataAll (getCellsFunction partNamesLookupExcelFile "Sheet1" "A1:D500") ;_ end of GetExcel
			    ) ;_ end of setq
			    (princ "part data loaded!\n")
			    (setDict "PartDataAllMaster" partDataAll)
			    (setDict "PartDataAllMasterLastSave" (VL-FILE-SYSTIME partNamesLookupExcelFile)) ;_ end of vlax-ldata-put
			    (vl-propagate 'partDataAll)
		  ) ;_ end of progn
	 ) ;_ end of if
) ;_ end of defun

(defun c:updatePart(/ ss)
  (setq allEnames (LoopSS (ssget '((0 . "INSERT")))) )
  (foreach partEname allEnames
	(setq raw (cdr (assoc 2 (entget partEname))))
	(setq raw (LM:StringSubst " " "_"  raw))
	(setq raw (LM:StringSubst "" "ELBOW " raw))
	(setq raw (strcat raw ".dwg"))
    (if (setq partLookupVal (assoc raw  partDataAll))
	(progn
    
    (attachXdata partEname "PART" (nth 1 partLookupVal))
    (attachXdata partEname "DESCRIPTION" (nth 2 partLookupVal)) ;_ end of attachXdata
    (attachXdata partEname "UOM" (nth 3 partLookupVal))
    (attachXdata partEname "MATERIAL" "FS")
    )
	)
  )
)