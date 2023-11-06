
(setq LibLocation (strcat (vl-filename-directory (findfile "AutoBOM.lsp")) "\\lib"))
(setq filenameStandardParts (strcat LibLocation "\\FITTINGS BASE NAMES" ))

(VL-LOAD-ALL (findfile "SUBFUNCTIONS.lsp")) ; VL-LOAD-ALL
(VL-LOAD-ALL (findfile  "CountPipes.LSP")) ; VL-LOAD-ALL


(getPartDataFromExcel)

(setq PipeODLookup
       '(
	 ("0.250" "0.54")
	 ("0.500" "0.84")
	 ("0.750" "1.05")
	 ("1.000" "1.315")
	 ("1.250" "1.66")
	 ("1.500" "1.9")
	 ("2.000" "2.375")
	 ("2.500" "2.875")
	 ("3.000" "3.5")
	 ("4.000" "4.5")
	 ("5.000" "5.563")
	 ("6.000" "6.625")
	 ("8.000" "8.625")
	 ("10.00" "10.75")
	 ("12.00" "12.75")
	 ("14.00" "14")
	 ("16.00" "16")
	 ("18.00" "18")
	 ("20.00" "20")
	 ("24.00" "24")
	)
) ;_ end of setq

(defun reverseLookupOD->NPS (OD /)
  (cadr (assoc OD (mapcar 'reverse PipeODLookup)))
) ;_ end of defun

(defun atan2 (x y /)
  (setq r (* (/ 180 pi) (atan y x)))
  (if (= (abs r) r)
    r
    (+ r 360)
  ) ;_ end of if
) ;_ end of defun

;; rtos wrapper  -  Lee Mac
;; A wrapper for the rtos function to negate the effect of DIMZIN

(defun LM:rtos (real units prec / dimzin result)
  (setq dimzin (getvar 'dimzin))
  (setvar 'dimzin 0)
  (setq result (vl-catch-all-apply 'rtos (list real units prec)))
  (setvar 'dimzin dimzin)
  (if (not (vl-catch-all-error-p result))
    result
  ) ;_ end of if
) ;_ end of defun

 ;_ end of defun


(defun insertPart (partType OD p1 /)

  (if (= (LENGTH P1) 2)
    (progn
      (setq p2 (cadr p1)
	    p1 (car p1)
      ) ;_ end of setq
      (setq
	rotation (- 0
		    (atan2 (cadr (trans p2 0 1)) (car (trans p2 0 1)))
		 ) ;_ end of -
      ) ;_ end of setq
    ) ;_ end of progn
    (setq rotation "")
  ) ;_ end of if

  (setq partNPS (reverseLookupOD->NPS OD))

  (if (< (atof partNPS) 2.5)
    (setq pipeSize "SW"
	  endAdd   " SW 3000.dwg"
    ) ;_ end of setq
    (setq pipeSize "BW"
	  endAdd   " BW.dwg"
    ) ;_ end of setq
  ) ;_ end of if

  (setq	partFilenameLocation
	 (strcat filenameStandardParts
		 "\\"		 pipeSize
		 "\\"		 partType
		 "\\"
		) ;_ end of strcat

	partFilename
	 (strcat partNPS
		 " "
		 partType
		 endAdd
	 ) ;_ end of strcat
  ) ;_ end of setq
  (setq	partFileLocation
	 (strcat partFilenameLocation
		 partFilename
	 ) ;_ end of strcat
  ) ;_ end of setq


  (command
    "-insert"
    partFileLocation
    "_NON"
    (trans p1 0 1)
    ""
    ""
    rotation
  ) ;_ end of command

  
  (setq lastPartEname (entlast))
  (setq partLookupVal (assoc partFilename partDataAll))
  (attachXdata lastPartEname "PART" (nth 1 partLookupVal))
  (attachXdata lastPartEname "DESCRIPTION" (nth 2 partLookupVal)) ;_ end of attachXdata
  (attachXdata lastPartEname "UOM" (nth 3 partLookupVal))
  (attachXdata lastPartEname "MATERIAL" "FS")

) ;_ end of defun



(defun insertPipe (pp1 pp2 first OD /)
  (setq	fudgefactor
	 (if (>= (atof (reverseLookupOD->NPS OD)) 2.5)
	   (* (atof (reverseLookupOD->NPS OD)) 1.5)
	   (atof (reverseLookupOD->NPS OD))
	 ) ;_ end of if
  ) ;_ end of setq
     ;how much to recede pipe from intersection based on 90 size. by extending/receding the start/end 3d points Z component. [INCHES]
  (setq	p1n (list (car (trans pp1 0 1)) ; x
		  (cadr (trans pp1 0 1)) ; y
		  (+ (caddr (trans pp1 0 1))
		     (if first
		       0
		       fudgefactor
		     ) ;_ end of if
		  ) ; z
	    ) ;_ end of list
  ) ;_ end of setq
  (setq	p2n (list (car (trans pp2 0 1)) ; x
		  (cadr (trans pp2 0 1)) ; y 
		  (- (caddr (trans pp2 0 1)) fudgefactor) ; z
	    ) ;_ end of list
  ) ;_ end of setq
  (command
     ;make pipe cylinder from point 1 (p1) to point 2 (p2) with diameter of actual pipe OD
    "_.cylinder"
    "_NON"
    p1n
    "_NON"
    (/ (atof OD) 2)
    "_NON"
    p2n
  ) ;_ end of command

  (attachXdata (entlast) "MATERIAL" "FS")
  (attachXdata (entlast) "PIPE" "T")
) ;_ end of defun

(defun makeUcs (p1u p2u /)
  (command ; make new ucs in with Z axis in direction of pipe
    "_.ucs"
    "_ZA"
    "_NON"
    (trans p1u 0 1) ;WCS TO UCS
    "_NON"
    (trans p2u 0 1) ;WCS TO UCS
  ) ;_ end of command
) ;_ end of defun

(defun prevUcs ()
  (command
    "_.ucs"
    "_P"
  ) ;_ end of command
) ;_ end of defun

(defun checkIfStraight (p2 p3 /)
  (and
    (< (abs (- (car (trans p3 0 1)) (car (trans p2 0 1)))) 0.1)
    (< (abs (- (cadr (trans p3 0 1)) (cadr (trans p2 0 1))))
       0.1
    ) ;_ end of <
  ) ;_ end of and
) ;_ end of defun



(defun MaKePipeRun (p1 p2 p3 OD first s plist /)
  (setvar "NAVVCUBEDISPLAY" 0)
  (setvar "UCSICON" 0)
  (setvar "GRIDDISPLAY" 0)

  (if (not first)
    (progn
      (makeUcs p3 p1)
      (if (checkIfStraight p2 p3)
	(progn
	  (setq p1 p3)
	  (setq
	    plist (vl-remove (nth (- (length plist) 2) plist) plist)
	  ) ;_ end of setq
	  (entdel (entlast))
	  (prevUcs)
	) ;_ end of progn
	(progn

	  (insertPart "90" OD (list p1 p2))
	  (prevUcs)
	) ;_ end of progn
      ) ;_ end of if
    ) ;_ end of progn
  ) ;_ end of if

  (makeUcs p1 p2) ; make new ucs in with Z axis in direction of pipe  
  (insertPipe p1 p2 first OD)
  (prevUcs)

  (setvar "NAVVCUBEDISPLAY" (car s))
  (setvar "UCSICON" (cadr s))
  (setvar "GRIDDISPLAY" (caddr s))
  plist

) ;_ end of defun



(defun getPipeOD ()
  (setq prevPipeOD (reverseLookupOD->NPS G:pipeOD))
  (setq	msg (strcat "Pipe NPS"
		    (if	prevPipeOD
		      (strcat " [" prevPipeOD "]")
		      ""
		    ) ;_ end of if
		    ":"
	    ) ;_ end of strcat
  ) ;_ end of setq

  (setq getpipeNPS (getreal msg))
  (setq	addTrailingZerosPipe
	 (if getpipeNPS
	   (substr (LM:rtos getpipeNPS 2 3) 1 5)
	   999
	 ) ;_ end of if
  ) ;_ end of setq
  (setq pipeODi (car (getAV addTrailingZerosPipe PipeODLookup)))
  (if pipeODi ;_ end of setq
    (setq G:pipeOD pipeODi)
    (if	(and prevPipeOD (null pipeODi))
      prevPipeOD
      (getPipeOD)
    ) ;_ end of if
  ) ;_ end of if
) ;_ end of defun

(defun getpointWCS (p1 msg /)
  (if (not p1)
    (trans (getpoint msg) 1 0)
    (trans (getpoint (trans p1 0 1) msg) 1 0)
  ) ;_ end of if
) ;_ end of defun

(setq G:pipeOD nil)

(defun c:AutoPIPE (/ *error* plist settings)

  (setq	settings (list (getvar "NAVVCUBEDISPLAY")
		       (getvar "UCSICON")
		       (getvar "GRIDDISPLAY")
		 ) ;_ end of list
  ) ;_ end of setq
  (setq plist nil)

  (defun *error* (msg)
    (command-s "_.ucs" "p")
    (princ (strcat "\nMake Pipe Exit! Reason = " msg "\n"))
    (setvar "NAVVCUBEDISPLAY" (car settings))
    (setvar "UCSICON" (cadr settings))
    (setvar "GRIDDISPLAY" (caddr settings))
  ) ;_ end of defun

  (defun runAutoPipe (pipeODInput / p1 p2 p3)
    (if	(not plist)
     ; if no first point, it is the begining of the sequence. get pipe diameter
      (progn
	(setq G:pipeOD (getPipeOD))
	(setq p1 (getpointWCS nil "\nEnter the pipe 1st point: "))
	(setq p2 (getpointWCS p1 "\nEnter the pipe 2nd point: "))
     ;     (setq p1 (car (setq r (getp1 "\nEnter the pipe 1st point: " "\ndirection:"))))
     ;     (setq p2 (getp2 r "\nEnter the pipe 2nd point: " ))
	(setq plist (append plist (list p1 p2)))
	(setq plist (MaKePipeRun p1 p2 nil G:pipeOD T settings plist))
	(runAutoPipe G:pipeOD)
      ) ;_ end of progn
      (progn
	(setq p3 (nth (- (length plist) 2) plist))
	(setq p1 (nth (- (length plist) 1) plist))
     ; (setq r (getp1 p1 "\ndirection:"))
     ; (setq p2 (getp2 r "\nEnter the pipe 2nd point: " ))

	(setq p2 (getpointWCS p1 "\nEnter the pipe 2nd point [Free]: "))

	(setq plist (append plist (list p2)))
	(setq plist (MaKePipeRun p1 p2 p3 G:pipeOD nil settings plist))
	(runAutoPipe G:pipeOD)
      ) ;_ end of progn
    ) ;_ end of if
  ) ;_ end of defun
  (runAutoPipe nil)
) ;_ end of defun

(defun c:AutotPart ()

  (setq G:pipeOD (getPipeOD))
  (initget "90 45 Tee Coupling Cap Reducer")
  (setq select (getkword "insert [90/45/Tee/Coupling/Cap/Reducer]"))

  (if (= select "Reducer")
    (progn
      (initget "Coupling Insert Tee Eccentric Concentric")
      (setq
	select2	(getkword
		  "insert [Coupling/Insert/Tee/Eccentric/Concentric]"
		) ;_ end of getkword
      ) ;_ end of setq
      (setq pipeNPS2 (reverseLookupOD->NPS (getPipeOD)))
      (cond
	((= select2 "Tee") (setq select "REDUCER\\REDUCING TEE\\"))
      ) ;_ end of cond
    ) ;_ end of progn
    (insertPart select (reverseLookupOD->NPS G:pipeOD) '(0 0 0))

  ) ;_ end of if
) ;_ end of defun

(defun c:updatePart(/ ss partEname allEnames)
  (setq allEnames (LoopSS (ssget '((0 . "INSERT")))) )
  (foreach partEname allEnames
    (setq partLookupVal (assoc (STRCAT (CDR (ASSOC 2 (ENTGET partEname))) ".dwg")  partDataAll))
    (attachXdata partEname "PART" (nth 1 partLookupVal))
    (attachXdata partEname "DESCRIPTION" (nth 2 partLookupVal)) ;_ end of attachXdata
    (attachXdata partEname "UOM" (nth 3 partLookupVal))
    (attachXdata partEname "MATERIAL" "FS")
  )
)