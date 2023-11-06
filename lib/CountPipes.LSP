(setq PipeNomDLookup
       '(("0.54" "0.25")
         ("0.84" "0.5")
         ("1.05" "0.75")
         ("1.315" "1")
         ("1.66" "1.25")
         ("1.66" "1-1/4")
         ("1.9" "1.5")
         ("1.9" "1-1/2")
         ("2.375" "2")
         ("2.875" "2.5")
         ("2.875" "2-1/2")
         ("3.5" "3")
         ("4.5" "4")
         ("5.563" "5")
         ("6.625" "6")
         ("8.625" "8")
         ("10.75" "10")
         ("12.75" "12")
         ("14" "14")
         ("16" "16")
         ("18" "18")
         ("20" "20")
         ("24" "24")

        )
) ;_ end of setq

(defun pipe_geom (e / i v r l)
  (setq e (vlax-ename->vla-object e)
        v (vlax-get e 'Volume)
        i ((lambda (x)
             (if
               (equal (car x) (cadr x) (* 0.001 (car x)))
                (list (caddr x) (car x))
                (if
                  (equal (car x) (caddr x) (* 0.001 (car x)))
                   (list (cadr x) (car x))
                   (if
                     (equal (cadr x) (caddr x) (* 0.001 (cadr x)))
                      (list (car x) (cadr x))
                   ) ;_ end of if
                ) ;_ end of if
             ) ;_ end of if
           ) ;_ end of lambda
            (vlax-get e 'PrincipalMoments)
          )
  ) ;_ end of setq
  (if i
    (progn
      (setq r (sqrt (/ (* 2.0 (car i)) v))
            l (/ v (* pi r r))
      ) ;_ end of setq
      (if
        (equal (/ (cadr i) (/ (* v (+ (* 3.0 r r) (* l l))) 12.0))
               1.0
               1e-5
        ) ;_ end of equal
         (list (* 2.0 r) (LM:roundup (/ l 12) 0.5))
     ;round to nearest half foot



      ) ;_ end of if
    ) ;_ end of progn
  ) ;_ end of if
) ;_ end of defun

;; Round Up  -  Lee Mac
;; Rounds 'n' up to the nearest 'm'

(defun LM:roundup (n m)
  ((lambda (r)
     (cond ((equal 0.0 r 1e-8) n)
           ((< n 0) (- n r))
           ((+ n (- m r)))
     ) ;_ end of cond
   ) ;_ end of lambda
    (rem n m)
  )
) ;_ end of defun

(defun JP:GetBlockEnts (BlockNameI Filter / MainEnt NextEnt EntLst Lst)
  (setq MainEnt (tblobjname "block" BlockNameI))
   (setq NextEnt (cdr (assoc -2 (entget MainEnt)))
  ) ;_ end of setq
  (while NextEnt
    (setq EntLst (entget NextEnt))
    (if (wcmatch (cdr (assoc 0 EntLst)) (strcase Filter))
      (setq Lst (cons EntLst Lst))
    ) ;_ end of if
    (setq NextEnt (entnext NextEnt))
  ) ;_ end of while
  (reverse Lst)
) ;_ end of defun

;;;(setq pipelist (JP:GetBlockEnts "Drawing2" "3dSolid"))
;;;(setq
;;;  n1 nil
;;;) ;_ end of setq
;;;(foreach i pipelist
;;;  (setq pipespecs (pipe_geom (cdr (car i))))
;;;  (if pipespecs
;;;    (if (setq pnomd (assoc (rtos (car pipespecs)) PipeNomDLookup))
;;;      (setq
;;;        n1
;;;         (append
;;;           n1
;;;           (list
;;;             (append (list pnomd (cadr pipespecs) (cdr (assoc 8 i))))
;;;           ) ;_ end of list
;;;         ) ;_ end of append
;;;      ) ;_ end of setq
;;;    ) ;_ end of if
;;;  ) ;_ end of if
;;;) ;_ end of foreach
;;;(setq n1 (vl-remove nil n1))
;;;     ;(setq l (vl-sort l '(lambda (a b) (< (car a) (car b))))) sort array