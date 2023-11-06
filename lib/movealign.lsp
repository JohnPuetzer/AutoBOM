(defun ga (p1 p2 p3 /)
  (setq dir1 (mapcar '- p2 p1))
  (setq	dir2
	 (mapcar
	   'abs
	   (mapcar
	     '(lambda (x) (/ x (sqrt (apply '+ (mapcar '* dir1 dir1)))))
	     dir1
	   ) ;_ end of mapcar
	 ) ;_ end of mapcar

  ) ;_ end of setq
  (cond
    ((= (car dir2) 1) (list (car (mapcar '- p3 p1)) 0 0))
    ((= (cadr dir2) 1) (list 0 (cadr (mapcar '- p3 p1)) 0))
    ((= (caddr dir2) 1) (list 0 0 (caddr (mapcar '- p3 p1))))
  ) ;_ end of cond
) ;_ end of defun

(defun getp1 (msg1 msg2 /)
  (if (= (type msg1) 'STR)
    (setq a (getpoint msg1))
    (setq a msg1)
  ) ;_ end of if
  (setq	osm   (getvar 'osmode)
	omode (getvar "orthomode")
	pmode (getvar "autosnap")
  ) ;_ end of setq
  (setvar 'osmode (logior 16384 (getvar 'osmode)))
  (setvar "orthomode" 1)
  (setq b (getpoint a msg2))

  (setvar 'osmode osm)
  (setvar "orthomode" omode)
  (setvar "autosnap" pmode)
  (list (trans a 1 0 ) (trans b 1 0 ))

) ;_ end of defun
(defun getp2 (r msg /)

  (setq	a (car r)
	b (cadr r)
  ) ;_ end of setq
  (princ msg)
  (setq c (getpoint))
  (trans (mapcar '+ a (ga a b c)) 1 0)
) ;_ end of defun


(defun c:qe ()

  (command "move"
	   (ssget)
	   ""
	   (car (setq r (getp1 "\nbase point:" "\ndirection:" T)))
	   (getp2 r "\nsnap to point:")
  ) ;_ end of command
) ;_ end of defun
