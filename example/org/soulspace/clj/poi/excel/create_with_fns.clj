;
;   Copyright (c) Ludger Solbach. All rights reserved.
;   The use and distribution terms for this software are covered by the
;   Eclipse Public License 1.0 (http://opensource.org/licenses/eclipse-1.0.php)
;   which can be found in the file license.txt at the root of this distribution.
;   By using this software in any fashion, you are agreeing to be bound by
;   the terms of this license.
;   You must not remove this notice, or any other, from this software.
;
(ns org.soulspace.cmp.poi.example.excel.create-with-fns
  (:require [clojure.java.io :as io]
             [org.soulspace.cmp.poi.excel :as xl]))

;
; creating excel workbooks with functions
;
(defn create-example []
  (let [wb (xl/create-workbook {})]
    ;(println wb)
    (let [sht (xl/create-sheet wb {})]
      ;(println sht)
      (let [rw (xl/create-row sht {})]
        (xl/create-cell rw {} "N0001")
        (xl/create-cell rw {} "V13")
        (xl/create-cell rw {} 1.5))
      (let [rw (xl/create-row sht {})]
        (xl/create-cell rw {} "N0002")
        (xl/create-cell rw {} "V14")
        (xl/create-cell rw {} 2.3))
      (let [rw (xl/create-row sht {})]
        (xl/create-cell rw {} "N0003")
        (xl/create-cell rw {} "V15")
        (xl/create-cell rw {} 4.5)))
    (with-open [out (io/output-stream "TestFN.xlsx")]
      (xl/write-workbook out wb))))

(comment
  (create-example))