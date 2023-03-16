;
;   Copyright (c) Ludger Solbach. All rights reserved.
;   The use and distribution terms for this software are covered by the
;   Eclipse Public License 1.0 (http://opensource.org/licenses/eclipse-1.0.php)
;   which can be found in the file license.txt at the root of this distribution.
;   By using this software in any fashion, you are agreeing to be bound by
;   the terms of this license.
;   You must not remove this notice, or any other, from this software.
;
(ns org.soulspace.cmp.poi.example.excel.read-workbook1
  (:require [clojure.java.io :as io]
            [org.soulspace.cmp.poi.excel :as xl]))

(def wb (xl/create-workbook (io/file "data/test/TestSheet1.xlsx")
                  {:missingCellPolicy (xl/missing-cell-policy :create-null-as-blank)}))
;                  {:missingCellPolicy (xl/missing-cell-policy :return-null-and-blank)}
;                  {:missingCellPolicy (xl/missing-cell-policy :return-blank-as-null)}

(def sheet0 (xl/get-sheet wb 0))

(def row0 (xl/get-row sheet0 0))
(def row1 (xl/get-row sheet0 1))
(def row2 (xl/get-row sheet0 2))

;(all-sheet-values wb 4)

(defn all-row-cells
  "Returns a sequence of values of the sheets in the workbook."
  ([]
   (xl/all-row-values xl/*sheet* 0))
  ([sheet]
   (xl/all-row-values sheet 0))
  ([sheet min-cells]
   (map #(xl/get-all-cells % min-cells) (filter seq (xl/get-all-rows sheet)))))

(defn get-all-sheet-cells
  [min-cells]
  (map #(all-row-cells %  min-cells) (xl/get-sheets wb)))

;(println "(xl/get-sheets wb):" (xl/get-sheets wb))
;(println "(xl/get-rows sheet0):" (xl/get-rows sheet0))
;(println "(xl/get-cells row2):" (xl/get-cells row2))

;(println "(xl/cell-values row0):" (xl/cell-values row0))
;(println "(xl/cell-values row2):" (xl/cell-values row2))

;(println "(xl/row-values sheet0):" (xl/row-values sheet0))
;(println "(xl/sheet-values wb):" (xl/sheet-values wb))
