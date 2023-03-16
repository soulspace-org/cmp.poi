;;
;;   Copyright (c) Ludger Solbach. All rights reserved.
;;   The use and distribution terms for this software are covered by the
;;   Eclipse Public License 1.0 (http://opensource.org/licenses/eclipse-1.0.php)
;;   which can be found in the file license.txt at the root of this distribution.
;;   By using this software in any fashion, you are agreeing to be bound by
;;   the terms of this license.
;;   You must not remove this notice, or any other, from this software.
;;
(ns org.soulspace.cmp.poi.excel.examples
  (:require [org.soulspace.cmp.poi.excel :as xl])
  (:import [org.apache.poi.ss.usermodel IndexedColors]))

;; 
;; this is the way I like to create excel workbooks from Clojure
;;
(defn create-example
  "Creates new workbook, fills in some data and writes it to disk."
  []
  (xl/write-workbook 
    "TestMACRO.xlsx"
    (xl/new-workbook {}
              (let [style1 (xl/new-cell-style {:fillForegroundColor (xl/color-index IndexedColors/AQUA)
                                               :fillPattern (xl/cell-fill-style :solid-foreground)
                                               :alignment (xl/horizontal-alignment :center)})]
                (xl/new-sheet {}
                           (xl/new-row {}
                                (xl/new-cell {:cellStyle style1} "N0001")
                                (xl/new-cell {} "V13")
                                (xl/new-cell {} 1.5))
                           (xl/new-row {}
                                    (xl/new-cell {:cellStyle style1} "N0002")
                                    (xl/new-cell {} "V14")
                                    (xl/new-cell {} 2.3))
                           (xl/new-row {}
                                    (xl/new-cell {:cellStyle style1} "N0003")
                                    (xl/new-cell {} "V15")
                                    (xl/new-cell {} 4.5)))))))

(defn update-example
  "Updates an existing workbook with a new row of data in the first sheet."
  []
  (xl/with-workbook "TestMACRO.xlsx"
    (xl/select-sheet 0 {}
           (xl/new-row {}
                (xl/new-cell {} 1.5)
                (xl/new-cell {} 2.5)
                (xl/new-cell {} 3.5)))))

(comment
  (create-example)
  (update-example))