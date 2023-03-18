(ns org.soulspace.cmp.poi.excel-test
  (:require [clojure.test :refer :all]
            [clojure.java.io :as io]
            [org.soulspace.cmp.poi.excel :as xl])
  (:import [org.apache.poi.ss.usermodel Cell Row Sheet]))

(def create-null-as-blank (xl/missing-cell-policy :create-null-as-blank))
(def return-blank-as-null (xl/missing-cell-policy :return-blank-as-null))
(def return-null-and-blank (xl/missing-cell-policy :return-null-and-blank))

(defn count-cells
  "Returns the count of the cells in a given row.
   The cells are retrieved by 'get-all-cells'."
  ([wb sheet row]
   (-> wb
       (xl/get-sheet sheet)
       (xl/get-row row)
       (xl/get-all-cells)
       (count))))

(deftest workbook-test
  (let [wb (xl/create-workbook (io/file "data/test/Workbook1.xlsx")
                               {:missingCellPolicy return-null-and-blank})
        sheet1 (xl/get-sheet wb 0)
        sheet2 (xl/get-sheet wb 1)
        sheet3 (xl/get-sheet wb 2)]
    (are [x y] (= x y)
      (count (xl/get-sheets wb)) 3
      (count (xl/get-rows sheet1)) 1
      (count (xl/get-rows sheet2)) 4
      (count (xl/get-rows sheet3)) 6
      (xl/max-cell-num sheet1) 3
      (xl/max-cell-num sheet2) 3
      (xl/max-cell-num sheet3) 5
      (count-cells wb 0 0) 3
      (count-cells wb 1 0) 3
      (count-cells wb 2 0) 5
      (count-cells wb 2 1) 5
      (count-cells wb 2 2) 5
      (count-cells wb 2 3) 5)))

(deftest color-test
  (let [wb (xl/create-workbook {:missingCellPolicy return-null-and-blank})
        c1 (xl/color wb 0x000000)
        c2 (xl/color wb 0x777777)
        c3 (xl/color wb 0xFFFFFF)]
    (are [x y] (= x y)
      (.isIndexed c1) false
      (.isIndexed c2) false
      (.isIndexed c3) false)))

(comment
  (run-tests))