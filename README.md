cmp.poi
=======
The cmp.poi library is an [Apache POI](https://poi.apache.org/) wrapper in Clojure.
It contains functions and macros to create, read, update and write MS Excel workbooks.

Usage
-----
### Dependency
[![Clojars Project](https://img.shields.io/clojars/v/org.soulspace.clj/cmp.poi.svg)](https://clojars.org/org.soulspace.clj/cmp.poi)

### Example
In a new example namespace we require the excel namespace and import the IndexedColors class for a bit of color.

```
(ns org.soulspace.cmp.poi.excel.examples
  (:require [org.soulspace.cmp.poi.excel :as xl])
  (:import [org.apache.poi.ss.usermodel IndexedColors]))
```
Now we create a brand new excel workbook with some sample data and write it to disk.

```
(defn create-example
  "Create new workbook, fill in some data and write it to disk."
  []
  (xl/write-workbook 
    "TestMACRO.xlsx"
    (xl/new-workbook {}
              (let [style1 (xl/new-cell-style {:fillForegroundColor (xl/color-index IndexedColors/AQUA)
                                               :fillPattern (xl/fill-pattern-type :solid-foreground)
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

(create-example)
```

Finally we read the new excel workbook again and update a sheet with a new row of data.

```
(defn update-example
  "Update an existing workbook with a new row of data in the first sheet."
  []
  (xl/with-workbook "TestMACRO.xlsx"
    (xl/select-sheet 0 {}
           (xl/new-row {}
                (xl/new-cell {} 1.5)
                (xl/new-cell {} 2.5)
                (xl/new-cell {} 3.5)))))
                
(update-example)
```

Copyright
---------
© 2013-2021 Ludger Solbach

License
-------
[Eclipse Public License 1.0](http://www.eclipse.org/legal/epl-v10.html)

Code Repository
---------------
[CljComponents on GitHub](https://github.com/soulspace-org/cmp.poi)

