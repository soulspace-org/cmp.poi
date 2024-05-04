(defproject org.soulspace.clj/cmp.poi "0.6.5-SNAPSHOT"
  :description "The cmp.poi library is an Apache POI wrapper to create, read and update MS Excel files."
  :url "https://github.com/soulspace-org/cmp.poi"
  :license {:name "Eclipse Public License"
            :url  "http://www.eclipse.org/legal/epl-v10.html"}

  ; use deps.edn dependencies
;  :plugins [[lein-tools-deps "0.4.5"]]
;  :middleware [lein-tools-deps.plugin/resolve-dependencies-with-deps-edn]
;  :lein-tools-deps/config {:config-files [:install :user :project]}
  :dependencies [[org.clojure/clojure "1.10.1"]
                 [org.apache.poi/poi-ooxml "5.2.5"]
                 [org.soulspace.clj/clj.java "0.9.0"]]

  :test-paths ["test"]

  :repl-options {:init-ns org.soulspace.cmp.poi.excel}
  :profiles {:dev {:dependencies [[djblue/portal "0.49.1"]
                                  [criterium/criterium "0.4.6"]
                                  [com.clojure-goes-fast/clj-java-decompiler "0.3.4"]
                                      ; [expound/expound "0.9.0"]
                                  ]
                    ; :global-vars {*warn-on-reflection* true}
                   }}

  :scm {:name "git" :url "https://github.com/soulspace-org/cmp.poi"}
  :deploy-repositories [["clojars" {:sign-releases false :url "https://clojars.org/repo"}]])
