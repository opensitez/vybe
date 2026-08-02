;; vybe-test: wast/wat_types/import_table
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (import "env" "tbl" (table 1 funcref)))
