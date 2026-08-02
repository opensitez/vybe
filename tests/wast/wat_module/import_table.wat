;; vybe-test: wast/wat_module/import_table
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (import "env" "t" (table 1 funcref)))
