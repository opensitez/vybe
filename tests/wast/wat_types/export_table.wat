;; vybe-test: wast/wat_types/export_table
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (table $t 1 funcref) (export "tbl" (table $t)))
