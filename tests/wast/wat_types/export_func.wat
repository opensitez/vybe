;; vybe-test: wast/wat_types/export_func
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (func $f) (export "func" (func $f)))
