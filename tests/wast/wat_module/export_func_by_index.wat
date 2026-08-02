;; vybe-test: wast/wat_module/export_func_by_index
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func) (export "f" (func 0)))
