;; vybe-test: wast/wat_module/export_func
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (func $f) (export "f" (func $f)))
