;; vybe-test: wast/wat_module/export_global
;; origin: languages/wast/tests/wast/test_wat_module.rs
;; vybe-test-mode: compile

(module (global i32 (i32.const 42)) (export "g" (global 0)))
