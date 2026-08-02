;; vybe-test: wast/wat_types/export_global
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (global $g i32 (i32.const 0)) (export "g" (global $g)))
