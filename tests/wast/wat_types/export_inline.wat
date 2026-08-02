;; vybe-test: wast/wat_types/export_inline
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (func (export "f")) (global (export "g") i32 (i32.const 0)))
