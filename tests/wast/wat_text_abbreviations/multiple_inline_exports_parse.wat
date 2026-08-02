;; vybe-test: wast/wat_text_abbreviations/multiple_inline_exports_parse
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs
;; vybe-test-mode: compile

(module (func $f (export "a") (export "b") (result i32) i32.const 1))
