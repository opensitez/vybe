;; vybe-test: wast/wat_text_abbreviations/inline_global_export_parses
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs
;; vybe-test-mode: compile

(module (global (export "g") i32 (i32.const 0)))
