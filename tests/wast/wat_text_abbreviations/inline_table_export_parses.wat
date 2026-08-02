;; vybe-test: wast/wat_text_abbreviations/inline_table_export_parses
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs
;; vybe-test-mode: compile

(module (table (export "t") 1 funcref))
