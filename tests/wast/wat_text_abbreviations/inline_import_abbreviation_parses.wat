;; vybe-test: wast/wat_text_abbreviations/inline_import_abbreviation_parses
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs
;; vybe-test-mode: compile

(module (func $f (import "m" "n") (param i32) (result i32)))
