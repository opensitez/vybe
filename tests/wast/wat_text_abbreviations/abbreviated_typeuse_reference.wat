;; vybe-test: wast/wat_text_abbreviations/abbreviated_typeuse_reference
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs
;; vybe-test-mode: compile

(module (type $t (func (param i32) (result i32))) (func (type $t) (param i32) (result i32) local.get 0))
