;; vybe-test: wast/wat_lexical/id_numeric_indices
;; origin: languages/wast/tests/wast/test_wat_lexical.rs
;; vybe-test-mode: compile

(module (func $0) (func $1) (func (export "test") call $0 call $1))
