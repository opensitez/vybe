;; vybe-test: wast/wat_lexical/string_invalid_escapes
;; origin: languages/wast/tests/wast/test_wat_lexical.rs
;; vybe-test-mode: compile-fail

(module (import "env" "\z" (func)))
