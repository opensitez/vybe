;; vybe-test: wast/wat_types/invalid_import_after_func_definition
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile-fail

(module (func) (import "env" "print" (func)))
