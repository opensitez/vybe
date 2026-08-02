;; vybe-test: wast/wat_types/type_duplicate_definitions
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (type (func (param i32))) (type (func (param i32))))
