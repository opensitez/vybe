;; vybe-test: wast/wat_stack_switching/cont_type_definition
;; origin: languages/wast/tests/wast/test_wat_stack_switching.rs
;; vybe-test-mode: compile

(module (type $ft (func)) (type $ct (cont $ft)))
