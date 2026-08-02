;; vybe-test: wast/wat_types/elem_passive_segment
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (func $f) (elem funcref (ref.func $f)))
