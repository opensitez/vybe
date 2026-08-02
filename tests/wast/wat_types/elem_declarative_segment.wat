;; vybe-test: wast/wat_types/elem_declarative_segment
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (func $f) (elem declare funcref (ref.func $f)))
