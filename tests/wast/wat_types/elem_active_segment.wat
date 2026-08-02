;; vybe-test: wast/wat_types/elem_active_segment
;; origin: languages/wast/tests/wast/test_wat_types.rs
;; vybe-test-mode: compile

(module (table 1 funcref) (func $f) (elem (i32.const 0) $f))
