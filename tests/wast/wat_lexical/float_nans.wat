;; vybe-test: wast/wat_lexical/float_nans
;; origin: languages/wast/tests/wast/test_wat_lexical.rs
;; vybe-test-mode: compile

(module (global f32 (f32.const nan)))
