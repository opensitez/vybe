;; vybe-test: wast/wat_lexical/int_invalid_octal_binary_wat
;; origin: languages/wast/tests/wast/test_wat_lexical.rs
;; vybe-test-mode: compile-fail

(module (global i32 (i32.const 0b1010)))
