;; vybe-test: wast/wat_text_abbreviations/test_folded_if_expression
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          (call $log (if (result i32) (i32.const 1) (then (i32.const 7)) (else (i32.const 8))))))
