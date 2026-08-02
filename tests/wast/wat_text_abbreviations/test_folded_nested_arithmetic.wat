;; vybe-test: wast/wat_text_abbreviations/test_folded_nested_arithmetic
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          (call $log (i32.mul (i32.add (i32.const 2) (i32.const 3)) (i32.const 4)))))
