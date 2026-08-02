;; vybe-test: wast/wat_text_abbreviations/test_folded_block_with_result
;; origin: languages/wast/tests/wast/test_wat_text_abbreviations.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          (call $log (block (result i32) (i32.const 100)))))
