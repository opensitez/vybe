;; vybe-test: wast/wat_multivalue/test_block_two_results
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          block (result i32 i32) i32.const 7 i32.const 8 end i32.add call $log))
