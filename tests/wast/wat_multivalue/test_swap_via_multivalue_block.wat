;; vybe-test: wast/wat_multivalue/test_swap_via_multivalue_block
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          i32.const 10 i32.const 20
          block (param i32 i32) (result i32 i32) end
          i32.sub call $log))
