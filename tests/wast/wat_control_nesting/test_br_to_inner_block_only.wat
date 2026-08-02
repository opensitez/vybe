;; vybe-test: wast/wat_control_nesting/test_br_to_inner_block_only
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          block block i32.const 1 call $log br 0 i32.const 2 call $log end
          i32.const 3 call $log end))
