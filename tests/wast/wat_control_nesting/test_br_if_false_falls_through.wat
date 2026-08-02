;; vybe-test: wast/wat_control_nesting/test_br_if_false_falls_through
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          block i32.const 1 call $log i32.const 0 br_if 0 i32.const 2 call $log end))
