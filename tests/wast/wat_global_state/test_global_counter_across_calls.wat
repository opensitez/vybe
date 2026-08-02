;; vybe-test: wast/wat_global_state/test_global_counter_across_calls
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $c (mut i32) (i32.const 0))
        (func $inc global.get $c i32.const 1 i32.add global.set $c)
        (func (export "_start")
          call $inc call $inc call $inc call $inc global.get $c call $log))
