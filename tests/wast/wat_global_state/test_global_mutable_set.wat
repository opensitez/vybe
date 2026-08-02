;; vybe-test: wast/wat_global_state/test_global_mutable_set
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g (mut i32) (i32.const 0))
        (func (export "_start") i32.const 77 global.set $g global.get $g call $log))
