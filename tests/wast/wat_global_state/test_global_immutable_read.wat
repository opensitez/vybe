;; vybe-test: wast/wat_global_state/test_global_immutable_read
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $g i32 (i32.const 314))
        (func (export "_start") global.get $g call $log))
