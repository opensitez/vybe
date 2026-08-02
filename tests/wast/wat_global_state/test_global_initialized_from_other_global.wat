;; vybe-test: wast/wat_global_state/test_global_initialized_from_other_global
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (global $base i32 (i32.const 10))
        (global $derived i32 (global.get $base))
        (func (export "_start") global.get $derived call $log))
