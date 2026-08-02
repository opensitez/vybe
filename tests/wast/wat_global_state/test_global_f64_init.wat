;; vybe-test: wast/wat_global_state/test_global_f64_init
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
        (global $pi f64 (f64.const 3.5))
        (func (export "_start") global.get $pi call $log_f64))
