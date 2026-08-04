;; vybe-test: wast/wat_global_state/test_global_f64_init
;; origin: languages/wast/tests/wast/test_wat_global_state.rs

(module
        (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f64 (param f64) (param f64)
    local.get 0
    local.get 1
    f64.ne
    if
      unreachable
    end)
        (global $pi f64 (f64.const 3.5))
        (func (export "_start") global.get $pi f64.const 3.5 call $vybe_check_f64))
