;; vybe-test: wast/wat_number_literals/test_large_float_exponent
;; origin: languages/wast/tests/wast/test_wat_number_literals.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f64 (param f64) (param f64)
    local.get 0
    local.get 1
    f64.ne
    if
      unreachable
    end)
  (func (export "_start") f64.const 1e100 f64.const 1e100 f64.div f64.const 1.0 call $vybe_check_f64)
)
