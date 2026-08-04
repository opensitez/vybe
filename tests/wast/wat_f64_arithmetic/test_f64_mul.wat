;; vybe-test: wast/wat_f64_arithmetic/test_f64_mul
;; origin: languages/wast/tests/wast/test_wat_f64_arithmetic.rs

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
  (func (export "_start")
  f64.const 3.0
  f64.const 2.5
  f64.mul
  f64.const 7.5 call $vybe_check_f64
)
)
