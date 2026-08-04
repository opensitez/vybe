;; vybe-test: wast/wat_f32_arithmetic/test_f32_max
;; origin: languages/wast/tests/wast/test_wat_f32_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f32 (param f32) (param f32)
    local.get 0
    local.get 1
    f32.ne
    if
      unreachable
    end)
  (func (export "_start")
  f32.const 3.0
  f32.const 5.0
  f32.max
  f32.const 5.0 call $vybe_check_f32
)
)
