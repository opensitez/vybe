;; vybe-test: wast/wat_f32_arithmetic/test_f32_nearest_up
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
  f32.const 3.8
  f32.nearest
  f32.const 4.0 call $vybe_check_f32
)
)
