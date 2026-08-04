;; vybe-test: wast/wat_float_conversions/test_f64_reinterpret_i64
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs

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
  i64.const 4607182418800017408 ;; 0x3FF0000000000000 = 1.0
  f64.reinterpret_i64
  f64.const 1.0 call $vybe_check_f64
)
)
