;; vybe-test: wast/wat_float_conversions/test_f32_reinterpret_i32
;; origin: languages/wast/tests/wast/test_wat_float_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 1065353216 ;; 0x3f800000 = 1.0f
  f32.reinterpret_i32
  call $log_f32
)
)
