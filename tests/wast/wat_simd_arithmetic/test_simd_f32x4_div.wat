;; vybe-test: wast/wat_simd_arithmetic/test_simd_f32x4_div
;; origin: languages/wast/tests/wast/test_wat_simd_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  v128.const f32x4 1.5 2.5 10.0 4.5
  v128.const f32x4 0.5 1.5 2.0 3.5
  f32x4.div
  f32x4.extract_lane 2
  call $log_f32
)
)
