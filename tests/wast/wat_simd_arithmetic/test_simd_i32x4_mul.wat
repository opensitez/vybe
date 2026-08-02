;; vybe-test: wast/wat_simd_arithmetic/test_simd_i32x4_mul
;; origin: languages/wast/tests/wast/test_wat_simd_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  v128.const i32x4 10 20 30 40
  v128.const i32x4 5 25 25 35
  i32x4.mul
  i32x4.extract_lane 0
  call $log
)
)
