;; vybe-test: wast/wat_simd_f32x4/test_f32x4_sqrt
;; origin: languages/wast/tests/wast/test_wat_simd_f32x4.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const f32x4 16.0 0 0 0 f32x4.sqrt f32x4.extract_lane 0 call $log_f32)
)
