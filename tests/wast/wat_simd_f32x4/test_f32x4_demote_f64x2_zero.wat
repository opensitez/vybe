;; vybe-test: wast/wat_simd_f32x4/test_f32x4_demote_f64x2_zero
;; origin: languages/wast/tests/wast/test_wat_simd_f32x4.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const f64x2 3.5 0 f32x4.demote_f64x2_zero f32x4.extract_lane 0 call $log_f32)
)
