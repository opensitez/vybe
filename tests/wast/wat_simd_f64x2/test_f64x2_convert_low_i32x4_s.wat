;; vybe-test: wast/wat_simd_f64x2/test_f64x2_convert_low_i32x4_s
;; origin: languages/wast/tests/wast/test_wat_simd_f64x2.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i32x4 -7 0 0 0 f64x2.convert_low_i32x4_s f64x2.extract_lane 0 call $log_f64)
)
