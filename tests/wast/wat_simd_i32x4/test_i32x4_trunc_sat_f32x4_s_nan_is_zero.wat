;; vybe-test: wast/wat_simd_i32x4/test_i32x4_trunc_sat_f32x4_s_nan_is_zero
;; origin: languages/wast/tests/wast/test_wat_simd_i32x4.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const f32x4 nan 0 0 0 i32x4.trunc_sat_f32x4_s i32x4.extract_lane 0 call $log)
)
