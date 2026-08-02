;; vybe-test: wast/wat_simd_i16x8/test_i16x8_splat_truncates_to_16
;; origin: languages/wast/tests/wast/test_wat_simd_i16x8.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const 0x1FFFF i16x8.splat i16x8.extract_lane_u 0 call $log)
)
