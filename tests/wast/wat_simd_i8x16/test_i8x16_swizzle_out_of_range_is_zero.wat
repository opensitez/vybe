;; vybe-test: wast/wat_simd_i8x16/test_i8x16_swizzle_out_of_range_is_zero
;; origin: languages/wast/tests/wast/test_wat_simd_i8x16.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 10 20 30 40 50 60 70 80 90 100 110 120 5 15 25 35
        v128.const i8x16 16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.swizzle i8x16.extract_lane_u 0 call $log)
)
