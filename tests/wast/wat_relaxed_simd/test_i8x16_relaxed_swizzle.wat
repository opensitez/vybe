;; vybe-test: wast/wat_relaxed_simd/test_i8x16_relaxed_swizzle
;; origin: languages/wast/tests/wast/test_wat_relaxed_simd.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 10 20 30 40 50 60 70 80 90 100 110 120 5 15 25 35
        v128.const i8x16 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.relaxed_swizzle i8x16.extract_lane_u 0 call $log)
)
