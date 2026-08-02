;; vybe-test: wast/wat_relaxed_simd/test_f64x2_relaxed_min
;; origin: languages/wast/tests/wast/test_wat_relaxed_simd.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const f64x2 3.0 0 v128.const f64x2 8.0 0
        f64x2.relaxed_min f64x2.extract_lane 0 call $log_f64)
)
