;; vybe-test: wast/wat_simd_i64x2/test_i64x2_shr_s
;; origin: languages/wast/tests/wast/test_wat_simd_i64x2.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i64x2 -1024 0 i32.const 2 i64x2.shr_s i64x2.extract_lane 0 call $log_i64)
)
