;; vybe-test: wast/wat_simd_i32x4/test_i32x4_le_u
;; origin: languages/wast/tests/wast/test_wat_simd_i32x4.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i32x4 1 0 0 0 v128.const i32x4 -1 0 0 0
        i32x4.le_u i32x4.extract_lane 0 call $log)
)
