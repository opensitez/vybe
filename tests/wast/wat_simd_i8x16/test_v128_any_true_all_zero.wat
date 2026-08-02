;; vybe-test: wast/wat_simd_i8x16/test_v128_any_true_all_zero
;; origin: languages/wast/tests/wast/test_wat_simd_i8x16.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.any_true call $log)
)
