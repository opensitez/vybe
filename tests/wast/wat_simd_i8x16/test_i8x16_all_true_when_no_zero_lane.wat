;; vybe-test: wast/wat_simd_i8x16/test_i8x16_all_true_when_no_zero_lane
;; origin: languages/wast/tests/wast/test_wat_simd_i8x16.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        v128.const i8x16 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16
        i8x16.all_true call $log)
)
