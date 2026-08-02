;; vybe-test: wast/wat_i32_arithmetic/test_i32_rem_s_neg_pos
;; origin: languages/wast/tests/wast/test_wat_i32_arithmetic.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") i32.const -20 i32.const 3 i32.rem_s call $log)
)
