;; vybe-test: wast/wat_numeric_edge_cases/test_i32_rem_s_min_by_neg1_is_zero
;; origin: languages/wast/tests/wast/test_wat_numeric_edge_cases.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const -2147483648 i32.const -1 i32.rem_s call $log)
)
