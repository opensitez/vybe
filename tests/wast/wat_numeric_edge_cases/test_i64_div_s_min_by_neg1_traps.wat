;; vybe-test: wast/wat_numeric_edge_cases/test_i64_div_s_min_by_neg1_traps
;; origin: languages/wast/tests/wast/test_wat_numeric_edge_cases.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i64.const -9223372036854775808 i64.const -1 i64.div_s call $log_i64)
)
