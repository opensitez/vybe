;; vybe-test: wast/wat_numeric_edge_cases/test_i32_div_u_by_zero_traps
;; origin: languages/wast/tests/wast/test_wat_numeric_edge_cases.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const 5 i32.const 0 i32.div_u call $log)
)
