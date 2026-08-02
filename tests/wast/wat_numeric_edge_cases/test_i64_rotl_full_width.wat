;; vybe-test: wast/wat_numeric_edge_cases/test_i64_rotl_full_width
;; origin: languages/wast/tests/wast/test_wat_numeric_edge_cases.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i64.const 0x8000000000000000 i64.const 1 i64.rotl call $log_i64)
)
