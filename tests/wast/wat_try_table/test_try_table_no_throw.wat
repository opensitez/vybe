;; vybe-test: wast/wat_try_table/test_try_table_no_throw
;; origin: languages/wast/tests/wast/test_wat_try_table.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  (try_table (result i32)
    i32.const 99)
  call $log)
)
