;; vybe-test: wast/wat_br_table/test_br_table_empty
;; origin: languages/wast/tests/wast/test_wat_br_table.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  block (result i32)
    i32.const 10
    i32.const 0
    br_table 0
  end
  call $log
)
)
