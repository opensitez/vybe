;; vybe-test: wast/wat_table_ops/test_table_grow_fail
;; origin: languages/wast/tests/wast/test_wat_table_ops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (table 5 5 funcref)
(func (export "_start")
  ref.null func
  i32.const 1
  table.grow 0
  call $log
)
)
