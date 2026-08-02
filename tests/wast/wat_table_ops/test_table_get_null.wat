;; vybe-test: wast/wat_table_ops/test_table_get_null
;; origin: languages/wast/tests/wast/test_wat_table_ops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (table 2 funcref)
(func $f1 (result i32) i32.const 42)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 1
  table.get 0
  ref.is_null
  call $log
)
)
