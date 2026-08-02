;; vybe-test: wast/wat_table_ops/test_table_fill_oob
;; origin: languages/wast/tests/wast/test_wat_table_ops.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (table 5 funcref)
(func $f1)
(func (export "_start")
  i32.const 3
  ref.func $f1
  i32.const 3
  table.fill 0
  i32.const 1
  call $log
)
)
