;; vybe-test: wast/wat_table_ops/test_table_copy
;; origin: languages/wast/tests/wast/test_wat_table_ops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (table $t1 5 funcref)
(table $t2 5 funcref)
(func $f1)
(elem (table $t1) (i32.const 2) $f1)
(func (export "_start")
  i32.const 0
  i32.const 2
  i32.const 1
  table.copy $t2 $t1
  i32.const 0
  table.get $t2
  ref.is_null
  call $log
)
)
