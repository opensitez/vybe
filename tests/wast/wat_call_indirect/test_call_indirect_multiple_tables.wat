;; vybe-test: wast/wat_call_indirect/test_call_indirect_multiple_tables
;; origin: languages/wast/tests/wast/test_wat_call_indirect.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $sig (func (result i32)))
(table $t1 1 funcref)
(table $t2 1 funcref)
(func $f1 (result i32) i32.const 42)
(func $f2 (result i32) i32.const 99)
(elem (table $t1) (i32.const 0) $f1)
(elem (table $t2) (i32.const 0) $f2)
(func (export "_start")
  i32.const 0
  call_indirect $t2 (type $sig)
  call $log
)
)
