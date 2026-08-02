;; vybe-test: wast/wat_elem_segments/test_elem_drop
;; origin: languages/wast/tests/wast/test_wat_elem_segments.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (table 5 funcref)
(func $f1 (result i32) i32.const 42)
(elem $e $f1)
(func (export "_start")
  elem.drop $e
  i32.const 0
  i32.const 0
  i32.const 1
  table.init $e
  i32.const 0
  call $log
)
)
