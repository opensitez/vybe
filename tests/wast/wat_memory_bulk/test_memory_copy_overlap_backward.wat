;; vybe-test: wast/wat_memory_bulk/test_memory_copy_overlap_backward
;; origin: languages/wast/tests/wast/test_wat_memory_bulk.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 10) "abcdef")
(func (export "_start")
  i32.const 8 
  i32.const 10 
  i32.const 4
  memory.copy
  i32.const 8
  i32.load8_u
  call $log)
)
