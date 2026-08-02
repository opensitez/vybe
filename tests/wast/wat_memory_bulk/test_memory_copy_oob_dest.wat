;; vybe-test: wast/wat_memory_bulk/test_memory_copy_oob_dest
;; origin: languages/wast/tests/wast/test_wat_memory_bulk.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(func (export "_start")
  i32.const 65530
  i32.const 0
  i32.const 10
  memory.copy
  i32.const 0
  call $log)
)
