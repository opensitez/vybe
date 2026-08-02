;; vybe-test: wast/wat_memory_store/test_memory_store_align
;; origin: languages/wast/tests/wast/test_wat_memory_store.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(func (export "_start") 
  i32.const 1 
  i32.const 42
  i32.store align=1 
  i32.const 1 
  i32.load 
  call $log)
)
