;; vybe-test: wast/wat_memory_store/test_memory_store8_i64
;; origin: languages/wast/tests/wast/test_wat_memory_store.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(func (export "_start") 
  i32.const 0 
  i64.const 300 
  i64.store8 
  i32.const 0 
  i64.load8_u 
  call $log_i64)
)
