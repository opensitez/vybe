;; vybe-test: wast/wat_memory_grow/test_memory_size_after_failed_grow
;; origin: languages/wast/tests/wast/test_wat_memory_grow.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1 2)
(func (export "_start") 
  i32.const 5
  memory.grow
  drop
  memory.size
  call $log)
)
