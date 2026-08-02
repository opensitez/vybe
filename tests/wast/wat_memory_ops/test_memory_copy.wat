;; vybe-test: wast/wat_memory_ops/test_memory_copy
;; origin: languages/wast/tests/wast/test_wat_memory_ops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(func (export "_start")
  (i32.const 0) ;; dest
  (i32.const 255) ;; val
  (i32.const 4) ;; len
  memory.fill
  
  (i32.const 10) ;; dest
  (i32.const 0) ;; src
  (i32.const 4) ;; len
  memory.copy

  (i32.const 10)
  i32.load
  call $log
)
)
