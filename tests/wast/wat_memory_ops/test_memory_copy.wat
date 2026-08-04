;; vybe-test: wast/wat_memory_ops/test_memory_copy
;; origin: languages/wast/tests/wast/test_wat_memory_ops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
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
  i32.const -1 call $vybe_check_i32
)
)
