;; vybe-test: wast/wat_memory_grow/test_memory_grow_multiple
;; origin: languages/wast/tests/wast/test_wat_memory_grow.rs

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
  i32.const 2
  memory.grow
  drop
  i32.const 3
  memory.grow
  i32.const 3 call $vybe_check_i32)
)
