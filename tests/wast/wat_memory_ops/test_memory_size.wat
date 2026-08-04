;; vybe-test: wast/wat_memory_ops/test_memory_size
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
  (memory 2)
(func (export "_start")
  memory.size
  i32.const 2 call $vybe_check_i32
)
)
