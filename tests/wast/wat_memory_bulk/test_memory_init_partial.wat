;; vybe-test: wast/wat_memory_bulk/test_memory_init_partial
;; origin: languages/wast/tests/wast/test_wat_memory_bulk.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data $d "data")
(func (export "_start")
  i32.const 10
  i32.const 1
  i32.const 2
  memory.init $d
  i32.const 10
  i32.load8_u
  call $log)
)
