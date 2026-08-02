;; vybe-test: wast/wat_memory_load/test_memory_load8_u
;; origin: languages/wast/tests/wast/test_wat_memory_load.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\80")
(func (export "_start") i32.const 0 i32.load8_u call $log)
)
