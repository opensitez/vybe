;; vybe-test: wast/wat_memory_load/test_memory_load_offset
;; origin: languages/wast/tests/wast/test_wat_memory_load.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 10) "\10\20\30\40")
(func (export "_start") i32.const 5 i32.load offset=5 call $log)
)
