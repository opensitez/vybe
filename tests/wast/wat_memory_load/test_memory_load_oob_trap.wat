;; vybe-test: wast/wat_memory_load/test_memory_load_oob_trap
;; origin: languages/wast/tests/wast/test_wat_memory_load.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(func (export "_start") i32.const 65536 i32.load call $log)
)
