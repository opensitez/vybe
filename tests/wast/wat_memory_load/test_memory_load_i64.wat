;; vybe-test: wast/wat_memory_load/test_memory_load_i64
;; origin: languages/wast/tests/wast/test_wat_memory_load.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 8) "\01\02\03\04\05\06\07\08")
(func (export "_start") i32.const 8 i64.load call $log_i64)
)
