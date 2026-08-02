;; vybe-test: wast/wat_call_direct/test_call_simple
;; origin: languages/wast/tests/wast/test_wat_call_direct.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $helper (result i32)
  i32.const 42
)
(func (export "_start")
  call $helper
  call $log
)
)
