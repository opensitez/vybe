;; vybe-test: wast/wat_return_call/test_return_call_direct
;; origin: languages/wast/tests/wast/test_wat_return_call.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $f1 (result i32)
  i32.const 42)
(func (export "_start") (result i32)
  return_call $f1
)
)
