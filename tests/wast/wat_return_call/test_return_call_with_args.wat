;; vybe-test: wast/wat_return_call/test_return_call_with_args
;; origin: languages/wast/tests/wast/test_wat_return_call.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(func (export "_start") (result i32)
  i32.const 10
  i32.const 20
  return_call $add
)
)
