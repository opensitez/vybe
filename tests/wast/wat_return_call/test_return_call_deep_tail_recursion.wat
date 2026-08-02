;; vybe-test: wast/wat_return_call/test_return_call_deep_tail_recursion
;; origin: languages/wast/tests/wast/test_wat_return_call.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $sum (param $n i32) (param $acc i32) (result i32)
  local.get $n
  i32.eqz
  if (result i32)
    local.get $acc
  else
    local.get $n
    i32.const 1
    i32.sub
    local.get $acc
    local.get $n
    i32.add
    return_call $sum
  end)
(func (export "_start") (result i32)
  i32.const 300
  i32.const 0
  return_call $sum
)
)
