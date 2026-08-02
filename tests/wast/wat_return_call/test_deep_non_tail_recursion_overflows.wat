;; vybe-test: wast/wat_return_call/test_deep_non_tail_recursion_overflows
;; origin: languages/wast/tests/wast/test_wat_return_call.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $sum (param $n i32) (result i32)
  local.get $n
  i32.eqz
  if (result i32)
    i32.const 0
  else
    local.get $n
    local.get $n
    i32.const 1
    i32.sub
    call $sum
    i32.add
  end)
(func (export "_start") (result i32)
  i32.const 20000
  call $sum
)
)
