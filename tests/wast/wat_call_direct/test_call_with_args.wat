;; vybe-test: wast/wat_call_direct/test_call_with_args
;; origin: languages/wast/tests/wast/test_wat_call_direct.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add
)
(func (export "_start")
  i32.const 10
  i32.const 20
  call $add
  i32.const 30 call $vybe_check_i32
)
)
