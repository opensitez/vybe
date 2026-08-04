;; vybe-test: wast/wat_if_else/test_if_param_result
;; origin: languages/wast/tests/wast/test_wat_if_else.rs

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
  (func (export "_start")
  i32.const 10
  i32.const 1
  if (param i32) (result i32)
    i32.const 5
    i32.add
  else
    i32.const 2
    i32.mul
  end
  i32.const 15 call $vybe_check_i32
)
)
