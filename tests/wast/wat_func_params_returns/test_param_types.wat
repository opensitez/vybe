;; vybe-test: wast/wat_func_params_returns/test_param_types
;; origin: languages/wast/tests/wast/test_wat_func_params_returns.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_f64 (param f64) (param f64)
    local.get 0
    local.get 1
    f64.ne
    if
      unreachable
    end)
  (func $f1 (param $i i32) (param $f f32) (param $l i64) (param $d f64) (result f64)
  local.get $d)
(func (export "_start")
  i32.const 10
  f32.const 1.0
  i64.const 20
  f64.const 42.5
  call $f1
  f64.const 42.5 call $vybe_check_f64
)
)
