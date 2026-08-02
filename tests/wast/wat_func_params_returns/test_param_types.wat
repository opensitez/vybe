;; vybe-test: wast/wat_func_params_returns/test_param_types
;; origin: languages/wast/tests/wast/test_wat_func_params_returns.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $f1 (param $i i32) (param $f f32) (param $l i64) (param $d f64) (result f64)
  local.get $d)
(func (export "_start")
  i32.const 10
  f32.const 1.0
  i64.const 20
  f64.const 42.5
  call $f1
  call $log_f64
)
)
