;; vybe-test: wast/wat_func_params_returns/test_param_index_access
;; origin: languages/wast/tests/wast/test_wat_func_params_returns.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $f1 (param i32) (param i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(func (export "_start")
  i32.const 10
  i32.const 20
  call $f1
  call $log
)
)
