;; vybe-test: wast/wat_func_params_returns/test_param_ref_null
;; origin: languages/wast/tests/wast/test_wat_func_params_returns.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i32)))
(func $f1 (param $s (ref null $S)) (result i32)
  local.get $s
  ref.is_null)
(func (export "_start")
  ref.null $S
  call $f1
  call $log
)
)
