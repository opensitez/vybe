;; vybe-test: wast/wat_func_params_returns/test_param_ref_non_null
;; origin: languages/wast/tests/wast/test_wat_func_params_returns.rs

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
  (type $S (struct (field i32)))
(func $f1 (param $s (ref null $S)) (result i32)
  local.get $s
  ref.is_null)
(func (export "_start")
  i32.const 42
  struct.new $S
  call $f1
  i32.const 0 call $vybe_check_i32
)
)
