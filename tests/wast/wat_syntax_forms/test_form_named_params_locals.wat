;; vybe-test: wast/wat_syntax_forms/test_form_named_params_locals
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

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
  (func $sq (param $x i32) (result i32) (local $y i32)
  local.get $x local.get $x i32.mul local.set $y
  local.get $y)
(func (export "_start") i32.const 6 call $sq i32.const 36 call $vybe_check_i32)
)
