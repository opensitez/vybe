;; vybe-test: wast/wat_syntax_forms/test_form_table_inline_elem
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
  (type $v (func (result i32)))
(func $f (result i32) i32.const 55)
(table funcref (elem $f))
(func (export "_start") i32.const 0 call_indirect (type $v) i32.const 55 call $vybe_check_i32)
)
