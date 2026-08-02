;; vybe-test: wast/wat_syntax_forms/test_form_table_inline_elem
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $v (func (result i32)))
(func $f (result i32) i32.const 55)
(table funcref (elem $f))
(func (export "_start") i32.const 0 call_indirect (type $v) call $log)
)
