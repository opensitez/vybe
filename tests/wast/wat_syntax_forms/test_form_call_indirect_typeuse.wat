;; vybe-test: wast/wat_syntax_forms/test_form_call_indirect_typeuse
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $ii (func (param i32) (result i32)))
(func $double (param i32) (result i32) local.get 0 i32.const 2 i32.mul)
(table 1 funcref)
(elem (i32.const 0) $double)
(func (export "_start")
  i32.const 21 i32.const 0 call_indirect (type $ii) call $log)
)
