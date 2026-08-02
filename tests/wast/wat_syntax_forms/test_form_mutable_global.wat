;; vybe-test: wast/wat_syntax_forms/test_form_mutable_global
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $g (mut i32) (i32.const 5))
(func (export "_start")
  i32.const 10 global.set $g
  global.get $g call $log)
)
