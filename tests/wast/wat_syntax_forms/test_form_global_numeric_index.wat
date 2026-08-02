;; vybe-test: wast/wat_syntax_forms/test_form_global_numeric_index
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global (mut i32) (i32.const 3))
(func (export "_start")
  i32.const 8 global.set 0
  global.get 0 call $log)
)
