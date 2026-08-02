;; vybe-test: wast/wat_syntax_forms/test_form_start_function
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $g (mut i32) (i32.const 0))
(func $init i32.const 99 global.set $g)
(start $init)
(func (export "_start") global.get $g call $log)
)
