;; vybe-test: wast/wat_syntax_forms/test_form_folded_if_result
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  (if (result i32) (i32.const 1) (then (i32.const 1)) (else (i32.const 0)))
  call $log)
)
