;; vybe-test: wast/wat_syntax_forms/test_form_inline_export_memory
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
  (memory (export "mem") 1)
(func (export "_start")
  i32.const 0 i32.const 42 i32.store
  i32.const 0 i32.load i32.const 42 call $vybe_check_i32)
)
