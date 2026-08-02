;; vybe-test: wast/wat_if_else/test_if_else_true
;; origin: languages/wast/tests/wast/test_wat_if_else.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 1
  if
    i32.const 42
    call $log
  else
    i32.const 99
    call $log
  end
)
)
