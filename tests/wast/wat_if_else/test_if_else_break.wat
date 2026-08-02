;; vybe-test: wast/wat_if_else/test_if_else_break
;; origin: languages/wast/tests/wast/test_wat_if_else.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  block (result i32)
    i32.const 0
    if (result i32)
      i32.const 42
    else
      i32.const 99
      br 1
    end
  end
  call $log
)
)
