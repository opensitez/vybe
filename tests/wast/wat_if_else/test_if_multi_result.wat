;; vybe-test: wast/wat_if_else/test_if_multi_result
;; origin: languages/wast/tests/wast/test_wat_if_else.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 1
  if (result i32 i32)
    i32.const 10
    i32.const 20
  else
    i32.const 30
    i32.const 40
  end
  i32.add
  call $log
)
)
