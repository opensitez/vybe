;; vybe-test: wast/wat_control_nesting/test_return_exits_function_early
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const 5 call $log return i32.const 6 call $log)
)
