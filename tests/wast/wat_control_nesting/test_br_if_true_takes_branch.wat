;; vybe-test: wast/wat_control_nesting/test_br_if_true_takes_branch
;; origin: languages/wast/tests/wast/test_wat_control_nesting.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        block i32.const 7 call $log i32.const 1 br_if 0 i32.const 8 call $log end)
)
