;; vybe-test: wast/wat_unreachable/test_unreachable_in_if_true
;; origin: languages/wast/tests/wast/test_wat_unreachable.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 1
  if
    unreachable
  else
    i32.const 42
    call $log
  end
)
)
