;; vybe-test: wast/wat_unreachable/test_unreachable_polymorphic_stack
;; origin: languages/wast/tests/wast/test_wat_unreachable.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  unreachable
  i32.add ;; consumes values that don't exist, valid due to polymorphic stack
  call $log
)
)
