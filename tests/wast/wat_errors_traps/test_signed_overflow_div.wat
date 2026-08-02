;; vybe-test: wast/wat_errors_traps/test_signed_overflow_div
;; origin: languages/wast/tests/wast/test_wat_errors_traps.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const -2147483648 i32.const -1 i32.div_s call $log)
)
