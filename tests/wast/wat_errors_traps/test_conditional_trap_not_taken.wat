;; vybe-test: wast/wat_errors_traps/test_conditional_trap_not_taken
;; origin: languages/wast/tests/wast/test_wat_errors_traps.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const 0 if unreachable end i32.const 99 call $log)
)
