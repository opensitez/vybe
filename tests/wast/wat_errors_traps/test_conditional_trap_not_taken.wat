;; vybe-test: wast/wat_errors_traps/test_conditional_trap_not_taken
;; origin: languages/wast/tests/wast/test_wat_errors_traps.rs

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
  (func (export "_start")
        i32.const 0 if unreachable end i32.const 99 i32.const 99 call $vybe_check_i32)
)
