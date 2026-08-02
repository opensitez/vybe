;; vybe-test: wast/wat_loop/test_loop_no_yield_from_continue
;; origin: languages/wast/tests/wast/test_wat_loop.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 0
  block (result i32)
    loop (param i32) (result i32)
      drop
      i32.const 10
      br 1
      i32.const 99
      br 0
    end
  end
  call $log)
)
