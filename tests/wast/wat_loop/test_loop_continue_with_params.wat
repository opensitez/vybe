;; vybe-test: wast/wat_loop/test_loop_continue_with_params
;; origin: languages/wast/tests/wast/test_wat_loop.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 0
  block (param i32) (result i32)
    loop (param i32) (result i32)
      local.set 0
      local.get 0
      local.get 0
      i32.const 5
      i32.eq
      br_if 1
      i32.const 1
      i32.add
      br 0
    end
  end
  call $log)
)
