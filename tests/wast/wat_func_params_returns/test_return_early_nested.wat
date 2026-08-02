;; vybe-test: wast/wat_func_params_returns/test_return_early_nested
;; origin: languages/wast/tests/wast/test_wat_func_params_returns.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $f1 (result i32)
  block
    block
      i32.const 42
      return
    end
  end
  i32.const 99)
(func (export "_start")
  call $f1
  call $log
)
)
