;; vybe-test: wast/wat_block/test_block_params_break
;; origin: languages/wast/tests/wast/test_wat_block.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 5
  block (param i32) (result i32)
    i32.const 10
    i32.add
    br 0
  end
  call $log)
)
