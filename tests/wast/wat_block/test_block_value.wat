;; vybe-test: wast/wat_block/test_block_value
;; origin: languages/wast/tests/wast/test_wat_block.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  block (result i32)
    i32.const 10
  end
  call $log)
)
