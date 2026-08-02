;; vybe-test: wast/wat_multivalue/test_block_result_value
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        block (result i32) i32.const 5 i32.const 6 i32.add end call $log)
)
