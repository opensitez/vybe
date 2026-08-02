;; vybe-test: wast/wat_multivalue/test_block_param_consumed
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const 3 block (param i32) (result i32) i32.const 4 i32.add end call $log)
)
