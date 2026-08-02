;; vybe-test: wast/wat_multivalue/test_if_multi_value_result
;; origin: languages/wast/tests/wast/test_wat_multivalue.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        i32.const 1 if (result i32) i32.const 100 else i32.const 200 end call $log)
)
