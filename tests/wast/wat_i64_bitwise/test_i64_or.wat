;; vybe-test: wast/wat_i64_bitwise/test_i64_or
;; origin: languages/wast/tests/wast/test_wat_i64_bitwise.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i64.const 0x0F
  i64.const 0x30
  i64.or
  call $log_i64
)
)
