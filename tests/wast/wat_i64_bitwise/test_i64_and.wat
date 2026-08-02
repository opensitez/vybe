;; vybe-test: wast/wat_i64_bitwise/test_i64_and
;; origin: languages/wast/tests/wast/test_wat_i64_bitwise.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i64.const 0x0F
  i64.const 0x33
  i64.and
  call $log_i64
)
)
