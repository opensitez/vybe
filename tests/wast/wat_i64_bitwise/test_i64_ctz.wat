;; vybe-test: wast/wat_i64_bitwise/test_i64_ctz
;; origin: languages/wast/tests/wast/test_wat_i64_bitwise.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i64.const 0x8000000000000000
  i64.ctz
  call $log_i64
)
)
