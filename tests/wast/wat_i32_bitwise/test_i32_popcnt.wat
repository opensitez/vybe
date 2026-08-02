;; vybe-test: wast/wat_i32_bitwise/test_i32_popcnt
;; origin: languages/wast/tests/wast/test_wat_i32_bitwise.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 0x0F0F0F0F
  i32.popcnt
  call $log
)
)
