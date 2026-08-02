;; vybe-test: wast/wat_integer_conversions/test_i32_wrap_i64
;; origin: languages/wast/tests/wast/test_wat_integer_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i64.const 4294967338 ;; 0x10000002A -> wraps to 42
  i32.wrap_i64
  call $log
)
)
