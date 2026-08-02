;; vybe-test: wast/wat_integer_conversions/test_i64_extend16_s_pos
;; origin: languages/wast/tests/wast/test_wat_integer_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i64.const 32767
  i64.extend16_s
  call $log_i64
)
)
