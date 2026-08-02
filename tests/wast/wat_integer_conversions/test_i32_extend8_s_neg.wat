;; vybe-test: wast/wat_integer_conversions/test_i32_extend8_s_neg
;; origin: languages/wast/tests/wast/test_wat_integer_conversions.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  i32.const 255 ;; -1 as i8
  i32.extend8_s
  call $log
)
)
