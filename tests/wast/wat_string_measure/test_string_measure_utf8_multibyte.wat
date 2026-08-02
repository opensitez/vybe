;; vybe-test: wast/wat_string_measure/test_string_measure_utf8_multibyte
;; origin: languages/wast/tests/wast/test_wat_string_measure.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\e2\82\ac") ;; euro sign, 3 bytes
(func (export "_start")
  i32.const 0
  i32.const 3
  string.new_utf8
  string.measure_utf8
  call $log
)
)
