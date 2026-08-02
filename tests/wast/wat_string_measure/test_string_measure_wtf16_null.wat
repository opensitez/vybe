;; vybe-test: wast/wat_string_measure/test_string_measure_wtf16_null
;; origin: languages/wast/tests/wast/test_wat_string_measure.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  ref.null string
  string.measure_wtf16
  call $log
)
)
