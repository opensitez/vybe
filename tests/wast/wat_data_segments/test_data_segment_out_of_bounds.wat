;; vybe-test: wast/wat_data_segments/test_data_segment_out_of_bounds
;; origin: languages/wast/tests/wast/test_wat_data_segments.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 65535) "hello")
(func (export "_start")
  i32.const 42
  call $log
)
)
