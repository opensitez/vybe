;; vybe-test: wast/wat_data_segments/test_data_segment_multi_memory
;; origin: languages/wast/tests/wast/test_wat_data_segments.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory $m1 1)
(memory $m2 1)
(data (memory $m2) (i32.const 0) "hello")
(func (export "_start")
  i32.const 0
  i32.load8_u $m2
  call $log
)
)
