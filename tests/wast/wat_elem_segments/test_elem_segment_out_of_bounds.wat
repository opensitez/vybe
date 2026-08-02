;; vybe-test: wast/wat_elem_segments/test_elem_segment_out_of_bounds
;; origin: languages/wast/tests/wast/test_wat_elem_segments.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (table 2 funcref)
(func $f1)
(elem (i32.const 5) $f1)
(func (export "_start")
  i32.const 42
  call $log
)
)
