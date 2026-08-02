;; vybe-test: wast/wat_elem_segments/test_elem_segment_declarative
;; origin: languages/wast/tests/wast/test_wat_elem_segments.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $f1 (result i32) i32.const 42)
(elem declare $f1)
(func (export "_start")
  ref.func $f1
  drop
  i32.const 42
  call $log
)
)
