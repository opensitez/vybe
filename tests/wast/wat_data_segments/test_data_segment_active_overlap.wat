;; vybe-test: wast/wat_data_segments/test_data_segment_active_overlap
;; origin: languages/wast/tests/wast/test_wat_data_segments.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (memory 1)
(data (i32.const 0) "hello")
(data (i32.const 1) "world")
(func (export "_start")
  i32.const 1
  i32.load8_u
  i32.const 119 call $vybe_check_i32
)
)
