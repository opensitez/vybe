;; vybe-test: wast/wat_elem_segments/test_elem_segment_active_overlap
;; origin: languages/wast/tests/wast/test_wat_elem_segments.rs

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
  (table 5 funcref)
(func $f1)
(func $f2)
(elem (i32.const 0) $f1)
(elem (i32.const 0) $f2)
(func (export "_start")
  i32.const 0
  table.get 0
  ref.func $f2
  ref.eq
  i32.const 1 call $vybe_check_i32
)
)
