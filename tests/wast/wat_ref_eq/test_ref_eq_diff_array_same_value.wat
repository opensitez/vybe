;; vybe-test: wast/wat_ref_eq/test_ref_eq_diff_array_same_value
;; origin: languages/wast/tests/wast/test_wat_ref_eq.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $A (array i32))
(func (export "_start")
  i32.const 42
  i32.const 5
  array.new $A
  i32.const 42
  i32.const 5
  array.new $A
  ref.eq
  call $log
)
)
