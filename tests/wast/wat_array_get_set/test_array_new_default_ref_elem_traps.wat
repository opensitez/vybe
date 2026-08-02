;; vybe-test: wast/wat_array_get_set/test_array_new_default_ref_elem_traps
;; origin: languages/wast/tests/wast/test_wat_array_get_set.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Inner (struct (field i32)))
(type $A (array (ref null $Inner)))
(func (export "_start")
  i32.const 2
  array.new_default $A
  i32.const 0
  array.get $A
  struct.get $Inner 0
  call $log
)
)
