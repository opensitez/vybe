;; vybe-test: wast/wat_struct_new/test_struct_new_default_ref_field_traps
;; origin: languages/wast/tests/wast/test_wat_struct_new.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Inner (struct (field i32)))
(type $Outer (struct (field (ref null $Inner))))
(func (export "_start")
  struct.new_default $Outer
  struct.get $Outer 0
  struct.get $Inner 0
  call $log
)
)
