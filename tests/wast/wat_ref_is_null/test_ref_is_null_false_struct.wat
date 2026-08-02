;; vybe-test: wast/wat_ref_is_null/test_ref_is_null_false_struct
;; origin: languages/wast/tests/wast/test_wat_ref_is_null.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i32)))
(func (export "_start")
  i32.const 42
  struct.new $S
  ref.is_null
  call $log
)
)
