;; vybe-test: wast/wat_struct_get/test_struct_get_default_local_null
;; origin: languages/wast/tests/wast/test_wat_struct_get.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Point (struct (field i32) (field i32)))
(func (export "_start") (local $p (ref null $Point))
  local.get $p
  struct.get $Point 0
  call $log
)
)
