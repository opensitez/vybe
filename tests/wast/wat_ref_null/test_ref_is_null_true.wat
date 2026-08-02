;; vybe-test: wast/wat_ref_null/test_ref_is_null_true
;; origin: languages/wast/tests/wast/test_wat_ref_null.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $t (struct (field i32)))
(func (export "_start")
  (local $s (ref null $t))
  (local.set $s (ref.null $t))
  (ref.is_null (local.get $s))
  call $log
)
)
