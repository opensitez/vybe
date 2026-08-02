;; vybe-test: wast/wat_ref_null/test_br_on_non_null_success
;; origin: languages/wast/tests/wast/test_wat_ref_null.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $t (struct (field i32)))
(func (export "_start")
  (local $s (ref null $t))
  (local.set $s (struct.new $t (i32.const 42)))
  (block $L (result (ref $t))
    (br_on_non_null $L (local.get $s))
    (return (i32.const 0))
  )
  (struct.get $t 0)
  call $log
)
)
