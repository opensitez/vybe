;; vybe-test: wast/wat_ref_null/test_br_on_null_success
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
  (block $L
    (drop (br_on_null $L (local.get $s)))
    (i32.const 0)
    call $log
    return
  )
  (i32.const 1)
  call $log
)
)
