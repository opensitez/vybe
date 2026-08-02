;; vybe-test: wast/wat_call_direct/test_call_recursive
;; origin: languages/wast/tests/wast/test_wat_call_direct.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $fact (param $n i32) (result i32)
  (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
    (then (i32.const 1))
    (else
      (i32.mul
        (local.get $n)
        (call $fact (i32.sub (local.get $n) (i32.const 1)))
      )
    )
  )
)
(func (export "_start")
  i32.const 5
  call $fact
  call $log
)
)
