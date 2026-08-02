;; vybe-test: wast/wat_ref_is_null/test_ref_is_null_branch
;; origin: languages/wast/tests/wast/test_wat_ref_is_null.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  ref.null func
  ref.is_null
  if
    i32.const 42
    call $log
  else
    i32.const 99
    call $log
  end
)
)
