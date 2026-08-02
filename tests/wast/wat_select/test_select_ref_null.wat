;; vybe-test: wast/wat_select/test_select_ref_null
;; origin: languages/wast/tests/wast/test_wat_select.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $f1)
(func (export "_start")
  ref.func $f1
  ref.null func
  i32.const 0
  select (result funcref)
  ref.is_null
  call $log
)
)
