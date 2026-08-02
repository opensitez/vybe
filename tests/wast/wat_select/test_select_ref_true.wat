;; vybe-test: wast/wat_select/test_select_ref_true
;; origin: languages/wast/tests/wast/test_wat_select.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $f1)
(func $f2)
(func (export "_start")
  ref.func $f1
  ref.func $f2
  i32.const 1
  select (result funcref)
  ref.is_null
  call $log
)
)
