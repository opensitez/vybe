;; vybe-test: wast/wat_func_locals/test_local_ref_null
;; origin: languages/wast/tests/wast/test_wat_func_locals.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") (local $r funcref)
  local.get $r
  ref.is_null
  call $log
)
)
