;; vybe-test: wast/wat_func_locals/test_local_array_null
;; origin: languages/wast/tests/wast/test_wat_func_locals.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $A (array i32))
(func (export "_start") (local $a (ref null $A))
  local.get $a
  ref.is_null
  call $log
)
)
