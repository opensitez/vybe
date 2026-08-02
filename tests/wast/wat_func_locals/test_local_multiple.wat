;; vybe-test: wast/wat_func_locals/test_local_multiple
;; origin: languages/wast/tests/wast/test_wat_func_locals.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") (local $x i32) (local $y i32)
  i32.const 10
  local.set $x
  i32.const 20
  local.set $y
  local.get $x
  local.get $y
  i32.add
  call $log
)
)
