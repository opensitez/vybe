;; vybe-test: wast/wat_globals_advanced/test_global_different_types
;; origin: languages/wast/tests/wast/test_wat_globals_advanced.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $gi i32 (i32.const 42))
(global $gf f32 (f32.const 3.14))
(global $gl i64 (i64.const 99))
(func (export "_start")
  global.get $gl
  call $log_i64
)
)
