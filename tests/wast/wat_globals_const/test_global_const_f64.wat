;; vybe-test: wast/wat_globals_const/test_global_const_f64
;; origin: languages/wast/tests/wast/test_wat_globals_const.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $g f64 (f64.const 3.14))
(func (export "_start")
  global.get $g
  i32.trunc_f64_s
  call $log
)
)
