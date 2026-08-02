;; vybe-test: wast/wat_globals_const/test_global_const_multiple
;; origin: languages/wast/tests/wast/test_wat_globals_const.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $a i32 (i32.const 10))
(global $b i32 (i32.const 20))
(func (export "_start")
  global.get $a
  global.get $b
  i32.add
  call $log
)
)
