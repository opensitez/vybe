;; vybe-test: wast/wat_globals_mut/test_global_mut_f64
;; origin: languages/wast/tests/wast/test_wat_globals_mut.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $g (mut f64) (f64.const 3.14))
(func (export "_start")
  global.get $g
  f64.const 10.0
  f64.add
  global.set $g
  global.get $g
  i32.trunc_f64_s
  call $log
)
)
