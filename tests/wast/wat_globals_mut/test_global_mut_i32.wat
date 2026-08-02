;; vybe-test: wast/wat_globals_mut/test_global_mut_i32
;; origin: languages/wast/tests/wast/test_wat_globals_mut.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $g (mut i32) (i32.const 42))
(func (export "_start")
  global.get $g
  i32.const 100
  i32.add
  global.set $g
  global.get $g
  call $log
)
)
