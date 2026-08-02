;; vybe-test: wast/wat_globals_mut/test_global_mut_init_expr
;; origin: languages/wast/tests/wast/test_wat_globals_mut.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $a i32 (i32.const 10))
(global $b (mut i32) (global.get $a))
(func (export "_start")
  global.get $b
  i32.const 20
  i32.add
  global.set $b
  global.get $b
  call $log
)
)
