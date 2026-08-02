;; vybe-test: wast/wat_globals_advanced/test_global_init_expr
;; origin: languages/wast/tests/wast/test_wat_globals_advanced.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (global $g1 i32 (i32.const 10))
(global $g2 i32 (global.get $g1))
(func (export "_start")
  global.get $g2
  call $log
)
)
