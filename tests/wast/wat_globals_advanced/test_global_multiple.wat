;; vybe-test: wast/wat_globals_advanced/test_global_multiple
;; origin: languages/wast/tests/wast/test_wat_globals_advanced.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (global $g1 (mut i32) (i32.const 10))
(global $g2 (mut i32) (i32.const 20))
(func (export "_start")
  global.get $g1
  global.get $g2
  i32.add
  i32.const 30 call $vybe_check_i32
)
)
