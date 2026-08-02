;; vybe-test: wast/wat_globals_advanced/test_global_ref_null_traps
;; origin: languages/wast/tests/wast/test_wat_globals_advanced.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i32)))
(global $g (mut (ref null $S)) (ref.null $S))
(func (export "_start")
  global.get $g
  struct.get $S 0
  call $log
)
)
