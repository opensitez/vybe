;; vybe-test: wast/wat_globals_advanced/test_global_ref_null_is_null
;; origin: languages/wast/tests/wast/test_wat_globals_advanced.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i32)))
(global $g (mut (ref null $S)) (ref.null $S))
(func (export "_start")
  (ref.is_null (global.get $g))
  call $log
)
)
