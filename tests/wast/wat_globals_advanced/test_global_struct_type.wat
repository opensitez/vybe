;; vybe-test: wast/wat_globals_advanced/test_global_struct_type
;; origin: languages/wast/tests/wast/test_wat_globals_advanced.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $S (struct (field i32)))
(global $g (mut (ref null $S)) (ref.null $S))
(func (export "_start")
  i32.const 42
  struct.new $S
  global.set $g
  global.get $g
  struct.get $S 0
  call $log
)
)
