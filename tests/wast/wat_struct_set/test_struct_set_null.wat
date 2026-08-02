;; vybe-test: wast/wat_struct_set/test_struct_set_null
;; origin: languages/wast/tests/wast/test_wat_struct_set.rs
;; vybe-test-mode: run-fail

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Point (struct (field (mut i32)) (field (mut i32))))
(func (export "_start") (local $p (ref null $Point))
  ref.null $Point
  local.set $p
  
  local.get $p
  i32.const 42
  struct.set $Point 0
  
  i32.const 0
  call $log
)
)
