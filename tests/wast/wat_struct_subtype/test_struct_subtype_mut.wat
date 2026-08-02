;; vybe-test: wast/wat_struct_subtype/test_struct_subtype_mut
;; origin: languages/wast/tests/wast/test_wat_struct_subtype.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Base (struct (field (mut i32))))
(type $Sub (struct_subtype (field (mut i32)) (field (mut i32)) $Base))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  struct.new $Sub
  local.set $s
  
  local.get $s
  i32.const 42
  struct.set $Base 0
  
  local.get $s
  struct.get $Base 0
  call $log
)
)
