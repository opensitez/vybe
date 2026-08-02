;; vybe-test: wast/wat_struct_subtype/test_struct_subtype_deep
;; origin: languages/wast/tests/wast/test_wat_struct_subtype.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Base (struct (field i32)))
(type $Sub1 (struct_subtype (field i32) (field i32) $Base))
(type $Sub2 (struct_subtype (field i32) (field i32) (field i32) $Sub1))
(func (export "_start") (local $s (ref null $Base))
  i32.const 10
  i32.const 20
  i32.const 30
  struct.new $Sub2
  local.set $s
  
  local.get $s
  struct.get $Base 0
  call $log
)
)
