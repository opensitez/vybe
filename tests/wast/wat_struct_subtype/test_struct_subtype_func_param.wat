;; vybe-test: wast/wat_struct_subtype/test_struct_subtype_func_param
;; origin: languages/wast/tests/wast/test_wat_struct_subtype.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Base (struct (field i32)))
(type $Sub (struct_subtype (field i32) (field i32) $Base))
(func $f1 (param $b (ref null $Base)) (result i32)
  local.get $b
  struct.get $Base 0)
(func (export "_start")
  i32.const 99
  i32.const 88
  struct.new $Sub
  call $f1
  call $log
)
)
