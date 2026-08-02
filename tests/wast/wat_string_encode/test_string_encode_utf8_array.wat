;; vybe-test: wast/wat_string_encode/test_string_encode_utf8_array
;; origin: languages/wast/tests/wast/test_wat_string_encode.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(type $A (array (mut i8)))
(data (i32.const 0) "hello")
(func (export "_start") (local $a (ref null $A))
  i32.const 5
  array.new_default $A
  local.set $a
  
  i32.const 0
  i32.const 5
  string.new_utf8
  
  local.get $a
  i32.const 0
  string.encode_utf8_array
  drop
  
  local.get $a
  i32.const 1
  array.get_u $A
  call $log
)
)
