;; vybe-test: wast/wat_stringref_encodings/test_string_new_wtf8_array
;; origin: languages/wast/tests/wast/test_wat_stringref_encodings.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $A (array (mut i8)))
(func (export "_start") (local $a (ref null $A))
  i32.const 104 i32.const 105 array.new_fixed $A 2
  local.set $a
  local.get $a i32.const 0 i32.const 2 string.new_wtf8_array
  string.measure_utf8
  call $log)
)
