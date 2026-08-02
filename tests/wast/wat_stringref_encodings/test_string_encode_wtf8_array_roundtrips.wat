;; vybe-test: wast/wat_stringref_encodings/test_string_encode_wtf8_array_roundtrips
;; origin: languages/wast/tests/wast/test_wat_stringref_encodings.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $A (array (mut i8)))
(memory 1)
(data (i32.const 0) "\48\69")
(func (export "_start") (local $a (ref null $A))
  ;; make a 2-element array, encode "Hi" into it, decode back, compare.
  i32.const 0 i32.const 0 array.new_fixed $A 2
  local.set $a
  i32.const 0 i32.const 2 string.new_utf8   ;; "Hi"
  local.get $a i32.const 0 string.encode_wtf8_array   ;; returns count
  drop
  local.get $a i32.const 0 i32.const 2 string.new_wtf8_array  ;; "Hi" again
  i32.const 0 i32.const 2 string.new_utf8
  string.eq
  call $log)
)
