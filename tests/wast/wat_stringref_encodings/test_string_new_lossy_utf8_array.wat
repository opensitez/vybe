;; vybe-test: wast/wat_stringref_encodings/test_string_new_lossy_utf8_array
;; origin: languages/wast/tests/wast/test_wat_stringref_encodings.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (type $A (array (mut i8)))
(func (export "_start") (local $a (ref null $A))
  i32.const 65 i32.const 66 i32.const 67 array.new_fixed $A 3
  local.set $a
  local.get $a i32.const 0 i32.const 3 string.new_lossy_utf8_array
  string.measure_utf8
  i32.const 3 call $vybe_check_i32)
)
