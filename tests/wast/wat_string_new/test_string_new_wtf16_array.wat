;; vybe-test: wast/wat_string_new/test_string_new_wtf16_array
;; origin: languages/wast/tests/wast/test_wat_string_new.rs

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
  (type $A (array (mut i16)))
(func (export "_start") (local $a (ref null $A))
  i32.const 104 ;; 'h'
  i32.const 105 ;; 'i'
  array.new_fixed $A 2
  local.set $a
  
  local.get $a
  i32.const 0
  i32.const 2
  string.new_wtf16_array
  string.measure_utf8
  i32.const 2 call $vybe_check_i32
)
)
