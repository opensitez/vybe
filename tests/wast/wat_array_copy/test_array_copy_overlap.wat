;; vybe-test: wast/wat_array_copy/test_array_copy_overlap
;; origin: languages/wast/tests/wast/test_wat_array_copy.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Arr (array (mut i32)))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 10
  i32.const 20
  i32.const 30
  array.new_fixed $Arr 3
  local.set $a
  
  local.get $a
  i32.const 1
  local.get $a
  i32.const 0
  i32.const 2
  array.copy $Arr $Arr
  
  local.get $a
  i32.const 1
  array.get $Arr
  call $log
)
)
