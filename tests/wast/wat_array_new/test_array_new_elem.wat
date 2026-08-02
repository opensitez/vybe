;; vybe-test: wast/wat_array_new/test_array_new_elem
;; origin: languages/wast/tests/wast/test_wat_array_new.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Arr (array (mut funcref)))
(func $f)
(elem $e $f)
(func (export "_start") (local $a (ref null $Arr))
  i32.const 0
  i32.const 1
  array.new_elem $Arr $e
  local.set $a

  local.get $a
  array.len
  call $log
)
)
