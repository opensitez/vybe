;; vybe-test: wast/wat_array_new/test_array_new_fixed
;; origin: languages/wast/tests/wast/test_wat_array_new.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Arr (array i32))
(func (export "_start") (local $a (ref null $Arr))
  i32.const 10
  i32.const 20
  i32.const 30
  array.new_fixed $Arr 3
  local.set $a
  
  local.get $a
  i32.const 1
  array.get $Arr
  call $log
)
)
