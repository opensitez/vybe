;; vybe-test: wast/wat_array_new/test_array_new_data
;; origin: languages/wast/tests/wast/test_wat_array_new.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (type $Arr (array (mut i8)))
(data $d "data")
(func (export "_start") (local $a (ref null $Arr))
  i32.const 0
  i32.const 4
  array.new_data $Arr $d
  local.set $a
  
  local.get $a
  i32.const 0
  array.get_u $Arr
  call $log
)
)
